// CompressCall.cpp

#include "StdAfx.h"

#include <wchar.h>

#include "../../../Common/IntToString.h"
#include "../../../Common/MyCom.h"
#include "../../../Common/Random.h"
#include "../../../Common/StringConvert.h"

#include "../../../Windows/DLL.h"
#include "../../../Windows/ErrorMsg.h"
#include "../../../Windows/FileDir.h"
#include "../../../Windows/FileMapping.h"
#include "../../../Windows/FileName.h"
#include "../../../Windows/MemoryLock.h"
#include "../../../Windows/ProcessUtils.h"
#include "../../../Windows/Synchronization.h"

#include "../FileManager/StringUtils.h"
#include "../FileManager/RegistryUtils.h"

#include "ZipRegistry.h"
#include "CompressCall.h"

using namespace NWindows;

#define MY_TRY_BEGIN try {

#define MY_TRY_FINISH } \
  catch(...) { ErrorMessageHRESULT(E_FAIL); return E_FAIL; }

#define MY_TRY_FINISH_VOID } \
  catch(...) { ErrorMessageHRESULT(E_FAIL); }

#define k7zGui  "7zG.exe"

// 21.07 : we can disable wildcard
// #define ISWITCH_NO_WILDCARD_POSTFIX "w-"
#define ISWITCH_NO_WILDCARD_POSTFIX

#define kShowDialogSwitch  " -ad"
#define kEmailSwitch  " -seml."
#define kArchiveTypeSwitch  " -t"
#define kIncludeSwitch  " -i" ISWITCH_NO_WILDCARD_POSTFIX
#define kArcIncludeSwitches  " -an -ai" ISWITCH_NO_WILDCARD_POSTFIX
#define kHashIncludeSwitches  kIncludeSwitch
#define kStopSwitchParsing  " --"

static NCompression::CInfo m_RegistryInfo;
extern HWND g_HWND;

static void ErrorMessage(LPCWSTR message)
{
  MessageBoxW(g_HWND, message, L"7-Zip ZS", MB_ICONERROR | MB_OK);
}

static void ErrorMessageHRESULT(HRESULT res, LPCWSTR s = NULL)
{
  UString s2 = NError::MyFormatMessage(res);
  if (s)
  {
    s2.Add_LF();
    s2 += s;
  }
  ErrorMessage(s2);
}

static HRESULT Call7zGui(const UString &params,
    // LPCWSTR curDir,
    bool waitFinish,
    NSynchronization::CBaseEvent *event)
{
  UString imageName = fs2us(NWindows::NDLL::GetModuleDirPrefix());
  imageName += k7zGui;

  CProcess process;
  const WRes wres = process.Create(imageName, params, NULL); // curDir);
  if (wres != 0)
  {
    const HRESULT hres = HRESULT_FROM_WIN32(wres);
    ErrorMessageHRESULT(hres, imageName);
    return hres;
  }
  if (waitFinish)
    process.Wait();
  else if (event != NULL)
  {
    HANDLE handles[] = { process, *event };
    ::WaitForMultipleObjects(Z7_ARRAY_SIZE(handles), handles, FALSE, INFINITE);
  }
  return S_OK;
}

static void AddLagePagesSwitch(UString &params)
{
  if (ReadLockMemoryEnable())
  #ifndef UNDER_CE
  if (NSecurity::Get_LargePages_RiskLevel() == 0)
  #endif
    params += " -slp";
}

class CRandNameGenerator
{
  CRandom _random;
public:
  CRandNameGenerator() { _random.Init(); }
  void GenerateName(UString &s, const char *prefix)
  {
    s += prefix;
    s.Add_UInt32((UInt32)(unsigned)_random.Generate());
  }
};

static HRESULT CreateMap(const UStringVector &names,
    CFileMapping &fileMapping, NSynchronization::CManualResetEvent &event,
    UString &params)
{
  size_t totalSize = 1;
  {
    FOR_VECTOR (i, names)
      totalSize += (names[i].Len() + 1);
  }
  totalSize *= sizeof(wchar_t);
  
  CRandNameGenerator random;

  UString mappingName;
  for (;;)
  {
    random.GenerateName(mappingName, "7zMap");
    const WRes wres = fileMapping.Create(PAGE_READWRITE, totalSize, GetSystemString(mappingName));
    if (fileMapping.IsCreated() && wres == 0)
      break;
    if (wres != ERROR_ALREADY_EXISTS)
      return HRESULT_FROM_WIN32(wres);
    fileMapping.Close();
  }
  
  UString eventName;
  for (;;)
  {
    random.GenerateName(eventName, "7zEvent");
    const WRes wres = event.CreateWithName(false, GetSystemString(eventName));
    if (event.IsCreated() && wres == 0)
      break;
    if (wres != ERROR_ALREADY_EXISTS)
      return HRESULT_FROM_WIN32(wres);
    event.Close();
  }

  params.Add_Char('#');
  params += mappingName;
  params.Add_Colon();
  char temp[32];
  ConvertUInt64ToString(totalSize, temp);
  params += temp;
  
  params.Add_Colon();
  params += eventName;

  LPVOID data = fileMapping.Map(FILE_MAP_WRITE, 0, totalSize);
  if (!data)
    return E_FAIL;
  CFileUnmapper unmapper(data);
  {
    wchar_t *cur = (wchar_t *)data;
    *cur++ = 0; // it means wchar_t strings (UTF-16 in WIN32)
    FOR_VECTOR (i, names)
    {
      const UString &s = names[i];
      const unsigned len = s.Len() + 1;
      wmemcpy(cur, (const wchar_t *)s, len);
      cur += len;
    }
  }
  return S_OK;
}

int FindRegistryFormat(const UString &name)
{
  FOR_VECTOR (i, m_RegistryInfo.Formats)
  {
    const NCompression::CFormatOptions &fo = m_RegistryInfo.Formats[i];
    if (name.IsEqualTo_NoCase(GetUnicodeString(fo.FormatID)))
      return i;
  }
  return -1;
}

int FindRegistryFormatAlways(const UString &name)
{
  int index = FindRegistryFormat(name);
  if (index < 0)
  {
    NCompression::CFormatOptions fo;
    fo.FormatID = GetSystemString(name);
    index = m_RegistryInfo.Formats.Add(fo);
  }
  return index;
}

HRESULT CompressFiles(
    const UString &arcPathPrefix,
    const UString &arcName,
    const UString &arcType,
    bool addExtension,
    const UStringVector &names,
    bool email, bool showDialog, bool waitFinish)
{
  MY_TRY_BEGIN
  UString params ('a');

  CFileMapping fileMapping;
  NSynchronization::CManualResetEvent event;
  params += kIncludeSwitch;
  RINOK(CreateMap(names, fileMapping, event, params))

  if (!arcType.IsEmpty() && ((arcType == L"7z") || (arcType == L"zip")))
  {
    int index;
    params += kArchiveTypeSwitch;
    params += arcType;
    m_RegistryInfo.Load();
    index = FindRegistryFormatAlways(arcType);
    if (index >= 0)
    {
      char temp[64];
      const NCompression::CFormatOptions &fo = m_RegistryInfo.Formats[index];

      if (!fo.Method.IsEmpty())
      {
        params += (arcType == L"7z") ? " -m0=" : " -mm=";
        params += fo.Method;
      }

      if (fo.Level)
      {
        params += " -mx=";
        ConvertUInt32ToString(fo.Level, temp);
        params += temp;
      }

      if (fo.Dictionary && (arcType == L"7z"))
      {
        params += " -md=";
        ConvertUInt32ToString(fo.Dictionary, temp);
        params += temp;
        params += "b";
      }

      if (fo.BlockLogSize && (arcType == L"7z"))
      {
        params += " -ms=";
        ConvertUInt64ToString(1ULL << fo.BlockLogSize, temp);
        params += temp;
        params += "b";
      }

      if (fo.NumThreads && fo.NumThreads != -1)
      {
        params += " -mmt=";
        ConvertUInt32ToString(fo.NumThreads, temp);
        params += temp;
      }

      if (!fo.Options.IsEmpty())
      {
        UStringVector strings;
        SplitString(fo.Options, strings);
        FOR_VECTOR (i, strings)
        {
          params += " -m";
          params += strings[i];
        }
      }
    }
  }

  if (email)
    params += kEmailSwitch;

  if (showDialog)
    params += kShowDialogSwitch;

  AddLagePagesSwitch(params);

  if (arcName.IsEmpty())
    params += " -an";

  if (addExtension)
    params += " -saa";
  else
    params += " -sae";

  params += kStopSwitchParsing;
  params.Add_Space();
  
  if (!arcName.IsEmpty())
  {
    params += GetQuotedString(
    // #ifdef UNDER_CE
      arcPathPrefix +
    // #endif
    arcName);
  }
  
  // ErrorMessage(params);
  return Call7zGui(params,
      // (arcPathPrefix.IsEmpty()? 0: (LPCWSTR)arcPathPrefix),
      waitFinish, &event);
  MY_TRY_FINISH
}

static void ExtractGroupCommand(const UStringVector &arcPaths, UString &params, bool isHash)
{
  AddLagePagesSwitch(params);
  params += (isHash ? kHashIncludeSwitches : kArcIncludeSwitches);
  CFileMapping fileMapping;
  NSynchronization::CManualResetEvent event;
  HRESULT result = CreateMap(arcPaths, fileMapping, event, params);
  if (result == S_OK)
    result = Call7zGui(params, false, &event);
  if (result != S_OK)
    ErrorMessageHRESULT(result);
}

void ExtractArchives(const UStringVector &arcPaths, const UString &outFolder, bool showDialog, bool elimDup, UInt32 writeZone)
{
  MY_TRY_BEGIN
  UString params ('x');
  if (!outFolder.IsEmpty())
  {
    params += " -o";
    params += GetQuotedString(outFolder);
  }
  if (elimDup)
    params += " -spe";
  if (writeZone != (UInt32)(Int32)-1)
  {
    params += " -snz";
    params.Add_UInt32(writeZone);
  }
  if (showDialog)
    params += kShowDialogSwitch;
  ExtractGroupCommand(arcPaths, params, false);
  MY_TRY_FINISH_VOID
}


void SmartExtractArchives(const UStringVector &arcPaths, const UString &baseFolder, UInt32 writeZone) {
  MY_TRY_BEGIN

      // For each archive, determine smart extraction path
      for (unsigned i = 0; i < arcPaths.Size(); i++) {
    const UString &arcPath = arcPaths[i];

    // Get archive file name without path
    UString arcName;
    int slashPos = arcPath.ReverseFind_PathSepar();
    if (slashPos >= 0)
      arcName = arcPath.Ptr(slashPos + 1);
    else
      arcName = arcPath;

    // Get folder name by removing extension
    UString folderName;
    int dotPos = arcName.ReverseFind_Dot();
    if (dotPos > 0)
      folderName = arcName.Left(dotPos);
    else
      folderName = arcName;

    // Check for multi-part archives (.001, .part01, etc.)
    folderName.TrimRight();
    dotPos = folderName.ReverseFind_Dot();
    if (dotPos > 0) {
      const UString ext2 = folderName.Ptr(dotPos + 1);
      if (ext2.IsEqualTo_Ascii_NoCase("001") ||
          ext2.IsEqualTo_Ascii_NoCase("part001") ||
          ext2.IsEqualTo_Ascii_NoCase("part01") ||
          ext2.IsEqualTo_Ascii_NoCase("part1")) {
        folderName.DeleteFrom(dotPos);
        folderName.TrimRight();
      }
    }

    // Use 7z.exe to list archive contents
    UString imageName = fs2us(NWindows::NDLL::GetModuleDirPrefix());
    imageName += L"7z.exe";

    UString cmdLine = GetQuotedString(imageName);
    cmdLine += L" l -slt ";
    cmdLine += GetQuotedString(arcPath);

    // Create pipe for capturing output
    HANDLE hReadPipe, hWritePipe;
    SECURITY_ATTRIBUTES sa;
    sa.nLength = sizeof(SECURITY_ATTRIBUTES);
    sa.bInheritHandle = TRUE;
    sa.lpSecurityDescriptor = NULL;

    AString content;
    bool captureSuccess = false;

    if (::CreatePipe(&hReadPipe, &hWritePipe, &sa, 0)) {
      ::SetHandleInformation(hReadPipe, HANDLE_FLAG_INHERIT, 0);

      STARTUPINFOW si;
      memset(&si, 0, sizeof(si));
      si.cb = sizeof(si);
      si.dwFlags = STARTF_USESHOWWINDOW | STARTF_USESTDHANDLES;
      si.wShowWindow = SW_HIDE;
      si.hStdOutput = hWritePipe;
      si.hStdError = hWritePipe;

      PROCESS_INFORMATION pi;
      memset(&pi, 0, sizeof(pi));

      if (::CreateProcessW(NULL, (LPWSTR)(LPCWSTR)cmdLine, NULL, NULL, TRUE,
                           CREATE_NO_WINDOW, NULL, NULL, &si, &pi)) {
        ::CloseHandle(hWritePipe);

        // Read output from pipe
        char buffer[4096];
        DWORD bytesRead;
        while (
            ::ReadFile(hReadPipe, buffer, sizeof(buffer), &bytesRead, NULL) &&
            bytesRead > 0) {
          content.AddFrom(buffer, bytesRead);
        }

        ::WaitForSingleObject(pi.hProcess, INFINITE);
        ::CloseHandle(pi.hThread);
        ::CloseHandle(pi.hProcess);
        ::CloseHandle(hReadPipe);
        captureSuccess = true;
      } else {
        ::CloseHandle(hReadPipe);
        ::CloseHandle(hWritePipe);
      }
    }

    // If failed to capture output, fallback to default extraction
    if (!captureSuccess) {
      UString outFolder = baseFolder;
      NFile::NName::NormalizeDirPathPrefix(outFolder);
      outFolder += folderName;
      UStringVector singleArc;
      singleArc.Add(arcPath);
      ExtractArchives(singleArc, outFolder, false, false, writeZone);
      continue;
    }

    // Parse the output to count top-level items
    int topLevelCount = 0;
    UString singleTopLevelName;
    bool hasSingleTopLevelFolder = false;
    UStringVector topLevelNames;

    {
      // Parse content line by line to find all paths
      int lineStart = 0;

      // Helper lambda to extract and clean path from "Path = " line
      auto extractPath = [](const AString &line) -> UString {
        AString pathStr = line.Mid(7, line.Len() - 7);
        UString path = GetUnicodeString(pathStr);
        // Remove trailing whitespace
        while (path.Len() > 0) {
          wchar_t c = path.Back();
          if (c == L'\r' || c == L'\n' || c == L' ' || c == L'\t') {
            path.DeleteBack();
          } else {
            break;
          }
        }
        return path;
      };

      for (int pos = 0; pos <= (int)content.Len(); pos++) {
        if (pos == (int)content.Len() || content[pos] == '\n') {
          AString line = content.Mid(lineStart, pos - lineStart);
          line.Trim();
          lineStart = pos + 1;

          if (line.IsPrefixedBy("Path = ")) {
            UString currentPath = extractPath(line);

            // Skip paths with colon (archive paths, not contents)
            if (currentPath.IsEmpty() || currentPath.Find(L':') >= 0) {
              continue;
            }

            // Extract top-level name
            int sepPos = currentPath.Find(WCHAR_PATH_SEPARATOR);
            UString topName =
                (sepPos >= 0) ? currentPath.Left(sepPos) : currentPath;

            // Add to list if not already present
            bool found = false;
            for (unsigned j = 0; j < topLevelNames.Size(); j++) {
              if (topLevelNames[j] == topName) {
                found = true;
                break;
              }
            }
            if (!found) {
              topLevelNames.Add(topName);
            }
          }
        }
      }

      topLevelCount = topLevelNames.Size();

      // Check if single top-level item is a folder
      if (topLevelCount == 1) {
        singleTopLevelName = topLevelNames[0];
        hasSingleTopLevelFolder = false;
        lineStart = 0;

        for (int pos = 0; pos <= (int)content.Len(); pos++) {
          if (pos == (int)content.Len() || content[pos] == '\n') {
            AString line = content.Mid(lineStart, pos - lineStart);
            line.Trim();
            lineStart = pos + 1;

            if (line.IsPrefixedBy("Path = ")) {
              UString path = extractPath(line);
              if (!path.IsEmpty() && path.Find(L':') < 0 &&
                  path.Find(WCHAR_PATH_SEPARATOR) >= 0) {
                hasSingleTopLevelFolder = true;
                break;
              }
            }
          }
        }
      }
    }

    // Decide extraction path based on top-level item count
    UString outFolder = baseFolder;
    NFile::NName::NormalizeDirPathPrefix(outFolder);

    if (topLevelCount == 0) {
      // Failed to analyze archive, fallback to creating subfolder
      outFolder += folderName;
    } else if (topLevelCount == 1 && hasSingleTopLevelFolder) {
      // Single top-level folder: extract directly to base folder
      // The folder inside archive will be extracted as-is
    } else if (topLevelCount == 1) {
      // Single top-level file: extract directly to base folder
    } else {
      // Multiple top-level items: extract to subfolder with archive name
      outFolder += folderName;
    }

    // Extract the archive
    UStringVector singleArc;
    singleArc.Add(arcPath);
    ExtractArchives(singleArc, outFolder, false, false, writeZone);
  }

  MY_TRY_FINISH_VOID
}

void TestArchives(const UStringVector &arcPaths, bool hashMode)
{
  MY_TRY_BEGIN
  UString params ('t');
  if (hashMode)
  {
    params += kArchiveTypeSwitch;
    params += "hash";
  }
  ExtractGroupCommand(arcPaths, params, false);
  MY_TRY_FINISH_VOID
}


void CalcChecksum(const UStringVector &paths,
    const UString &methodName,
    const UString &arcPathPrefix,
    const UString &arcFileName)
{
  MY_TRY_BEGIN

  if (!arcFileName.IsEmpty())
  {
    CompressFiles(
      arcPathPrefix,
      arcFileName,
      UString("hash"),
      false, // addExtension,
      paths,
      false, // email,
      false, // showDialog,
      false  // waitFinish
      );
    return;
  }

  UString params ('h');
  if (!methodName.IsEmpty())
  {
    params += " -scrc";
    params += methodName;
    /*
    if (!arcFileName.IsEmpty())
    {
      // not used alternate method of generating file
      params += " -scrf=";
      params += GetQuotedString(arcPathPrefix + arcFileName);
    }
    */
  }
  ExtractGroupCommand(paths, params, true);
  MY_TRY_FINISH_VOID
}

void Benchmark(bool totalMode)
{
  MY_TRY_BEGIN
  UString params ('b');
  if (totalMode)
    params += " -mm=*";
  AddLagePagesSwitch(params);
  const HRESULT result = Call7zGui(params, false, NULL);
  if (result != S_OK)
    ErrorMessageHRESULT(result);
  MY_TRY_FINISH_VOID
}
