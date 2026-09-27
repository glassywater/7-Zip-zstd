@echo off
REM Build GUI installers of 7-Zip ZS (darkmode builds only)

SET COPYCMD=/Y /B
SET COPTS=-m0=lzma -mx9 -ms=on -mf=bcj2
SET URL=https://www.7-zip.org/a/7z2603.exe
SET VERSION=26.03
SET SZIP="C:\Program Files\7-Zip\7z.exe"

SET WD=%cd%
SET SKEL=%WD%\skel

REM Download our skeleton files
mkdir %SKEL%
cd %SKEL%
curl %URL% -L -o 7-Zip.exe
IF %errorlevel% NEQ 0 EXIT 1
%SZIP% x 7-Zip.exe
IF %errorlevel% NEQ 0 EXIT 1
del 7-Zip.exe
goto start

@rem Doit function
:doit
SET ARCH=%~1
SET ZIP32=%~2
SET BIN=%~3
echo Doing %ARCH% in SOURCE=%BIN%

cd %SKEL%
del *.exe *.dll *.sfx
FOR %%f IN (7z.dll 7z.exe 7z.sfx 7za.dll 7za.exe 7zCon.sfx 7zFM.exe 7zG.exe 7-zip.dll 7zxa.dll Uninstall.exe) DO (
  copy %BIN%\%%f %%f || EXIT 1
)
IF NOT "%ZIP32%" == "" (
  copy %ZIP32% 7-zip32.dll || EXIT 1
)
%SZIP% a ..\%ARCH%.7z %COPTS% || EXIT 1
cd %WD%
copy %BIN%\Install.exe + %ARCH%.7z 7z%VERSION%-zstd-%ARCH%.exe || EXIT 1
del %ARCH%.7z
goto :eof
REM end of doit function.

:start

call :doit x86   ""                        "%WD%\bin-x86"   || EXIT 1
call :doit x64   "%WD%\bin-x86\7-zip.dll"  "%WD%\bin-x64"   || EXIT 1
call :doit arm64 ""                        "%WD%\bin-arm64" || EXIT 1

REM Cleanup
cd %WD%
rd /S /Q %SKEL%
