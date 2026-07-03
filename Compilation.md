# 7-Zip-zstd 编译指南 (Windows x64 MSVC)

> 必须在 **cmd.exe** 中执行，不能使用 PowerShell（PowerShell 中 `set` 不设置环境变量）。

## 1. 切换到 mine 分支（包含自定义功能）

```cmd
cd /d D:\work\7z\7-Zip-zstd
git checkout mine
git submodule update --init DarkMode/lib
```

## 2. 初始化 MSVC 开发环境

```cmd
call "D:\Program Files (x86)\tmp\Common7\Tools\VsDevCmd.bat" -arch=x64 -host_arch=x64
```

验证工具链可用：

```cmd
where nmake
where cl
```

## 3. 设置环境变量并编译

```cmd
cd /d D:\work\7z\7-Zip-zstd\CPP
set PLATFORM=x64
set VC=msvc
set SUBSYS=6.00
set APPVEYOR_BUILD_FOLDER=%CD%
build-it.cmd clean
build-it.cmd
```

- `build-it.cmd clean`：清理上次的构建产物（必须，否则可能触发 NMAKE U1071 错误）
- `build-it.cmd`：执行编译

## 4. 编译产物

输出目录：`CPP\bin-msvc-x64\`

包含以下文件：
- `7z.exe` / `7za.exe` / `7zz.exe` — 命令行版本
- `7zFM.exe` — 文件管理器
- `7z.dll` / `7zxa.dll` / `7za.dll` — 核心库
- `zstd.dll` / `brotli.dll` / `lz4.dll` / `lz5.dll` / `lizard.dll` / `flzma2.dll` — 编解码器插件
- `7-zip.dll` — Explorer 右键菜单扩展
- `7z.sfx` / `7zCon.sfx` — 自解压模块
- `Install.exe` / `Uninstall.exe` — 安装/卸载程序

## 注意事项

- 编译依赖 DarkMode 子模块，首次编译前必须执行 `git submodule update --init DarkMode/lib`
- 如果遇到 `NMAKE U1071: 目标"x64"无法在递归中构建`，先执行 `build-it.cmd clean` 再重新编译
- 如果遇到 `NMAKE U1073: don't know how to make 'x64\Darkmodelib.obj'`，说明子模块未初始化
