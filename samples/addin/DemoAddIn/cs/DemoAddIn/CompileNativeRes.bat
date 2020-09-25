:: Resource Compiler (rc.exe) is included in the Windows 10 SDK
:: and can be installed using the Visual Studio Installer.
:: https://docs.microsoft.com/en-us/windows/win32/menurc/resource-compiler

:: Solid Edge AddIns require native resources for BITMAP, PNG, etc.
:: If you examine the .csproj, you will see <Win32Resource>Native.res</Win32Resource>.
:: This instructs MSBUILD to embed Native.res into the $(TargetPath).

:: Update Native.rc and MyConstants.cs as needed.
:: This batch file must be executed manually but could be added
:: to pre-build events.
@echo off

set RC_PATH="Native.rc"
set RES_PATH="Native.res"

:: RC_PATH -> RES_PATH
rc.exe /fo%RES_PATH% %RC_PATH%