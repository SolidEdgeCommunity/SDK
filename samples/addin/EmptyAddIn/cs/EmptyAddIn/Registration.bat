:: If you are changing <OutputPath>bin\Debug\</OutputPath> to something like ..\..\..\MyProduct,
:: you are probably doing it wrong and this script will not work for you.

:: Execute WHERE RegAsm.exe from a command prompt like the Developer Command Prompt for VS 2019.
:: It will return C:\Windows\Microsoft.NET\Framework\v4.0.30319\RegAsm.exe which is the 32 bit version.
:: Using the 32 bit version of RegAsm.exe will not correctly register your addin for modern Solid Edge (x64).
:: %SystemRoot%\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe is the 64 bit version and is the correct
:: version for registering modern Solid Edge AddIns.

@echo off

:: Update to match <AssemblyName> as necessary.
set ASSEMBLY_NAME=EmptyAddIn

:: If you left <OutputPath> alone, this will resolve to the correct path.
set ADDIN_PATH="%~dp0bin\Debug\%ASSEMBLY_NAME%.dll"

:: set REGASM_X32="%SystemRoot%\Microsoft.NET\Framework\v4.0.30319\RegAsm.exe"
set REGASM_X64="%SystemRoot%\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe"

CLS

WHOAMI /Groups | FIND "S-1-16-12288" >NUL
IF ERRORLEVEL 1 (
	ECHO This batch file requires elevated privileges.
	EXIT /B 1
)

:menu
echo [Options]
echo 1 Register %ADDIN_PATH%
echo 2 Unregister %ADDIN_PATH%
echo 3 Quit

:choice
set /P C=Enter selection:
if "%C%"=="1" goto register
if "%C%"=="2" goto unregister
if "%C%"=="3" goto end
goto choice

:register
echo.
%REGASM_X64% /codebase %ADDIN_PATH%
goto end

:unregister
echo.
%REGASM_X64% /u %ADDIN_PATH%
goto end

:end
