::@echo off
setlocal

%TW-PROJECTS-DRIVE%
path %TW-PROJECTS-DRIVE%%TW-PROJECTS-PATH%\Build;%PATH%

set BIN-PATH=%TW-PROJECTS-PATH%\Bin\TradeWright.Common

if not exist "%BIN-PATH%" mkdir "%BIN-PATH%"

set PROGFILES=%ProgramFiles%
if exist "%ProgramFiles(x86)%" set PROGFILES=%ProgramFiles(x86)%

:: you may need to edit the following to locate your copy of midl.exe. 
:: It's for Visual Studio 2008 installed in the default location
path %ProgramFiles%\Microsoft SDKs\Windows\v6.0A\bin;%PATH%
path %PROGFILES%\Microsoft Visual Studio 9.0\VC\VCPackages;%PATH%
path %SystemRoot%\Microsoft.NET\Framework\v2.0.50727;%PATH%
path %SystemRoot%\Microsoft.NET\Framework\v3.5;%PATH%
path %PROGFILES%\Microsoft Visual Studio 9.0\Common7\Tools;%PATH%
path %PROGFILES%\Microsoft Visual Studio 9.0\VC\BIN;%PATH%
path %PROGFILES%\Microsoft Visual Studio 9.0\Common7\IDE;%PATH%

set INCLUDE="%PROGFILES%\Microsoft Visual Studio 9.0\VC\ATLMFC\INCLUDE;%PROGFILES%\Microsoft Visual Studio 9.0\VC\INCLUDE;%ProgramFiles%\Microsoft SDKs\Windows\v6.0A\include;"
set WindowsSdkDir=%ProgramFiles%\Microsoft SDKs\Windows\v6.0A\

:: we have to pushd into typelib to ensure midl
:: picks up the copy of oaidl.idl that's in there
pushd %TW-PROJECTS-PATH%\src\typelib

dir

@echo on
midl /mktyplib203 TWWin32API.idl /out %BIN-PATH%
@echo off
if errorlevel 1 pause

@echo on
regtlib %BIN-PATH%\TWWin32API.tlb
@echo off
if errorlevel 1 pause

popd