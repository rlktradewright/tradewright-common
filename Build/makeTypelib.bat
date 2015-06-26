@echo off
setlocal

%TW-PROJECTS-DRIVE%
path %TW-PROJECTS-DRIVE%%TW-PROJECTS-PATH%\Build;%PATH%

set BIN-PATH=%TW-PROJECTS-PATH%\Bin\TradeWright.Common

:: you may need to edit the following to locate your copy of midl.exe. 
:: It's for Visual Studio 2008 installed in the default location
path C:\Program Files\Microsoft SDKs\Windows\v6.0A\bin;%PATH%
path C:\Program Files\Microsoft Visual Studio 9.0\VC\VCPackages;%PATH%
path C:\WINDOWS\Microsoft.NET\Framework\v2.0.50727;%PATH%
path C:\WINDOWS\Microsoft.NET\Framework\v3.5;%PATH%
path C:\Program Files\Microsoft Visual Studio 9.0\Common7\Tools;%PATH%
path C:\Program Files\Microsoft Visual Studio 9.0\VC\BIN;%PATH%
path C:\Program Files\Microsoft Visual Studio 9.0\Common7\IDE;%PATH%

set INCLUDE="C:\Program Files\Microsoft Visual Studio 9.0\VC\ATLMFC\INCLUDE;C:\Program Files\Microsoft Visual Studio 9.0\VC\INCLUDE;C:\Program Files\Microsoft SDKs\Windows\v6.0A\include;"
set "WindowsSdkDir=C:\Program Files\Microsoft SDKs\Windows\v6.0A\"

:: we have to pushd into typelib to ensure midl
:: picks up the copy of oaidl.idl that's in there
pushd %TW-PROJECTS-PATH%\src\typelib

@echo on
midl /mktyplib203 TWWin32API.idl /out %BIN-PATH%
@echo off
if errorlevel 1 pause

@echo on
regtlib %BIN-PATH%\TWWin32API.tlb
@echo off
if errorlevel 1 pause

popd