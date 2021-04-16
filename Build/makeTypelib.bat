@echo off
setlocal

%TW-PROJECTS-DRIVE%
path %TW-PROJECTS-DRIVE%%TW-PROJECTS-PATH%\Build;%PATH%

set BIN-PATH=%TW-PROJECTS-PATH%\Bin\TradeWright.Common

if not exist "%BIN-PATH%" mkdir "%BIN-PATH%"

:: we have to pushd into typelib to ensure midl
:: picks up the copy of oaidl.idl that's in there
pushd %TW-PROJECTS-PATH%\src\typelib

@echo on
midl /mktyplib203 TWWin32API.idl /out %BIN-PATH%
@echo off
if errorlevel 1 goto :err

@echo on
regtlib %BIN-PATH%\TWWin32API.tlb
@echo off
if errorlevel 1 goto :err

popd
exit /B 0

:err
pause
exit /B 1
