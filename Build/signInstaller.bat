@echo off
setlocal

%TW-PROJECTS-DRIVE%
path %TW-PROJECTS-DRIVE%%TW-PROJECTS-PATH%\Build\Subscripts;%PATH%

call setTradeWrightCommonVersion

echo =================================
echo Signing the TradeWright Common merge module
echo.

signtool sign /t http://timestamp.comodoca.com %TW-PROJECTS-PATH%\Build\Setup\Bin\TradeWrightCommon%VB6-BUILD-MAJOR%-%VB6-BUILD-MINOR%-%VB6-BUILD-REVISION%.msm
if errorlevel 1 (
	echo Error signing installer
	set ERROR=1
	goto :err
)

echo =================================
echo Signing the TradeWright Common External Components merge module
echo.

signtool sign /t http://timestamp.comodoca.com %TW-PROJECTS-PATH%\Build\Setup\Bin\TradeWrightCommonExternalComponents%VB6-BUILD-MAJOR%-%VB6-BUILD-MINOR%-%VB6-BUILD-REVISION%.msm
if errorlevel 1 (
	echo Error signing installer
	set ERROR=1
	goto :err
)

echo =================================
echo Signing the TradeWright Common installer
echo.

signtool sign /t http://timestamp.comodoca.com %TW-PROJECTS-PATH%\Build\Setup\Bin\TradeWrightCommonInstaller%VB6-BUILD-MAJOR%-%VB6-BUILD-MINOR%-%VB6-BUILD-REVISION%.msi
if errorlevel 1 (
	echo Error signing installer
	set ERROR=1
	goto :err
)

exit /B 0

:err
exit /B 1

