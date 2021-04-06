@echo off
setlocal

:: un-registers the TradeWright Common dlls

path %CD%;%PATH%

call setTradeWrightCommonVersion

pushd Bin\TradeWright.Common

call unregisterComponent.bat TWUtilities dll
if errorlevel 1 goto :err

call unregisterComponent.bat ExtProps dll
if errorlevel 1 goto :err

call unregisterComponent.bat ExtEvents dll
if errorlevel 1 goto :err

call unregisterComponent.bat BusObjUtils dll
if errorlevel 1 goto :err

call unregisterComponent.bat TWControls ocx
if errorlevel 1 goto :err

call unregisterComponent.bat GraphicsUtils dll
if errorlevel 1 goto :err

call unregisterComponent.bat LayeredGraphics dll
if errorlevel 1 goto :err

call unregisterComponent.bat GraphObjUtils dll
if errorlevel 1 goto :err

call unregisterComponent.bat GraphObj dll
if errorlevel 1 goto :err

call unregisterComponent.bat SpriteControlLib dll
if errorlevel 1 goto :err

popd

pushd Bin\TradeWright.Common.ExternalComponents

call unregisterComponent.bat mscomctl OCX EXT
if errorlevel 1 goto :err

call unregisterComponent.bat TLBINF32 DLL EXT
if errorlevel 1 goto :err

popd

exit /B

:err
popd
pause
exit /B 1
