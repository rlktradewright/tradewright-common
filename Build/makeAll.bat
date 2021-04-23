@echo off
setlocal

call makeDlls.bat
call makeExes.bat 
call makeTradeWrightCommonExternalComponentsAssemblyManifest.bat

call makeComInteropFiles.bat

echo =================================
echo Make all completed
echo =================================
