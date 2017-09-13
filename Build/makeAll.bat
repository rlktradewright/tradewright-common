@echo off
setlocal

call makeDlls.bat
call makeExes.bat 
call makeTradeWrightCommonExternalComponentsAssemblyManifest.bat

pushd ..
::note we have to be in the tradewright-common folder to run makeComInteropFiles
call Build\makeComInteropFiles.bat
popd

echo =================================
echo Make all completed
echo =================================
