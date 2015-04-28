@echo off
echo =================================
echo Including %1 %2

echo Generating manifest for %BIN-PATH%\%1.%2 to %BIN-PATH%\%1.manifest
GenerateManifest /Bin:%BIN-PATH%\%1.%2 /Out:%BIN-PATH%\%1.manifest
if errorlevel 1 goto :err

:: ensure mt.exe can find object files when hashing
pushd %BIN-PATH%

echo Updating manifest hash
mt.exe -manifest %BIN-PATH%\%1.manifest -hashupdate -nologo
if errorlevel 1 (
	popd
	goto :err
)
popd

exit /B 0

:err
exit /B 1


