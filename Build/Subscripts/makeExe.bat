@echo off
echo =================================
echo Building %1

call setVersion
set FILENAME=%1%TWUTILS-MAJOR%%TWUTILS-MINOR%.exe

echo Setting version = %TWUTILS-MAJOR%.%TWUTILS-MINOR%.%TWUTILS-REVISION%
setprojectcomp %1\%1.vbp %TWUTILS-REVISION% -mode:N
if errorlevel 1 goto :err

vb6 /m %1\%1.vbp
if errorlevel 1 goto :err

generateAssemblyManifest %1 exe EMBED %2
if errorlevel 1 goto :err

exit /B 0

:err
pause
exit /B 1

