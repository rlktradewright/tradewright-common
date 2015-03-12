@echo off
echo =================================
echo Building %1

call setVersion
set FILENAME=%1%TWUTILS-MAJOR%%TWUTILS-MINOR%.exe

echo Setting version = %TWUTILS-MAJOR%.%TWUTILS-MINOR%.%TWUTILS-REVISION%
setprojectcomp %1\%1.vbp %TWUTILS-REVISION% -mode:N
if errorlevel 1 pause

vb6 /m %1\%1.vbp
if errorlevel 1 pause

if exist %1\%FILENAME%.manifest (
	echo Copying manifest to Bin
	copy %1\%FILENAME%.manifest %BIN-PATH%\
)

