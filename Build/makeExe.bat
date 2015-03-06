@echo off
echo =================================
echo Building %1

echo Setting revision version = %twutils-version%
setprojectcomp %1\%1.vbp %twutils-version% -mode:N
if errorlevel 1 pause

vb6 /m %1\%1.vbp
if errorlevel 1 pause

echo Setting revision version = 0
setprojectcomp %1\%1.vbp 0 -mode:N
if errorlevel 1 pause

if exist %1\%1%TW-VERSION%.exe.manifest (
	echo Copying manifest to Bin
	copy %1\%1%TW-VERSION%.exe.manifest %PROJECT-PATH%\Bin\
)

