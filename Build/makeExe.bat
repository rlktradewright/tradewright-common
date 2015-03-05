@echo off
echo =================================
echo Building %1
vb6 /m %1\%1.vbp
if errorlevel 1 pause
if exist %1\%1%TW-VERSION%.exe.manifest copy %1\%1%TW-VERSION%.exe.manifest %PROJECT-PATH%\Bin\

