@echo off
setlocal 

%TW-PROJECTS-DRIVE%
path %TW-PROJECTS-DRIVE%%TW-PROJECTS-PATH%\Build\Subscripts;%PATH%

set BIN-PATH=%TW-PROJECTS-PATH%\Bin

call setMyVersion.bat
call makeComponents.bat B
