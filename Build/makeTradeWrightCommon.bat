@echo off

%TW-PROJECTS-DRIVE%
path %TW-PROJECTS-DRIVE%%TW-PROJECTS-PATH%\Build;%PATH%
path %TW-PROJECTS-DRIVE%%TW-PROJECTS-PATH%\Build\Subscripts;%PATH%
path %TW-PROJECTS-DRIVE%%TW-PROJECTS-PATH%\Build\Tools;%PATH%
path C:\Program Files\Microsoft SDKs\Windows\v6.0A\bin;%PATH%

set BIN-PATH=%TW-PROJECTS-PATH%\Bin

call makeComponents B
