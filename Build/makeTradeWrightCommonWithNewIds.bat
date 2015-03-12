@echo off

%TW-PROJECTS-DRIVE%
path %TW-PROJECTS-DRIVE%%TW-PROJECTS-PATH%\Build;%PATH%

set BIN-PATH=%TW-PROJECTS-PATH%\Bin

call makeTradeWrightCommonProjects P
