@echo off

:: makeTradeWrightCommonProjects.bat
::
:: builds all the dll and ocx projects
::
:: Parameters:
::   %1 Binary compatibility setting- 'P' (project)or 'B' (binary)
::

set BINARY_COMPAT=B
if "%1" == "P" set BINARY_COMPAT=P
if "%1" == "B" set BINARY_COMPAT=B
if "%1" == "N" set BINARY_COMPAT=N

pushd %TW-PROJECTS-PATH%

:: note that we have to store the compatible version of
:: TWUtilities in the compat folder, because using the one
:: in Bin results in linker errors
call makedll.bat TWUtilities TWUtilities .dll %BINARY_COMPAT% /compat

if not "%BINARY_COMPAT%" == "B" call makeSetProjectComp.bat

call makedll.bat ExtProps ExtProps .dll %BINARY_COMPAT%
call makedll.bat ExtEvents ExtEvents .dll %BINARY_COMPAT%
call makedll.bat BusObjUtils BusObjUtils .dll %BINARY_COMPAT%
call makedll.bat TWControls TWControls .ocx %BINARY_COMPAT%
call makedll.bat GraphicsUtils GraphicsUtils .dll %BINARY_COMPAT%
call makedll.bat LayeredGraphics LayeredGraphics .dll %BINARY_COMPAT%
call makedll.bat GraphObjUtils GraphObjUtils .dll %BINARY_COMPAT%
call makedll.bat GraphObj GraphObj .dll %BINARY_COMPAT%
popd

pushd %TW-PROJECTS-PATH%\SampleApps\LayeredGraphicsTest\
call makedll.bat SpriteControlLib SpriteControlLib .dll %BINARY_COMPAT%
popd


