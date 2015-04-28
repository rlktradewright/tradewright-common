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
call makedll TWUtilities dll %BINARY_COMPAT% compat

if not "%BINARY_COMPAT%" == "B" call makeSetProjectComp

call makedll ExtProps dll %BINARY_COMPAT%
call makedll ExtEvents dll %BINARY_COMPAT%
call makedll BusObjUtils dll %BINARY_COMPAT%
call makedll TWControls ocx %BINARY_COMPAT%
call makedll GraphicsUtils dll %BINARY_COMPAT%
call makedll LayeredGraphics dll %BINARY_COMPAT%
call makedll GraphObj dll %BINARY_COMPAT%
call makedll GraphObjUtils dll %BINARY_COMPAT%
popd

pushd %TW-PROJECTS-PATH%\SampleApps\LayeredGraphicsTest\
call makedll SpriteControlLib dll %BINARY_COMPAT%
popd

call includeExternalLibrary TLBINF32 dll
call includeExternalLibrary MSCOMCTL ocx

