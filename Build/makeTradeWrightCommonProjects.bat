@echo off

:: makeTradeWrightCommonProjects.bat
::
:: builds all the dll and ocx projects
::
:: Parameters:
::   %1 Binary compatibility ('P' or 'B')
::

set PROJECT-DRIVE=E:
set PROJECT-PATH=\Projects\tradewright-common
set BIN-PATH=%PROJECT-PATH%\Bin

%PROJECT-DRIVE%
cd %PROJECT-PATH%

:: Ensure the Build folder is on the path for use of SetProjectComp.exe
path %PROJECT-DRIVE%%PROJECT-PATH%\Build;%PATH%

call setVersion

set binarycompat="B"
if "%1" == "P" set binarycompat="P"
if "%1" == "B" set binarycompat="B"
if "%1" == "N" set binarycompat="N"

:: note that we have to store the compatible version of
:: TWUtilities in the compat folder, because using the one
:: in Bin results in linker errors
call makedll TWUtilities dll %binarycompat% compat

if not %binarycompat% == "B" call makeSetProjectComp

call makedll ExtProps dll %binarycompat%
call makedll ExtEvents dll %binarycompat%
call makedll BusObjUtils dll %binarycompat%
call makedll TWControls ocx %binarycompat%
call makedll GraphicsUtils dll %binarycompat%
call makedll LayeredGraphics dll %binarycompat%
call makedll GraphObj dll %binarycompat%
call makedll GraphObjUtils dll %binarycompat%

pushd SampleApps\LayeredGraphicsTest\
call makedll SpriteControlLib dll %binarycompat%
popd
