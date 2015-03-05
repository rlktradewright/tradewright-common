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

:: note that we have to store the compatible version of
:: TWUtilities in the compat folder, because using the one
:: in Bin results in linker errors
call makedll TWUtilities dll %binarycompat% compat

if %binarycompat% == "B" goto setprojectcompisvalid

vb6 /m %PROJECT-PATH%\Tools\SetProjectComp\SetProjectComp.vbp
if errorlevel 1 pause
link /EDIT /SUBSYSTEM:CONSOLE %PROJECT-PATH%\Tools\SetProjectComp\SetProjectComp.exe
if errorlevel 1 pause
copy %PROJECT-PATH%\Tools\SetProjectComp\SetProjectComp.exe %PROJECT-PATH%\Build\SetProjectComp.exe

:setprojectcompisvalid

call makedll ExtProps dll %binarycompat%
call makedll ExtEvents dll %binarycompat%
call makedll BusObjUtils dll %binarycompat%
call makedll TWControls ocx %binarycompat%
call makedll GraphicsUtils dll %binarycompat%
call makedll LayeredGraphics dll %binarycompat%
call makedll GraphObj dll %binarycompat%

pushd SampleApps\LayeredGraphicsTest1\
call makedll SpriteControlLib dll %binarycompat%
popd