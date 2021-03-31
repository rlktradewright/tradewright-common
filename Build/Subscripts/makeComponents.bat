@echo off

:: makeComponents.bat
::
:: builds all the dll and ocx projects
::
:: Parameters:
::   %1 Binary compatibility setting- 'P' (project) or 'B' (binary)
::

set BINARY_COMPAT=B
if "%1" == "P" set BINARY_COMPAT=P
if "%1" == "B" set BINARY_COMPAT=B
if "%1" == "N" set BINARY_COMPAT=N

pushd %TW-PROJECTS-PATH%\src

echo =================================
echo Making components for TradeWright.Common
echo.

set BIN-PATH=%BIN-PATH%\TradeWright.Common

call makedll.bat TWUtilities TWUtilities /T:DLL /B:%BINARY_COMPAT%
if errorlevel 1 pause
call makedll.bat ExtProps ExtProps /T:DLL /B:%BINARY_COMPAT%
if errorlevel 1 pause
call makedll.bat ExtEvents ExtEvents /T:DLL /B:%BINARY_COMPAT%
if errorlevel 1 pause
call makedll.bat BusObjUtils BusObjUtils /T:DLL /B:%BINARY_COMPAT%
if errorlevel 1 pause
call makedll.bat TWControls TWControls /T:OCX /B:%BINARY_COMPAT%
if errorlevel 1 pause
call makedll.bat GraphicsUtils GraphicsUtils /T:DLL /B:%BINARY_COMPAT%
if errorlevel 1 pause
call makedll.bat LayeredGraphics LayeredGraphics /T:DLL /B:%BINARY_COMPAT%
if errorlevel 1 pause
call makedll.bat GraphObjUtils GraphObjUtils /T:DLL /B:%BINARY_COMPAT%
if errorlevel 1 pause
call makedll.bat GraphObj GraphObj /T:DLL /B:%BINARY_COMPAT%
if errorlevel 1 pause
popd

pushd %TW-PROJECTS-PATH%\src\SampleApps\LayeredGraphicsTest\
call makedll.bat SpriteControlLib SpriteControlLib /T:DLL /B:%BINARY_COMPAT%
if errorlevel 1 pause
popd


