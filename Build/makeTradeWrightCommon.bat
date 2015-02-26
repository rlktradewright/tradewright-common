set PROJECT-DRIVE=E:
set PROJECT-PATH=\Projects\tradewright-common

%PROJECT-DRIVE%
cd %PROJECT-PATH%

cd typelib
midl /mktyplib203 TWWin32API.idl
rem pause
regtlib TWWin32API.tlb
cd ..
rem pause
vb6 /m TWUtilities\TWUtilities.vbp
if errorlevel 1 pause
vb6 /m ExtProps\ExtProps.vbp
if errorlevel 1 pause
vb6 /m ExtEvents\ExtEvents.vbp
if errorlevel 1 pause

vb6 /m BusObjUtils\BusObjUtils.vbp
if errorlevel 1 pause

vb6 /m TWControls\TWControls.vbp
if errorlevel 1 pause

vb6 /m GraphicsUtils\GraphicsUtils.vbp
if errorlevel 1 pause
vb6 /m LayeredGraphics\LayeredGraphics.vbp
if errorlevel 1 pause
vb6 /m GraphObjUtils\GraphObjUtils.vbp
if errorlevel 1 pause
vb6 /m GraphObj\GraphObj.vbp
if errorlevel 1 pause

vb6 /m LayeredGraphics\SpriteControlLib\SpriteControlLib.vbp
if errorlevel 1 pause
vb6 /m LayeredGraphics\LayeredGraphicsTest1\LayeredGraphicsTest1.vbp
if errorlevel 1 pause
