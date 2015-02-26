set PROJECT-DRIVE=E:
set PROJECT-PATH=\Projects\tradewright-common

%PROJECT-DRIVE%
cd %PROJECT-PATH%

:: Ensure the Build folder is on the path for use of SetProjectComp.exe
path %PROJECT-DRIVE%%PROJECT-PATH%\Build;%PATH%

setprojectcomp TWUtilities\TWUtilities.vbp -mode:P
vb6 /m TWUtilities\TWUtilities.vbp
copy TWUtilities\TWUtilities40.dll TWUtilities\Compat\TWUtilities40.dll 
rem pause

vb6 /m %PROJECT-PATH%\Tools\SetProjectComp\SetProjectComp.vbp
link /EDIT /SUBSYSTEM:CONSOLE %PROJECT-PATH%\Tools\SetProjectComp\SetProjectComp.exe
copy %PROJECT-PATH%\Tools\SetProjectComp\SetProjectComp.exe %PROJECT-PATH%\Build\SetProjectComp.exe

setprojectcomp TWUtilities\TWUtilities.vbp -mode:B

setprojectcomp ExtProps\ExtProps.vbp -mode:P
vb6 /m ExtProps\ExtProps.vbp
rem pause
setprojectcomp ExtProps\ExtProps.vbp -mode:B

setprojectcomp ExtEvents\ExtEvents.vbp -mode:P
vb6 /m ExtEvents\ExtEvents.vbp
rem pause
setprojectcomp ExtEvents\ExtEvents.vbp -mode:B

setprojectcomp BusObjUtils\BusObjUtils.vbp -mode:P
vb6 /m BusObjUtils\BusObjUtils.vbp
rem pause
setprojectcomp BusObjUtils\BusObjUtils.vbp -mode:B

setprojectcomp TWControls\TWControls.vbp -mode:P
vb6 /m TWControls\TWControls.vbp
rem pause
setprojectcomp TWControls\TWControls.vbp -mode:B

setprojectcomp GraphicsUtils\GraphicsUtils.vbp -mode:P
vb6 /m GraphicsUtils\GraphicsUtils.vbp
rem pause
setprojectcomp GraphicsUtils\GraphicsUtils.vbp -mode:B

setprojectcomp LayeredGraphics\LayeredGraphics.vbp -mode:P
vb6 /m LayeredGraphics\LayeredGraphics.vbp
rem pause
setprojectcomp LayeredGraphics\LayeredGraphics.vbp -mode:B

setprojectcomp GraphObjUtils\GraphObjUtils.vbp -mode:P
vb6 /m GraphObjUtils\GraphObjUtils.vbp
rem pause
setprojectcomp GraphObjUtils\GraphObjUtils.vbp -mode:B

setprojectcomp GraphObj\GraphObj.vbp -mode:P
vb6 /m GraphObj\GraphObj.vbp
rem pause
setprojectcomp GraphObj\GraphObj.vbp -mode:B

setprojectcomp LayeredGraphics\SpriteControlLib\SpriteControlLib.vbp -mode:P
vb6 /m LayeredGraphics\SpriteControlLib\SpriteControlLib.vbp
rem pause
setprojectcomp LayeredGraphics\SpriteControlLib\SpriteControlLib.vbp -mode:B

vb6 /m LayeredGraphics\LayeredGraphicsTest1\LayeredGraphicsTest1.vbp
rem pause
