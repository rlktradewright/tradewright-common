set PROJECT-DRIVE=E:
set PROJECT-PATH=\Projects\tradewright-common

%PROJECT-DRIVE%
cd %PROJECT-PATH%

:: Ensure the Build folder is on the path for use of SetProjectComp.exe
path %PROJECT-DRIVE%%PROJECT-PATH%\Build;%PATH%

setprojectcomp TWUtilities\TWUtilities.vbp -mode:P
vb6 /m TWUtilities\TWUtilities.vbp
copy TWUtilities\TWUtilities40.dll TWUtilities\Compat\TWUtilities40.dll 
if errorlevel 1 pause
setprojectcomp TWUtilities\TWUtilities.vbp -mode:B

vb6 /m %PROJECT-PATH%\Tools\SetProjectComp\SetProjectComp.vbp
if errorlevel 1 pause
link /EDIT /SUBSYSTEM:CONSOLE %PROJECT-PATH%\Tools\SetProjectComp\SetProjectComp.exe
if errorlevel 1 pause
copy %PROJECT-PATH%\Tools\SetProjectComp\SetProjectComp.exe %PROJECT-PATH%\Build\SetProjectComp.exe

setprojectcomp ExtProps\ExtProps.vbp -mode:P
vb6 /m ExtProps\ExtProps.vbp
if errorlevel 1 pause
setprojectcomp ExtProps\ExtProps.vbp -mode:B

setprojectcomp ExtEvents\ExtEvents.vbp -mode:P
vb6 /m ExtEvents\ExtEvents.vbp
if errorlevel 1 pause
setprojectcomp ExtEvents\ExtEvents.vbp -mode:B

setprojectcomp BusObjUtils\BusObjUtils.vbp -mode:P
vb6 /m BusObjUtils\BusObjUtils.vbp
if errorlevel 1 pause
setprojectcomp BusObjUtils\BusObjUtils.vbp -mode:B

setprojectcomp TWControls\TWControls.vbp -mode:P
vb6 /m TWControls\TWControls.vbp
if errorlevel 1 pause
setprojectcomp TWControls\TWControls.vbp -mode:B

setprojectcomp GraphicsUtils\GraphicsUtils.vbp -mode:P
vb6 /m GraphicsUtils\GraphicsUtils.vbp
if errorlevel 1 pause
setprojectcomp GraphicsUtils\GraphicsUtils.vbp -mode:B

setprojectcomp LayeredGraphics\LayeredGraphics.vbp -mode:P
vb6 /m LayeredGraphics\LayeredGraphics.vbp
if errorlevel 1 pause
setprojectcomp LayeredGraphics\LayeredGraphics.vbp -mode:B

setprojectcomp GraphObjUtils\GraphObjUtils.vbp -mode:P
vb6 /m GraphObjUtils\GraphObjUtils.vbp
if errorlevel 1 pause
setprojectcomp GraphObjUtils\GraphObjUtils.vbp -mode:B

setprojectcomp GraphObj\GraphObj.vbp -mode:P
vb6 /m GraphObj\GraphObj.vbp
if errorlevel 1 pause
setprojectcomp GraphObj\GraphObj.vbp -mode:B

vb6 /m SampleApps\ClockTester\ClockTester.vbp
if errorlevel 1 pause

vb6 /m SampleApps\IntervalTimerTester\IntervalTimerTester.vbp
if errorlevel 1 pause

setprojectcomp SampleApps\LayeredGraphicsTest1\SpriteControlLib\SpriteControlLib.vbp -mode:P
vb6 /m SampleApps\LayeredGraphicsTest1\SpriteControlLib\SpriteControlLib.vbp
if errorlevel 1 pause
setprojectcomp SampleApps\LayeredGraphicsTest1\SpriteControlLib\SpriteControlLib.vbp -mode:B

vb6 /m SampleApps\LayeredGraphicsTest1\LayeredGraphicsTest1.vbp
if errorlevel 1 pause

vb6 /m SampleApps\TasksDemo\TasksDemo.vbp
if errorlevel 1 pause
