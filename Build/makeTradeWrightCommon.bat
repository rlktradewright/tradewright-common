set PROJECT-DRIVE=E:
set PROJECT-PATH=\Projects\tradewright-common

%PROJECT-DRIVE%
cd %PROJECT-PATH%

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

vb6 /m SampleApps\ClockTester\ClockTester.vbp
if errorlevel 1 pause
vb6 /m SampleApps\IntervalTimerTester\IntervalTimerTester.vbp
if errorlevel 1 pause
vb6 /m SampleApps\LayeredGraphicsTest1\SpriteControlLib\SpriteControlLib.vbp
if errorlevel 1 pause
vb6 /m SampleApps\LayeredGraphicsTest1\LayeredGraphicsTest1.vbp
if errorlevel 1 pause
vb6 /m SampleApps\TasksDemo\TasksDemo.vbp
if errorlevel 1 pause
