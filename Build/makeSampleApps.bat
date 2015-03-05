@echo off
set PROJECT-DRIVE=E:
set PROJECT-PATH=\Projects\tradewright-common

%PROJECT-DRIVE%
cd %PROJECT-PATH%

echo =================================
echo Building ClockTester
vb6 /m SampleApps\ClockTester\ClockTester.vbp
if errorlevel 1 pause

echo =================================
echo Building IntervalTimerTester
vb6 /m SampleApps\IntervalTimerTester\IntervalTimerTester.vbp
if errorlevel 1 pause

echo =================================
echo Building LayeredGraphicsTest1
vb6 /m SampleApps\LayeredGraphicsTest1\LayeredGraphicsTest1.vbp
if errorlevel 1 pause

echo =================================
echo Building TasksDemo
vb6 /m SampleApps\TasksDemo\TasksDemo.vbp
if errorlevel 1 pause
