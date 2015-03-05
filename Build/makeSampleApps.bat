@echo off
set PROJECT-DRIVE=E:
set PROJECT-PATH=\Projects\tradewright-common

%PROJECT-DRIVE%

path %PROJECT-DRIVE%%PROJECT-PATH%\Build;%PATH%

pushd %PROJECT-PATH%\SampleApps

call setversion

call makeExe ClockTester
call makeExe IntervalTimerTester
call makeExe LayeredGraphicsTest
call makeExe TasksDemo

popd