set PROJECT-DRIVE=E:
set PROJECT-PATH=\Projects\tradewright-common
set BIN-PATH=%PROJECT-PATH%\Bin

%PROJECT-DRIVE%
cd %PROJECT-PATH%

vb6 /m %PROJECT-PATH%\Tools\SetProjectComp\SetProjectComp.vbp
if errorlevel 1 pause
link /EDIT /SUBSYSTEM:CONSOLE %PROJECT-PATH%\Tools\SetProjectComp\SetProjectComp.exe
if errorlevel 1 pause
copy %PROJECT-PATH%\Tools\SetProjectComp\SetProjectComp.exe %PROJECT-PATH%\Build\SetProjectComp.exe
