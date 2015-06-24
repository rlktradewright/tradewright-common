%TW-PROJECTS-DRIVE%
pushd %TW-PROJECTS-PATH%\src

vb6 /m Tools\SetProjectComp\SetProjectComp.vbp
if errorlevel 1 pause
link /EDIT /SUBSYSTEM:CONSOLE Tools\SetProjectComp\SetProjectComp.exe
if errorlevel 1 pause
copy Tools\SetProjectComp\SetProjectComp.exe %TW-PROJECTS-PATH%\Build\Tools\SetProjectComp.exe

popd