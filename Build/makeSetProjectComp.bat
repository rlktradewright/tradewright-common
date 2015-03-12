%TW-PROJECTS-DRIVE%
pushd %TW-PROJECTS-PATH%

vb6 /m \Tools\SetProjectComp\SetProjectComp.vbp
if errorlevel 1 pause
link /EDIT /SUBSYSTEM:CONSOLE \Tools\SetProjectComp\SetProjectComp.exe
if errorlevel 1 pause
copy \Tools\SetProjectComp\SetProjectComp.exe \Build\SetProjectComp.exe

popd