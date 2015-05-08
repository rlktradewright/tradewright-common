@echo off
setlocal

%TW-PROJECTS-DRIVE%
set BIN-PATH=%TW-PROJECTS-PATH%\Bin

call makeExternalManifest.bat TLBINF32 dll
call makeExternalManifest.bat MSCOMCTL ocx
