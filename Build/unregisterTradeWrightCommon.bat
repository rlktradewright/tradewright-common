:: unregisters the TradeWright Common dlls

%TW-PROJECTS-DRIVE%
pushd %TW-PROJECTS-DRIVE%%TW-PROJECTS-PATH%\Bin

regsvr32 TWUtilities40.dll -u

regsvr32 ExtProps40.dll -u

regsvr32 ExtEvents40.dll -u

regsvr32 BusObjUtils40.dll -u

regsvr32 TWControls40.ocx -u

regsvr32 GraphicsUtils40.dll -u

regsvr32 LayeredGraphics40.dll -u

regsvr32 GraphObjUtils40.dll -u

regsvr32 GraphObj40.dll -u

regsvr32 SpriteControlLib40.dll -u

popd
