:: unregisters the TradeWright Common dlls

set PROJECT-DRIVE=E:
set BIN-PATH=\Projects\tradewright-common\Bin

%PROJECT-DRIVE%
cd %BIN-PATH%

regsvr32 TWUtilities40.dll -u

regsvr32 ExtProps40.dll -u

regsvr32 ExtEvents40.dll -u

regsvr32 BusObjUtils40.dll -u

regsvr32 TWControls40.ocx -u

regsvr32 GraphicsUtils40.dll -u

regsvr32 LayeredGraphics40.dll -u

regsvr32 GraphObjUtils40.dll -u

regsvr32 GraphObj40.dll -u

regsvr32 SpriteControlLib.dll -u


