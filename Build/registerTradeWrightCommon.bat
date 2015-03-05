:: registers the TradeWright Common dlls

set PROJECT-DRIVE=E:
set BIN-PATH=\Projects\tradewright-common\Bin

%PROJECT-DRIVE%
cd %BIN-PATH%


regsvr32 TWUtilities40.dll

regsvr32 ExtProps40.dll

regsvr32 ExtEvents40.dll

regsvr32 BusObjUtils40.dll

regsvr32 TWControls40.ocx

regsvr32 GraphicsUtils40.dll

regsvr32 LayeredGraphics40.dll

regsvr32 GraphObjUtils40.dll

regsvr32 GraphObj40.dll

regsvr32 SpriteControlLib.dll


