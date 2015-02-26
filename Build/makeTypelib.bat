set PROJECT-DRIVE=E:
set PROJECT-PATH=\Projects\tradewright-common

%PROJECT-DRIVE%
cd %PROJECT-PATH%

cd typelib
:: you may need to edit the following to locate your copy of midl.exe. 
:: It's for Visual Studio 2008 installed in the default location
path C:\Program Files\Microsoft SDKs\Windows\v6.0A\bin;%PATH%
path C:\Program Files\Microsoft Visual Studio 9.0\VC\VCPackages;%PATH%
path C:\WINDOWS\Microsoft.NET\Framework\v2.0.50727;%PATH%
path C:\WINDOWS\Microsoft.NET\Framework\v3.5;%PATH%
path C:\Program Files\Microsoft Visual Studio 9.0\Common7\Tools;%PATH%
path C:\Program Files\Microsoft Visual Studio 9.0\VC\BIN;%PATH%
path C:\Program Files\Microsoft Visual Studio 9.0\Common7\IDE;%PATH%

set INCLUDE="C:\Program Files\Microsoft Visual Studio 9.0\VC\ATLMFC\INCLUDE;C:\Program Files\Microsoft Visual Studio 9.0\VC\INCLUDE;C:\Program Files\Microsoft SDKs\Windows\v6.0A\include;"
set "WindowsSdkDir=C:\Program Files\Microsoft SDKs\Windows\v6.0A\"

midl /mktyplib203 TWWin32API.idl
::pause

regtlib TWWin32API.tlb
::pause

