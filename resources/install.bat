@echo off
echo Concluindo instalacao do GeoAssist...
echo.
regsvr32 /s SuperGrid.ocx

set pastaSistema=SYSTEM32
if exist %SystemRoot%\SYSWOW64\Wow64.dll set pastaSistema=SYSWOW64

if not exist %SystemRoot%\%pastaSistema%\COMDLG32.OCX copy COMDLG32.OCX %SystemRoot%\SYSTEM32
if not exist %SystemRoot%\%pastaSistema%\COMDLG32.DLL copy COMDLG32.DLL %SystemRoot%\SYSTEM32
if not exist %SystemRoot%\%pastaSistema%\msflxgrd.ocx copy msflxgrd.ocx %SystemRoot%\SYSTEM32
regsvr32 /s %SystemRoot%\%pastaSistema%\COMDLG32.DLL
regsvr32 /s %SystemRoot%\%pastaSistema%\COMDLG32.OCX
regsvr32 /s %SystemRoot%\%pastaSistema%\msflxgrd.ocx

echo.
echo Fim da instalacao.
