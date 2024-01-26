@echo off

setlocal enabledelayedexpansion

for /f "tokens=2 delims= " %%a in ('whoami /user') do (
    set "SID=%%a"
)

echo Current User SID: %SID%
echo.

::Set variables
for /f "delims=" %%i in ('C:\es.exe -w -i -sort-date-recently-changed-ascending -n 1 ksomisc.exe') do set "KSOMISC=%%i"
set WPSCLOUD=Software\Kingsoft\wpscloud
set WPSSWITCH=Software\Kingsoft\Office\6.0\Common\kic\switch
set KWPSSHELLEXT=Software\Kingsoft\Office\6.0\kwpsshellext
set WPSPLUGIN=Software\Kingsoft\Office\6.0\plugins
set WPSCOMMON=Software\Kingsoft\Office\6.0\Common


::Handle WPS Cloud
echo Trying remove WPS Cloud from This PC...

echo.
taskkill /f /t /im wpscloudsvr.exe

"%KSOMISC%" -unregyunpanlightweight -forceperusermode
"%KSOMISC%" -unregqingdriveshellext extraParams:dHlwZT1yZWdQZXJzb25hbERyaXZlfHVzZXJJZD0=
reg add "HKU\%SID%\%WPSCLOUD%" /v EnableVirtualCloudDisk /t REG_SZ /d "false" /f
echo.

echo Disable start
reg add "HKU\%SID%\%WPSCLOUD%\usercenter\tray" /v TrayStartWithWpsStart /t REG_SZ /d "false" /f
reg add "HKU\%SID%\%WPSCLOUD%\usercenter\tray" /v autorunTrayLastOperation /t REG_SZ /d "close" /f
reg add "HKU\%SID%\%WPSCLOUD%\usercenter\all\wpsbox" /v myAddBoot /t REG_SZ /d "false" /f
reg delete "HKU\%SID%\Software\Microsoft\Windows\CurrentVersion\Run" /v KWpsBox /f
echo.

echo Disable ShellExt
reg add "HKU\%SID%\%WPSCLOUD%" /v SendBy /t REG_SZ /d "false" /f
reg add "HKU\%SID%\%WPSCLOUD%" /v UploadSync /t REG_SZ /d "false" /f
echo.


::Remove WPS ShellExt
set WTF=7zDjnEI6YM7mtS0JAsswww==
reg add "HKU\%SID%\%WPSSWITCH%\user" /v 6GjQUxiYYdH8NxOVJUdpOAjJeIHv/0n3dzAUyF0tGYo= /t REG_SZ /d "%WTF%" /f
reg add "HKU\%SID%\%WPSSWITCH%\user" /v PMMpgkp8ls/4qfEiWKL9gvL4eyv2r9wg2ea0HbDpK48= /t REG_SZ /d "%WTF%" /f
reg add "HKU\%SID%\%WPSSWITCH%\user" /v VSXFEuhyZ+7MHjXCpG0+tBIZpYyWuOAgJbjlxkCYroM= /t REG_SZ /d "%WTF%" /f
reg add "HKU\%SID%\%WPSSWITCH%\user" /v jTQDZ60PH0TgGVc5XPR/PU7FlhQ8sfv+Ir0usxuF3ZU= /t REG_SZ /d "%WTF%" /f
reg add "HKU\%SID%\%WPSSWITCH%\user" /v kaQm19O36bskO7uwj2PeAJN5951KEClWCpFJR5RHyJs= /t REG_SZ /d "%WTF%" /f
reg add "HKU\%SID%\%WPSSWITCH%\user" /v pyoJeTexMJjB7c+B5ylpHt+wVeomk+s2MWuw1Y4FoIg= /t REG_SZ /d "%WTF%" /f
reg add "HKU\%SID%\%WPSSWITCH%\user" /v qssOqcEcuBksgfkgNmgDpW3dpPs4ssTllPQRRIC7smA= /t REG_SZ /d "%WTF%" /f
reg add "HKU\%SID%\%WPSSWITCH%\user" /v uWEXwvpGEA0DSF5Ae3PdqUF1uYEKqK0txnQophtU3m4= /t REG_SZ /d "%WTF%" /f
reg add "HKU\%SID%\%WPSSWITCH%\user" /v qA32MuVupIR0D+2e/hYicA== /t REG_SZ /d "%WTF%" /f
reg add "HKU\%SID%\%WPSSWITCH%\notify" /v FtJNUj1Xt6rGUkBGNcM/nQ== /t REG_SZ /d "%WTF%" /f
reg add "HKU\%SID%\%KWPSSHELLEXT%" /v disableDocPropPage /t REG_DWORD /d "1" /f
reg add "HKU\%SID%\%KWPSSHELLEXT%\contextmenu\photo" /v supportedit /t REG_SZ /d "false" /f
reg add "HKU\%SID%\%WPSPLUGIN%\kcdiskclean" /v wpsSceneRecEnable /t REG_SZ /d "false" /f
reg add "HKU\%SID%\%WPSCOMMON%\ksomisc" /v VasContextMenuShellNewUser /t REG_SZ /d "false" /f
reg add "HKU\%SID%\%WPSPLUGIN%\ksomisc" /v cancelassopdf /t REG_SZ /d "true" /f
echo.


::Remove WPS Preview Window
echo Disable pdfwpsv
for /f "delims=" %%i in ('C:\es.exe -w -i -sort-date-recently-changed-ascending -n 1 pdfwspv.dll') do set "PDFWPSV=%%i"
"%KSOMISC%" -unregpdfwspv
reg add "HKU\%SID%\%WPSPLUGIN%\pdfwpsv" /v hasUnregisterByUser /t REG_DWORD /d "1" /f
regsvr32 /s /u "PDFWPSV"
reg delete "HKU\%SID%_Classes\.pdfwpsshellnew" /f
echo.


::Remove WPS Association
echo Removeing WPS association...
"%KSOMISC%" -uncompatiblemso
"%KSOMISC%" -unassoword
"%KSOMISC%" -unassopowerpnt
"%KSOMISC%" -unassoexcel
"%KSOMISC%" -unassopdf
"%KSOMISC%" -unassopic
"%KSOMISC%" -unassoofd
"%KSOMISC%" -unassoopg
"%KSOMISC%" -unassocad
"%KSOMISC%" -unassodwg
"%KSOMISC%" -unassodxf
"%KSOMISC%" -unassoepub
"%KSOMISC%" -unassomobi
"%KSOMISC%" -unregister
"%KSOMISC%" -unregmtfont
"%KSOMISC%" -unregksee -forceperusermode
"%KSOMISC%" -unassoofficexml
"%KSOMISC%" -unreg_wps
"%KSOMISC%" -unreg_wpp
"%KSOMISC%" -unreg_et
"%KSOMISC%" -unregksee
"%KSOMISC%" -unreg_kso
"%KSOMISC%" -unreg_kde
"%KSOMISC%" -unregoledefhndlr
"%KSOMISC%" -unregmtfont
"%KSOMISC%" -rebuildicon
set "keywords=WPS. KWPS. ET. Kingsoft.Office ksoet ksoapp ksoCustomProtocol ksolink ksopdf ksoqing KsoWebStartup ksowpp ksowps ksoWPS"
for %%k in (%keywords%) do (
    set "keyword=%%k"
    for /f "tokens=1,*" %%a in ('reg query "HKU\%SID%_Classes" /s /f "!keyword!" 2^>nul') do (
        set "key=%%a"
        set "value=%%b"
        echo Found match: !key!
        reg delete "!key!" /f
    )
)

echo Done.