@echo off
title Kich hoat Office 2016 mien phi tuoitreit.vn!&cls&echo ============================================================================&echo #Phat trien boi Nguyen Huu Thang admin tuoitreit.vn&echo ============================================================================&echo.&echo #Ho tro:&echo - Microsoft Office Standard 2016&echo - Microsoft Office Professional Plus 2016&echo.&echo.&(if exist "%ProgramFiles%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles%\Microsoft Office\Office16")&(if exist "%ProgramFiles(x86)%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles(x86)%\Microsoft Office\Office16")&(for /f %%x in ('dir /b ..\root\Licenses16\proplusvl_kms*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul)&(for /f %%x in ('dir /b ..\root\Licenses16\proplusvl_mak*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul)&echo.&echo ============================================================================&echo Dang kich hoat vui long cho...&cscript //nologo ospp.vbs /unpkey:WFG99 >nul&cscript //nologo ospp.vbs /unpkey:DRTFM >nul&cscript //nologo ospp.vbs /unpkey:BTDRB >nul&cscript //nologo ospp.vbs /inpkey:XQNVK-8JYDB-WJ9W3-YJ8YR-WFG99 >nul&set i=1
:server
if %i%==1 set KMS_Sev=kms7.MSGuides.com
if %i%==2 set KMS_Sev=kms8.MSGuides.com
if %i%==3 set KMS_Sev=kms9.MSGuides.com
if %i%==4 goto notsupported
cscript //nologo ospp.vbs /sethst:%KMS_Sev% >nul&echo ============================================================================&echo.&echo.
cscript //nologo ospp.vbs /act | find /i "successful" && (echo.&echo ============================================================================&echo.&echo #Trang web: tuoitreit.vn&echo.&echo #Ho tro boi tuoitreit.vn&echo.&echo #Lien he voi tuoitreit.vn neu ban can ho tro.&echo.&echo #Cam on cac ban&echo #Chuc cac ban mot ngay vui ve!&echo.&echo ============================================================================&choice /n /c YN /m "Ban co muon tham dien dan tuoitreit.vn khong [Y,N]?" & if errorlevel 2 exit) || (echo Vui long kiem tra lai mang va tuong lua roi thu lai... & echo Xin cho mot chut... & echo. & echo. & set /a i+=1 & goto server)
explorer "https://tuoitreit.vn"&goto halt
:notsupported
echo.&echo ============================================================================&echo Phien ban ban Office dang dung khong duoc ho tro. Vui long kiem tra lai
:halt
pause
