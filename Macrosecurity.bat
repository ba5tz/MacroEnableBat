@echo off

:Menu
cls
color 2b
echo.
echo.
echo                                         =-=-=--=--=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
echo                                             MACRO SECURITY EXCEL ENABLE/DISABLE
echo                                         ----------------------------------------                                                                                           
echo                                                    ________  _ ____  _____ _    
echo                                                   /  __/\  \///   _\/  __// \   
echo                                                   ^|  \   \  / ^|  /  ^|  \  ^| ^|   
echo                                                   ^|  /_  /  \ ^|  \_ ^|  /_ ^| ^|_/\
echo                                                   \____\/__/\\\____/\____\\____/
echo.
echo                                         =-=-=--=--=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
color 0b                    
echo.
echo -------------------------
echo 1. Enable Macro Security
echo -------------------------
echo 2. Disable Macro Security
echo -------------------------
echo 3. Exit
echo -------------------------
echo.
echo.
echo.                                                                          Made by Andi Setiadi

set INPUT=
set /P INPUT= Silahkan Ketikan Pilhan (1 / 2 / 3) :
if /I '%INPUT%'=='1' goto aktifkan
if /I '%INPUT%'=='2' Goto Matikan
if /I '%INPUT%'=='3' Goto Keluar
echo Pilihan salah! Silahkan pilih kembali!
pause 
goto menu

:aktifkan
echo.
rem - silahkan sesuaikan dengan Versi office yg dipakai 16.0 atau 14.0 dll
reg add "HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Excel\Security" /v VBAWarnings /t REG_DWORD /d 1 /f
echo          *****    Macro enable   ******
pause
goto Menu

:Matikan
echo.
rem - silahkan sesuaikan dengan Versi office yg dipakai 16.0 atau 14.0 dll
reg add "HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Excel\Security" /v VBAWarnings /t REG_DWORD /d 2 /f
echo          *****    Macro disable   ******
pause
goto Menu

:Keluar
