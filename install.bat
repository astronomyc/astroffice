@echo off
set passsave=12345
:initialmenu
color 0d
title Astroffice
mode con:cols=45 lines=9
cls
echo /---------------WaitingProcess--------------\
echo /                                           \
echo /                1. Install                 \
echo /                2. Uninstall               \
echo /                3. Activate                \
echo /                                           \
echo /-------------------------------------------\
set /p stf1="Option: "
if %stf1%==1 set varpass=goinstallp
if %stf1%==2 set varpass=gouninstallp
if %stf1%==3 set varpass=goactivatep
if %stf1%==0 set varpass=rmrmofficep
if %stf1%==1 goto goinstall
if %stf1%==2 goto gouninstall
if %stf1%==3 goto goactivate
if %stf1%==0 goto rmrmoffice
goto initialmenu

:goactivate
goto passecho
:goactivatep
set /p pass=
if %pass%==%passsave% goto passcorrectac
goto initialmenu
:passcorrectac
Echo irm https://massgrave.dev/get iex| Clip.exe
start PowerShell
goto gofin


:rmrmoffice
goto passecho
:rmrmofficep
set /p pass=
if %pass%==%passsave% goto passcorrectrmrm
goto initialmenu
:passcorrectrmrm
del "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\StartUp\rmoffice.bat"
goto gofin

:goinstall
goto passecho
:goinstallp
set /p pass=
if %pass%==%passsave% goto passcorrectin
goto initialmenu
:passcorrectin

:configset
title SelectConfiguration
mode con:cols=45 lines=7
color 0b
cls
echo ----------------WaitingProcess---------------
echo -                                           -
echo -            Select Config to 1-5           -
echo -                                           -
echo ---------------------------------------------
set /p configsel= "Select Configuration: "
if %configsel%==1 set config=config1
if %configsel%==1 goto cfinstall
if %configsel%==2 set config=config2
if %configsel%==2 goto cfinstall
if %configsel%==3 set config=config3
if %configsel%==3 goto cfinstall
if %configsel%==4 set config=config4
if %configsel%==4 goto cfinstall
if %configsel%==5 set config=config5
if %configsel%==5 goto cfinstall
goto configset

:cfinstall
mode con:cols=45 lines=8
color 03
cls
title Installing
echo IIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIII
echo II                                         II
echo II                                         II
echo II               Installing                II
echo II                                         II
echo II                                         II
echo IIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIII
setup /configure configs/%config%.xml
goto copyrm

:copyrm
xcopy /Y /f "rmoffice.bat" "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\StartUp\"
goto gofin

:gouninstall
goto passecho
:gouninstallp
set /p pass=
if %pass%==%passsave% goto passcorrectun
goto initialmenu
:passcorrectun
mode con:cols=45 lines=8
color 04
cls
setup /configure uninstall.xml
goto gofin

:gofin
mode con:cols=45 lines=8
title Done
echo IIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIII
echo II                                         II
echo II                                         II
echo II      !!!!      COMPLETE      !!!!       II
echo II                                         II
echo II                                         II
echo IIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIII
pause>nul
goto initialmenu

:passecho
mode con:cols=45 lines=80
title PassWord
echo.
echo.
echo.
echo.
echo.
echo.
echo.
echo.
echo.
echo.
echo.
echo.
echo.
echo.
echo.
echo.
echo.
echo.
echo.
echo.
echo.
echo.
echo.
echo.
echo.
echo.
echo.
echo.
echo.
echo.
echo.
echo.
echo.
echo.
echo.
echo.
echo.
echo.
echo.
echo.
echo.
echo.
echo.
echo.
echo.
echo.
echo.
echo.
echo.
echo.
echo.
echo.
echo.
echo.
echo.
echo.
echo.
echo.
echo.
echo.
echo.
echo.
echo.
echo.
echo.
echo.
echo.
echo.
echo.
echo.
echo.
echo.
echo.
echo.
echo.
echo.
echo.
echo.
echo.
echo.
goto %varpass%