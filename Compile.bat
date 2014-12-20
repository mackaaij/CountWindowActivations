@ECHO OFF
SETLOCAL

:searchanau3
REM Check whether the number of .AU3 files equals 1. If not, abort script.
for /F %%i in ('dir /b *.au3 ^| find /c /v ""') do set numberoffiles=%%i
if NOT %numberoffiles%==1 echo There are NONE or more than one .AU3 files in this directory. & echo The script cannot determine which file to use. & echo Aborting script. & set error=true & goto :end

REM Find the name for the .Au3 file and place it in the variable filename.
for /F "delims=" %%i in ('dir *.au3 /b') do set filename=%%i
echo Found script file: %filename%
SET SCRIPT="%cd%\%filename%"
echo Script: %SCRIPT%

:searchicon
REM Check whether an icon can be found. If not, continue with default.
for /F %%i in ('dir /b *.ico ^| find /c /v ""') do set numberoffiles=%%i
if NOT %numberoffiles%==1 echo. & echo NO icon found. Continuing with default icon. & goto :noicon

REM Find the name for the .ico file and place it in the variable filename.
for /F "delims=" %%i in ('dir *.ico /b') do set filename=%%i
echo Found icon file: %filename%
SET ICON="%cd%\%filename%"
echo Icon: %ICON%

:yesicon
echo.
echo Compiling .EXE with icon...
"C:\Program Files\AutoIt3\Aut2Exe\Aut2Exe.exe" /in %SCRIPT% /nodecompile /icon %ICON%
goto :end

:noicon
echo.
echo Compiling .EXE with default icon...
"C:\Program Files\AutoIt3\Aut2Exe\Aut2Exe.exe" /in %SCRIPT% /nodecompile /icon "C:\Program Files\AutoIt3\Aut2Exe\Icons\AutoIt_HighColor.ico"

REM Script will only go here if no .au3 file is found since the call of aut2exe does not return
:end
if %error%==true color 4f

ECHO.
ECHO Press any key to close this screen.
PAUSE > nul
pause

ENDLOCAL