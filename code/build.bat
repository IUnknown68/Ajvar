@echo off

call "C:\Program Files (x86)\Microsoft Visual Studio 11.0\VC\bin\vcvars32.bat"

set /p VERSION=<..\build\version.txt

"C:\Program Files (x86)\Git\bin\sh.exe" --login -c "git pull -v --progress origin"
IF ERRORLEVEL 1 GOTO errorHandling

devenv Ajvar.sln /Rebuild "Debug|Win32"
IF ERRORLEVEL 1 GOTO errorHandling

devenv Ajvar.sln /Rebuild "Release|Win32"
IF ERRORLEVEL 1 GOTO errorHandling

devenv Ajvar.sln /Rebuild "Debug|x64"
IF ERRORLEVEL 1 GOTO errorHandling

devenv Ajvar.sln /Rebuild "Release|x64"
IF ERRORLEVEL 1 GOTO errorHandling

xcopy ..\docs\* Z:\www\ajvar /d /i /y /s

cd ..\deploy
del Ajvar-%VERSION%.zip
"C:\Program Files\7-Zip\7z" a -tzip "Ajvar-%VERSION%.zip" Ajvar-%VERSION%
copy "Ajvar-%VERSION%.zip" Z:\www\data\data\

exit

:errorHandling
exit /b 1
