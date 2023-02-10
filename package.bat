@echo off

echo Removing old package files

pushd %1 & for %%i in (.) do set curr=%%~ni

rd /s/q  %cd%\%curr%

del /f/s/q  %cd%\%curr%.zip

echo Pyinstaller start packing

pyinstaller -w -F --clean --distpath %cd%\%curr% %cd%/src/%curr%.py

xcopy .\icon\ %curr%\icon\ /e

echo Compressing package files

"C:\Program Files\7-Zip\7z.exe" a "%cd%\%curr%.zip" "%cd%\%curr%"

echo Press any key to exit

pause>nul

exit