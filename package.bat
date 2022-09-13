@echo off

echo pyinstaller packing

pushd %1 & for %%i in (.) do set curr=%%~ni

pyinstaller -w -F --clean --distpath %cd%/%curr% %cd%/src/%curr%.py

xcopy .\icon\ %curr%\icon\ /e

echo Press any key to exit

pause>nul

exit