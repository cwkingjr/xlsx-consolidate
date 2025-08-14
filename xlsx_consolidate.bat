@echo off
set app=xlsx_consolidate.exe
echo Invoking %app% "%1"
echo Please wait for the tool's results to be printed to the terminal.
%HOMEDRIVE%%HOMEPATH%\.local\bin\%app% --directory "%1"
pause