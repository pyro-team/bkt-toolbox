@echo off
%~d0
cd %~dp0
cd ..
@echo on

REM bin\ipy.exe -m unittest discover
REM pause

bin\ipy.exe -m unittest tests.test_taskpane
pause