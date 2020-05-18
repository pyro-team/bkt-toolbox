@echo off
%~d0
cd %~dp0
cd ..
@echo on

bin\ipy.exe -m unittest discover -v -s tests
pause