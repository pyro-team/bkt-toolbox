@echo off
%~d0
cd %~dp0
@echo on
ipy-2.7.10\ipy.exe -m bkt_install uninstall
pause
