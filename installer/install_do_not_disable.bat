@echo off
%~d0
cd %~dp0
@echo on
..\bin\ipy.exe -m bkt_install install --add_to_dndlist
pause