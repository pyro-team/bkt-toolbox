@echo off
%~d0
cd %~dp0
@echo on
..\bin\ipy.exe -m bkt_install -h
pause
..\bin\ipy.exe -m bkt_install install -h
pause
..\bin\ipy.exe -m bkt_install uninstall -h
pause
..\bin\ipy.exe -m bkt_install configure -h
pause
..\bin\ipy.exe -m bkt_install cleanup -h
pause