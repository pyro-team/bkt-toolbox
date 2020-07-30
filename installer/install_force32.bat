@echo off
%~d0
cd %~dp0
@echo on
..\bin\ipy.exe -m bkt_install install --force_office_bitness 32
pause