@echo off
%~d0
cd %~dp0
@echo on
..\bin\ipy.exe -m bkt_install configure --migrate_from "2.4"
pause