@echo off
%~d0
cd %~dp0
@echo on
ipy-2.7.9\ipy.exe -m bkt_install.build_stdlib
pause