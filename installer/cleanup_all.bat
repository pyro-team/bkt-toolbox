@echo off
%~d0
cd %~dp0
@echo on
..\bin\ipy.exe -m bkt_install cleanup --clear_cache --clear_config --clear_settings --clear_xml --clear_resiliency
pause