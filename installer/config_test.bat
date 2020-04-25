@echo off
%~d0
cd %~dp0
@echo on
ipy-2.7.9\ipy.exe -m bkt_install configure --add_folders "features\bkt_excel"
pause
ipy-2.7.9\ipy.exe -m bkt_install configure --add_folders "features\bkt_excel" "features\bkt_visio"
pause
ipy-2.7.9\ipy.exe -m bkt_install configure --add_folders "C:\test\features\does_not_exist"
pause
ipy-2.7.9\ipy.exe -m bkt_install configure --remove_folders "features\bkt_excel" "features\bkt_visio"
pause
ipy-2.7.9\ipy.exe -m bkt_install configure --set_config "show_exception" "True"
pause
ipy-2.7.9\ipy.exe -m bkt_install configure --set_config "show_exception" "False" --set_config "log_level" "WARNING"
pause