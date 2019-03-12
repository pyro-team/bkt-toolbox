%~d0
cd %~dp0
ipy-2.7.9\ipy.exe -m config --add_folder "features\bkt_excel"
ipy-2.7.9\ipy.exe -m config --add_folder "features\bkt_excel" --add_folder "features\bkt_visio"
ipy-2.7.9\ipy.exe -m config --add_folder "C:\test\features\does_not_exist"
pause