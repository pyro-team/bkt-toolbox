# -*- coding: utf-8 -*-



import os, glob, subprocess

#Build StdLib.DLL
ipath = 'ipy-2.7.9'
ipyc  = ipath + '\\ipyc.exe'

# any library files you need
gb1 = glob.glob(ipath + "\\Lib\\*.py")
gb2 = glob.glob(ipath + "\\Lib\\*\\*.py")
gb3 = glob.glob(ipath + "\\Lib\\*\\*\\*.py")
gb = list(set(gb1 + gb2 + gb3))

gb = [ipyc,"/main:bkt_install\\StdLib.py","/embed","/platform:x86","/target:dll"] + gb
subprocess.call(gb)

print("Made StdLib")