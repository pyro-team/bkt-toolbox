
:: FIND MSBUILD.EXE in C:\Windows\Microsoft.Net\Framework
:: LOOK FOR FRAMEWORK VERSION 4.5, 4.0 or 3.5
:: see http://stackoverflow.com/a/17736442, http://stackoverflow.com/a/23028195
@echo off
set "framework_path="
for /d %%a in ("C:\Windows\Microsoft.Net\Framework\*") do if "%%~nxa"=="v4.0.30319" (set "framework_path=%%a")
if "%framework_path%"=="" (
	for /d %%a in ("C:\Windows\Microsoft.Net\Framework\*") do if "%%~nxa"=="v4.0" (set "framework_path=%%a")
)
if "%framework_path%"=="" (
	for /d %%a in ("C:\Windows\Microsoft.Net\Framework\*") do if "%%~nxa"=="v3.5" (set "framework_path=%%a")
)


:: RUN MSBUILD
@echo on
%framework_path%\msbuild bkt.sln /t:Rebuild /p:Configuration=Release "/p:Platform=Any cpu"


:: COPY FILES
copy bkt-addin\bin\Release\*.* ..\bin
copy bkt-dev-addin\bin\Release\BKT.Dev.* ..\bin

