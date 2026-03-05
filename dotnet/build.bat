
:: FIND MSBUILD.EXE
:: First try Visual Studio 2022/2019 MSBuild (supports modern C#)
:: Then fall back to .NET Framework MSBuild
@echo off
set "msbuild_path="

:: Try VS 2022 (various editions)
for %%e in (Enterprise Professional Community BuildTools) do (
    if "%msbuild_path%"=="" if exist "C:\Program Files\Microsoft Visual Studio\2022\%%e\MSBuild\Current\Bin\MSBuild.exe" (
        set "msbuild_path=C:\Program Files\Microsoft Visual Studio\2022\%%e\MSBuild\Current\Bin\MSBuild.exe"
    )
)

:: Try VS 2022 (various editions)
for %%e in (Enterprise Professional Community BuildTools) do (
    if "%msbuild_path%"=="" if exist "C:\Program Files (x86)\Microsoft Visual Studio\2022\%%e\MSBuild\Current\Bin\MSBuild.exe" (
        set "msbuild_path=C:\Program Files (x86)\Microsoft Visual Studio\2022\%%e\MSBuild\Current\Bin\MSBuild.exe"
    )
)

:: Try VS 2019 (various editions)
for %%e in (Enterprise Professional Community BuildTools) do (
    if "%msbuild_path%"=="" if exist "C:\Program Files (x86)\Microsoft Visual Studio\2019\%%e\MSBuild\Current\Bin\MSBuild.exe" (
        set "msbuild_path=C:\Program Files (x86)\Microsoft Visual Studio\2019\%%e\MSBuild\Current\Bin\MSBuild.exe"
    )
)

:: Fall back to .NET Framework MSBuild (older C# compiler)
if "%msbuild_path%"=="" (
    for /d %%a in ("C:\Windows\Microsoft.Net\Framework\*") do if "%%~nxa"=="v4.0.30319" (set "msbuild_path=%%a\msbuild.exe")
)

if "%msbuild_path%"=="" (
    echo ERROR: MSBuild not found!
    pause
    exit /b 1
)

echo Using MSBuild: %msbuild_path%


:: RUN MSBUILD
@echo on
"%msbuild_path%" bkt.sln /t:Rebuild /p:Configuration=Release "/p:Platform=Any cpu"


:: COPY FILES
copy bkt-addin\bin\Release\*.* ..\bin
copy bkt-dev-addin\bin\Release\BKT.Dev.* ..\bin

