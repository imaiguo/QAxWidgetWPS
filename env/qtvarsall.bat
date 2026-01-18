
@echo off


@call "C:\Program Files\Microsoft Visual Studio\2022\Community\VC\Auxiliary\Build\vcvars64.bat" amd64


@if  /I "%~nx1%"=="amd64" (
    set "LIB=D:\devtools\Qt\Qt5.12.12\5.12.12\msvc2017_64\bin;C:\Program Files\OpenSSL-Win64;%LIB%"
    set "Path=%Path%;D:\devtools\Qt\Qt5.12.12\5.12.12\msvc2017_64\bin"
    echo "set msvc2017_64 env for qt ok"
    exit /B 0
)


@if  /I "%~nx1%"=="x86" (
    set "LIB=D:\devtools\Qt\Qt5.12.12\5.12.12\msvc2017\bin;%LIB%"
    set "Path=%Path%;D:\devtools\Qt\Qt5.12.12\5.12.12\msvc2017\bin"
    echo "set msvc2017 env for qt ok"
    exit /B 0
)

echo "set qt env failed, usage: qtvarsall.bat amd64/x86"