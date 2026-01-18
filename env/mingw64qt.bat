
@echo off

set Path=D:\devtools\Qt\Qt5.12.12\Tools\mingw730_64\bin;D:\devtools\cmake-3.26.0-rc1-windows-x86_64\bin;D:\devtools\gn;%Path%

set QT_BUILDDIR=D:\devtools\Qt\Qt5.12.12\5.12.12\mingw73_64
set LD_LIBRARY_PATH=%QT_BUILDDIR%\lib;%LD_LIBRARY_PATH%
set QT_QPA_PLATFORM_PLUGIN_PATH=%QT_BUILDDIR%\plugins\platforms
set QT_PLUGIN_PATH=%QT_BUILDDIR%\plugin
set PKG_CONFIG_PATH=%QT_BUILDDIR%\lib\pkgconfig;%PKG_CONFIG_PATH%
set Path=%QT_BUILDDIR%\bin;%QT_BUILDDIR%\gnuwin32\bin;D:\devtools\Strawberry\perl\bin;%Path%

echo "set mingw73_64 env for qt5.12.12 ok"