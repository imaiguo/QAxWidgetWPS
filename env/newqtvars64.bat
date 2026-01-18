
@call "C:\Program Files\Microsoft Visual Studio\2022\Community\VC\Auxiliary\Build\vcvars64.bat" amd64

set QT_BUILDDIR=D:\devtools\Qt\qteverywhere.5.15.2
set LD_LIBRARY_PATH=%QT_BUILDDIR%\lib;%LD_LIBRARY_PATH%
set QT_QPA_PLATFORM_PLUGIN_PATH=%QT_BUILDDIR%\plugins\platforms
set QT_PLUGIN_PATH=%QT_BUILDDIR%\plugin
set PKG_CONFIG_PATH=%QT_BUILDDIR%\lib\pkgconfig;%PKG_CONFIG_PATH%
set Path=%QT_BUILDDIR%\bin;%QT_BUILDDIR%\gnuwin32\bin;D:\devtools\Strawberry\perl\bin;%Path%