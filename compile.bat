@echo off
call "C:\Program Files\Microsoft Visual Studio\2022\Professional\Common7\Tools\VsDevCmd.bat" -arch=x64
cd ThreadSafeC
cl.exe /I"SDK\include" /c ThreadSafeC.cpp
link.exe /DLL /DEF:ThreadSafeC.def ThreadSafeC.obj SDK\lib\x64\XLCALL32.LIB SDK\lib\x64\frmwrk32.lib /OUT:ThreadSafeC.xll