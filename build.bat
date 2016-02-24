@echo off
del /Q /S build
mkdir build
mkdir build\cswf.doc
::msbuild io.vty.cswf.doc.sln /property:Configuration="Release" /t:clean /t:build
xcopy io.vty.cswf.srv\bin\Release\cswf.srv.exe* build\cswf.srv
cd build
7z a cswf.srv.zip cswf.srv
if not "%1"=="" (
 echo Upload package to fvm server %1
 fvm -u %1 cswf.srv 0.0.1 cswf.srv.zip
)
cd ..\