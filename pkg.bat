@echo off
del /Q /S build
mkdir build
mkdir build\cswf.doc
mkdir build\cswf.doc\test
::msbuild io.vty.cswf.doc.sln /property:Configuration="Release" /t:clean /t:build
xcopy io.vty.cswf.doc.console\bin\Release\cov_*.bat  build\cswf.doc
xcopy io.vty.cswf.doc.console\bin\Release\cswf-doc.exe*  build\cswf.doc
xcopy io.vty.cswf.doc.console\bin\Release\*.dll build\cswf.doc
xcopy io.vty.cswf.doc.test\test\* build\cswf.doc\test

cd build
7z a cswf.doc.zip cswf.doc
if not "%1"=="" (
 echo Upload package to fvm server %1
 fvm -u %1 cswf.doc 0.0.1 cswf.doc.zip
)
cd ..\