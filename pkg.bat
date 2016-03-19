@echo off
del /Q /S build
mkdir build
mkdir build\cswf.doc
mkdir build\cswf.doc\sdata_i
mkdir build\cswf.doc\sdata_i\test
call VsMSBuildCmd
msbuild io.vty.cswf.doc.sln /property:Configuration="Release" /p:Platform="x64" /t:clean /t:build
if not "%errorlevel%"=="0" goto :efail
xcopy io.vty.cswf.doc.console\bin\x64\Release\cov_*.sh  build\cswf.doc
xcopy io.vty.cswf.doc.console\bin\x64\Release\cswf-doc.exe*  build\cswf.doc
xcopy io.vty.cswf.doc.console\bin\x64\Release\*.dll build\cswf.doc
xcopy io.vty.cswf.doc.test\test\* build\cswf.doc\sdata_i\test

cd build
zip -r cswf.doc.zip cswf.doc
if not "%1"=="" (
 echo Upload package to fvm server %1
 fvm -u %1 cswf.doc 0.0.1 cswf.doc.zip
)
cd ..\
goto :esuccess

:efail
echo "Build fail"
exit 1

:esuccess
echo "Build success"