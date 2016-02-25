@echo off
msbuild
mstest /testcontainer:io.vty.cswf.doc.test\bin\Debug\doc.test.dll /testsettings:io.vty.cswf.doc.testrunconfig /resultsfile:io.vty.cswf.doc.trx