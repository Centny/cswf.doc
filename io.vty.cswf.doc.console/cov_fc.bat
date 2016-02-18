@echo off
set cnt=0
for %%A in (%1) do echo "%%~nA"
echo File count = %cnt%