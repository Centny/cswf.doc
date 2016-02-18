@echo off
convert %1 -resize %3 %2
for %%A in (%2) do set fn=%%~nA
set cnt=0
for %%A in (%fn%-*) do set /a cnt+=1
echo %cnt%