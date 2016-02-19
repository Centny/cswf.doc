@echo off
convert %1 -resize %4 %2
for %%A in (%2) do set fn=%%~nA
for %%A in (%fn%-*) do echo %A%