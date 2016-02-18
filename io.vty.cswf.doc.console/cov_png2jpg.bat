@echo off
convert %1 -background white -flatten -resize %4 %2
if "%5" == "rm" (
  del %1
)
echo %2