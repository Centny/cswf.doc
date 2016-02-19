@echo off
convert -density %4 -quality 100 %1 %2_%%d.jpg
for %%A in (%2_*) do echo %%A%
if "%5" == "rm" (
  del %1
)