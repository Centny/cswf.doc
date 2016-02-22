@echo off
echo %5
cswf-doc -l -o %2.json -exe_c cov_png2jpg.bat -exe_f %2-{0}.jpg -exe_a "%3 %4" -exe_p %5 -w %1 %2-{0}.png