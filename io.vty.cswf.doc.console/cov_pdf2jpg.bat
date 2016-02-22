@echo off
cswf-doc -l -o %2.json -exe_c cov_pdf2jpg_c.bat -exe_f %2 -exe_a "%3 %4" -exe_p %5 -x %1 %1