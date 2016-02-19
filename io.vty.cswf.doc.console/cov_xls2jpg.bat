@echo off
cswf-doc -l -o %2.json -exe_c cov_pdf2jpg_c.bat -exe_f %2_{0} -exe_a "%3 %4" -e %1 %2-{0}.pdf