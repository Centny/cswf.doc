@echo off
cswf-doc -l -o %2.json -exe_c cov_pdf2jpg_c.bat -exe_f %2_{0} -exe_a "%3 %4" -exe_p %5 -e %1 %2-{0}.pdf
echo.
echo ---------------------json-------------------------
echo.
echo.
echo [json]
type %2.json
echo.
echo [/json]
echo.