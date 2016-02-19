@echo off

del /S /Q /f out
rmdir out
mkdir out

echo Testing doc2jpg...
call cov_doc2jpg.bat xx.docx out\docx 480 rm
echo .
echo .
echo .

echo Testing pdf2jpg...
call cov_pdf2jpg.bat xx.pdf out\pdf 60
echo .
echo .
echo .

echo Testing ppt2jpg...
call cov_ppt2jpg.bat xx.pptx out\ppt 480 rm
echo .
echo .
echo .

echo Testing xls2jpg...
call cov_xls2jpg.bat xx.xlsx out\xls 80 rm
echo .
echo .
echo .

