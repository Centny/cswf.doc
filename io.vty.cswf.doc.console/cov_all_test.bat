@echo off

del /S /Q /f out
rmdir out
mkdir out

echo Testing doc2jpg...
call cov_doc2jpg.bat test\xx.docx out\docx 480 rm "http://127.0.0.1:8090/echo?tid=xa&process={0}"
echo .
echo .
echo .

echo Testing pdf2jpg...
call cov_pdf2jpg.bat test\xx.pdf out\pdf 60 " " "http://127.0.0.1:8090/echo?tid=xb&process={0}"
echo .
echo .
echo .

echo Testing ppt2jpg...
call cov_ppt2jpg.bat test\xx.pptx out\ppt 480 rm "http://127.0.0.1:8090/echo?tid=xc&process={0}"
echo .
echo .
echo .

echo Testing xls2jpg...
call cov_xls2jpg.bat test\xx.xlsx out\xls 80 rm "http://127.0.0.1:8090/echo?tid=xd&process={0}"
echo .
echo .
echo .

