#!/bin/bash
set -e
export PATH=`pwd`:`dirname ${0}`:$PATH
rm -rf out tmp

# echo Testing pdf2jpg_c...
# cov_pdf2jpg_c.sh test/xx.pdf test/pdf 60 no
# echo
# echo
# echo

echo Testing doc2jpg...
cov_doc2jpg.sh test/xx.docx test/doc 480 rm "http://127.0.0.1:8090/echo?tid=xa&process={0}" tmp out
echo
echo
echo

echo Testing pdf2jpg...
cov_pdf2jpg.sh test/xx.pdf test/pdf 60 no "http://127.0.0.1:8090/echo?tid=xb&process={0}" tmp out
echo
echo
echo

echo Testing ppt2jpg...
cov_ppt2jpg.sh test/xx.pptx test/ppt 480 rm "http://127.0.0.1:8090/echo?tid=xc&process={0}" tmp out
echo
echo
echo

echo Testing xls2jpg...
cov_xls2jpg.sh test/xx.xlsx test/xls 80 "rm" "http://127.0.0.1:8090/echo?tid=xd&process={0}" tmp out
echo
echo
echo

