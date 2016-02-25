#!/bin/bash
set -e
export PATH=`pwd`:`dirname ${0}`:$PATH
cd ws
rm -rf out
mkdir out

echo Testing doc2jpg...
cov_doc2jpg.sh test/xx.docx docx 480 rm "http://127.0.0.1:8090/echo?tid=xa&process={0}" tmp out
echo
echo
echo

sdfs
echo Testing pdf2jpg...
cov_pdf2jpg.sh test/xx.pdf pdf 60 no "http://127.0.0.1:8090/echo?tid=xb&process={0}" tmp out
echo
echo
echo

echo Testing ppt2jpg...
cov_ppt2jpg.sh test/xx.pptx ppt 480 rm "http://127.0.0.1:8090/echo?tid=xc&process={0}" tmp out
echo
echo
echo

echo Testing xls2jpg...
cov_xls2jpg.sh test/xx.xlsx xls 80 rm "http://127.0.0.1:8090/echo?tid=xd&process={0}" tmp out
echo
echo
echo

