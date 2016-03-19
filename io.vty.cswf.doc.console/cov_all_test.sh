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
cswf-doc -w -l test/xx.docx test/doc-{0}.jpg
echo
echo
echo

echo Testing doc2jpg1...
cswf-doc -w -l test/xx1.doc test/doc1-{0}.jpg
echo
echo
echo

echo Testing doc2jpg2...
cswf-doc -w -l test/xx2.doc test/doc2-{0}.jpg
echo
echo
echo

echo Testing pdf2jpg...
cswf-doc -pdf -l test/xx.pdf test/pdf-{0}.jpg
echo
echo
echo

echo Testing ppt2jpg...
cswf-doc -p -l test/xx.pptx test/pptx-{0}.jpg
echo
echo
echo

echo Testing ppt2jpg2...
cswf-doc -p -l test/xx1.ppt test/ppt1-{0}.jpg
echo
echo
echo

echo Testing xls2jpg...
cswf-doc -e -l test/xx.xlsx test/xlsx-{0}.jpg
echo
echo
echo


echo Testing xls2jpg2...
cswf-doc -e -l test/xx1.xls test/xls1-{0}.jpg
echo
echo
echo

echo Testing xls2jpg3...
cswf-doc -e -l test/xx2.xls test/xls2-{0}.jpg
echo
echo
echo

echo Testing xls2jpg4...
#cswf-doc -e -l test/xx3.xls test/xls3-{0}.jpg
echo
echo
echo
