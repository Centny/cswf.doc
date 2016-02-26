#!/bin/bash
set -e
twd=`pwd`
export PATH=$twd:`dirname ${0}`:$PATH
out_n=$2
#tmp
tmp_w=$6/$out_n
tmp_j=$tmp_w/out.json
tmp_f=$tmp_w/ws/$out_n
#out
out_w=$7
out_f=`dirname $out_w/$out_n`
#run converter
mkdir -p `dirname $tmp_f`
cswf-doc -l -o $tmp_j -exe_c "bash" -exe_f " -c 'cov_pdf2jpg_c.sh {0} $tmp_f $3 $4'" -exe_p $5 -prefix $tmp_w/ws/ -x $1 $1
#copy file to out
mkdir -p $out_f
if [ "$6" != "$7" ];then
 cp -rf $tmp_f* $out_f/
fi
#print result
echo
echo ---------------------result-------------------------
echo
echo
echo [json]
cat $tmp_j
echo
echo [/json]
echo
#clear
rm -rf $tmp_w
