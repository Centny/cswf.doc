#!/bin/bash
set -e
export PATH=`pwd`:`dirname ${0}`:$PATH
cswf-doc -l -o $2.json -exe_c bash -exe_f "-c cov_pdf2jpg_c.sh $6/$2 $3 $4" -exe_p $5 -x $1 $1
echo
echo ---------------------json-------------------------
echo
echo
echo [json]
cat $2.json
echo
echo [/json]
echo