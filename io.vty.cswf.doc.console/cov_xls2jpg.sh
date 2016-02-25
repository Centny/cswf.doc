#!/bin/bash
set -e
export PATH=`pwd`:`dirname ${0}`:$PATH
cswf-doc -l -o $2.json -exe_c bash -exe_f "-c cov_pdf2jpg_c.sh $2_{0} $3 $4" -exe_p $5 -e $1 $2-{0}.pdf
echo
echo ---------------------json-------------------------
echo
echo
echo [json]
cat $2.json
echo
echo [/json]
echo