#!/bin/bash
set -e
export PATH=`pwd`:`dirname ${0}`:$PATH
cswf-doc -l -o $2.json -exe_c bash -exe_f "-c cov_png2jpg.sh $2-{0}.jpg $3 $4" -exe_p $5 -p $1 $2-{0}.png
echo
echo ---------------------json-------------------------
echo
echo
echo [json]
cat $2.json
echo
echo [/json]
echo