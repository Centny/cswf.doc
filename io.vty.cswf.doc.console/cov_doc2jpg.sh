#!/bin/bash
set -e
tpwd=`pwd`
export PATH=$tpwd:`dirname ${0}`:$PATH
mkdir -p $6
mkdir -p $7
#
cd $6
cswf-doc -l -o $2.json -exe_c bash -exe_f "-c $2-{0}.jpg $3 $4" -exe_p $5 -w $1 $2-{0}.png
cd $tpwd
#
if [ "$6" != "$7" ];then
 cp -rf $6/* $7/
fi
echo
echo ---------------------json-------------------------
echo
echo
echo [json]
cat $2.json
echo
echo [/json]
echo