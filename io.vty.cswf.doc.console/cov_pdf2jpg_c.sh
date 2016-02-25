#!/bin/bash
set -e
export PATH=`pwd`:`dirname ${0}`:$PATH
convert -density $4 -quality 100 $1 $2_%d.jpg
ls -l $2_*
if [ "$5" == "rm" ];then
  rm -f $1
fi