#!/bin/bash
set -e
export PATH=`pwd`:`dirname ${0}`:$PATH
mkdir -p `dirname $2`
convert -density $3 -quality 100 $1 $2_%d.jpg
ls -a $2_*.jpg
if [ "$4" == "rm" ];then
  rm -f $1
fi