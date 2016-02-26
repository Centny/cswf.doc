#!/bin/bash
set -e
export PATH=`pwd`:`dirname ${0}`:$PATH
convert $1 -background white -flatten -resize $3 $2
if [ "$4" == "rm" ]; then
  rm -f $1
fi
echo $2