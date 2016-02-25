#!/bin/bash
set -e
export PATH=`pwd`:`dirname ${0}`:$PATH
convert $1 -background white -flatten -resize $4 $2
if [ "$5" == "rm"]; then
  rm -f $1
fi
echo $2