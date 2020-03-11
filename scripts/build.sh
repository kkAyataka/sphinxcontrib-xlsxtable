#!/bin/bash

# variables
readonly PROJ_NAME=sphinxcontrib_xlsxtable


cd "`dirname "${0}"`"
cd ../

# cleanup
if [ -d ./build ]; then
  rm -r ./build
fi

if [ -d ./dist ]; then
  rm -r ./dist
fi

if [ -d ./${PROJ_NAME}.egg-info]; then
  rm -r ./${PROJ_NAME}.egg-info
fi

# build
python setup.py sdist bdist_wheel
