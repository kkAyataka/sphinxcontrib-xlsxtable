#!/bin/bash

# checks env
if [ -z ${PYPI_USER} -o  -z ${PYPI_PASS} ]; then
  echo PYPI_USER or PYPI_PASS is not setuped
  exit 1
fi

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

# upload
if [ -d ~/.pypirc ]; then
  twine upload -u ${PYPI_USER} -p ${PYPI_PASS} --repository-url https://test.pypi.org/legacy/ dist/*
else
  twine upload -u ${PYPI_USER} -p ${PYPI_PASS} dist/*
fi
