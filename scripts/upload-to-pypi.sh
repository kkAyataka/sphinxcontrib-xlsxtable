#!/bin/bash

# checks env
if [ -z ${PYPI_USER} -o  -z ${PYPI_PASS} ]; then
  echo PYPI_USER or PYPI_PASS is not setuped
  exit 1
fi

cd "`dirname "${0}"`"
cd ../

# build
./scripts/build.sh

# upload
if [ -d ~/.pypirc ]; then
  twine upload -u ${PYPI_USER} -p ${PYPI_PASS} --repository-url https://test.pypi.org/legacy/ dist/*
else
  twine upload -u ${PYPI_USER} -p ${PYPI_PASS} dist/*
fi
