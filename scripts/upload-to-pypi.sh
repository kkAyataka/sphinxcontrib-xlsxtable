#!/bin/bash

cd "`dirname "${0}"`"
cd ../

# load env file
if [ -e ./scripts/env ]; then
  source ./scripts/env
fi

# checks env
if [ -z ${PYPI_USER} -o  -z ${PYPI_PASS} ]; then
  echo PYPI_USER or PYPI_PASS is not setuped
  exit 1
fi

# build
./scripts/build.sh

# upload
if [ ${PYPI_PRODUCTION} == 1 ]; then
  twine upload -u ${PYPI_USER} -p ${PYPI_PASS} dist/*
else
  twine upload -u ${PYPI_USER} -p ${PYPI_TEST_PASS} --repository-url https://test.pypi.org/legacy/ dist/*
fi
