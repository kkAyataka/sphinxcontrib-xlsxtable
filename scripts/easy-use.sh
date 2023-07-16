#!/bin/bash

cd "`dirname "$0"`"

# Build
echo ""
./build.sh
if [[ $? != 0 ]]; then
    echo "build error" >&2
    exit 1
else
    echo -e "\033[0;32m========== Build: OK ==========\033[0m"
fi

# Install
echo ""
pip install --no-deps --force-reinstall ../dist/sphinxcontrib_xlsxtable-1.1.0-py3-none-any.whl
if [[ $? != 0 ]]; then
    echo "pip install error" >&2
    exit 1
else
    echo -e "\033[0;32m========== Install: OK ==========\033[0m"
fi

# Run Sphinx"
echo ""
make -C ../test clean
make -C ../test html
if [[ $? != 0 ]]; then
    echo "make error" >&2
    exit 1
else
    echo -e "\033[0;32m========== Sphinx: OK ==========\033[0m"
fi

# RUn cli
echo ""
python -m sphinxcontrib.xlsxtable --sheet=Sheet1 --header-rows=1 ../test/_res/sample.xlsx
if [[ $? != 0 ]]; then
    echo "cli error" >&2
    exit 1
else
    echo -e "\033[0;32m========== CLI: OK ==========\033[0m"
fi

# Run test
cd ../
python setup.py test
if [[ $? != 0 ]]; then
    echo -e "\033[31m========== Test: Error ==========\033[0m"
    exit 1
else
    echo -e "\033[0;32m========== Test: OK ==========\033[0m"
fi
