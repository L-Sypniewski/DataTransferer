curl -s https://codecov.io/bash > codecov
chmod +x codecov
./codecov -f "coverage.opencover.xml" -t $CODECOV_TOKEN