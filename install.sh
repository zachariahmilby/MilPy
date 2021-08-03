_cwd="$PWD"
cd _cwd
pip3 install .
python3 -m sphinx.cmd.build -b html "${_cwd}/builddocs" "${_cwd}/docs" -E