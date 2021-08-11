_cwd="$PWD"
cd $PWD
pip3 install .
python3 -m sphinx.cmd.build -b html "${_cwd}/builddocs" "${_cwd}/docs" -E