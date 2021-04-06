#!/usr/bin/env bash

installdir="tools"
rcfile=".bashrc"

scripts=(
    "iniget.py"
)

sourcefiles=(
#    ".tools"
)

mkdir -p "${HOME}/${installdir}"

for script in "${scripts[@]}"
do
    cp "./${script}" "${HOME}/${installdir}/"
done

for sourcefile in "${sourcefiles[@]}"
do
    cp "./${sourcefile}" "${HOME}/${installdir}/"
    
    grep -qxF '. "~/${installdir}/${sourcefile}"' "${HOME}/${rcfile}" || \
    echo '. "~/${installdir}/${sourcefile}"' >> "${HOME}/${rcfile}"
done

( echo "${PATH}" | grep -qF "${HOME}/${installdir}/" ) || \
echo 'export PATH=$PATH:"${HOME}/${installdir}/"' >> "${HOME}/${rcfile}"

