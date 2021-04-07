#!/usr/bin/env bash

installdir="tools"  # directory under ${HOME} to install files in
rcfile=".bashrc"    # file within ${HOME} to link rcsubfiles in

scripts=(           # executable scripts
    "iniget.py"
)

rcsubfiles=(        # run command files
    ".gitconf"
)

sourcedir="$(pwd)/"
targetdir="${HOME}/${installdir}/"
targetrcf="${HOME}/${rcfile}"


appendifmissing() {
    if [[ -f "$1" ]]
    then
        grep -qxF "$2" "$1" || echo "$2" >> "$1"
        return 0
    else
        return 1
    fi
}

# Create tools installation directory
mkdir -p "${targetdir}"

# Copy the required files into install dir
allfiles=( "${scripts[@]}" "${rcsubfiles[@]}" )
for f in "${allfiles[@]}"
do
    cp "${sourcedir}${f}" "${targetdir}"
done

# Add entries for rc subfiles into main rc file
for f in "${rcsubfiles[@]}"
do
    appendifmissing "${targetrcf}" ". ~/\"${installdir}\"/\"${f}\""
done

# If necessary, check if path to installdir exists within main rc file
if [[ "${#scripts[@]}" -ne "0" ]]
then
    if ! ( echo ":${PATH}:" | grep -qF ":${targetdir}/:" )
    then
        appendifmissing "${targetrcf}" "export PATH=\"\${PATH}:${targetdir}\""
        export PATH="\${PATH}:${targetdir}"
    fi
fi
