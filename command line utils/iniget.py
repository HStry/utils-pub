#!/usr/bin/env python3
"""
iniget.py [file] [section] [key]...

tiny command line script to get values from ini files."""

from pathlib import Path
from configparser import ConfigParser, NoSectionError, NoOptionError

def errprint(e, exitcode=None):
    if isinstance(e, Exception):
        print("{}: {}".format(e.__class__.__name__, str(e)), file=sys.stderr)
    else:
        print(e, file=sys.stderr)
    if exitcode is not None:
        sys.exit(exitcode)

if __name__ == '__main__':
    import sys
    
    try:
        file, section, key = sys.argv[1:]
    except ValueError:
        args = sys.argv[1:]
        if not args:
            errprint(__doc__, 0)
        
        errprint("Too {} arguments provided"
                 .format(('few', 'many')[len(args) > 3]))
        errprint(__doc__, 9)
    
    file = Path(file).expanduser().absolute()
    if not file.exists():
        errprint("File '{}' does not exist".format(str(file)))
        errprint(__doc__, 9)
    
    ini = ConfigParser()
    try:
        with open(str(file), 'r') as f:
            ini.read_file(f)
    except Exception as e:
        errprint("file '{}' not formatted as expected, "
                 "unable to continue".format(str(file)))
        errprint(__doc__, 9)
    
    try:
        print(ini.get(section, key))
    except NoSectionError:
        errprint("Section '{}' not available".format(section), 1)
    except NoOptionError:
        errprint("Key '{}' not available".format(key), 1)
    
    sys.exit(0)
