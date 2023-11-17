from cx_Freeze import setup, Executable

base = None    

executables = [Executable("createFingerprint.py", base=base)]

packages = ["idna"]
options = {
    'build_exe': {    
        'packages':packages,
    },    
}

setup(
    name = "SW-Fingerprint-Generator",
    options = options,
    version = "1.0.0",
    description = 'Fingerabdruckgenerator und Gleichteilnummergenerator für SW-Teile',
    executables = executables
)