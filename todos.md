X   csv schreiben
-   csv auslesen
-   fingerprint erstellen
-   fingerprint auswerten, ob es gleich teile gibt
-   gleiche teile in solid works schreiben


# Bauen der Exe: 
`pyinstaller .\createFingerprint.py`


# def createFingerprint(line):
#     return 

# lines = []

# with open("01_Holzliste.csv", "r") as f:
#     lines = f.readlines()
    
# lines = lines[1:]    


# # for line in lines:
# #     attributes = line.split(";")
    
# #     hash = hashlib.sha256(line.encode('utf-8')).hexdigest()
# #     print(hash)


# print(lines[0].split(";"))