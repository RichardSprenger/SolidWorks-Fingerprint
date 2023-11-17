import hashlib
import csv
import argparse

# Input und Output Datei aus Kommandozeile einlesen
parser = argparse.ArgumentParser(description='Erstelle Fingerabdruck und Gleichteilnummer aus Import CSV Datei')
parser.add_argument('-i', '--importfilename', dest='importfilename', metavar='importfilename', type=str, help='Absoluter Pfad der Importdatei')
parser.add_argument('-o', '--outputfilename', dest='outputfilename', metavar='outputfilename', type=str, help='Absoluter Pfad der Outputdatei')

args = parser.parse_args()

# Output Datei schreiben
def writeFile(filename, list):
    # Output öffnen
    with open(filename, 'w', newline='') as csvfile:
        # Alle Attribute aus der Liste auslesen
        fieldnames = list[0].keys()
        
        # CSV Writer erstellen
        writer = csv.DictWriter(csvfile, fieldnames=fieldnames, delimiter=';')

        # Header schreiben
        writer.writeheader()
        # Alle Zeilen schreiben
        for row in list:
            writer.writerow(row)

# Attribute der CSV Datei, welche für die Fingerprinterstellung genutzt werden
IDENTIFIZIERUNGSATTRIBUTE = {
    "Laenge",
    "Breite",
    "Dicke",
    "Gewicht",
    "Material",
    "SW-MX",
    "SW-MY",
    "SW-MZ",
    "Px",
    "Py",
    "Pz"
}

# Input Datei einlesen
items = []

reader = csv.DictReader(open(args.importfilename, 'r'), delimiter=';')

for row in reader:
    # Fingerprint erstellen -> Alle Attribute aneinanderhängen und hashen
    string = ""
    for att in IDENTIFIZIERUNGSATTRIBUTE:
        string += row[att]
    
    # Hash erstellen
    hash = hashlib.sha256(string.encode('utf-8')).hexdigest()
    
    # Hash speichern
    row['fingerprint'] = hash
    items.append(row)

# Gleichteilnummer herausfinden
results = {}
newItems = []
gleichteilnummer = 0

for key, value in enumerate(items):
    # Schauen ob es das teil schon einmal gibt
    if value['fingerprint'] not in results.values():
        # wenn nicht erst die Gleichteilnummer erhöhen und den Fingerprint speichern
        gleichteilnummer += 1
        results[gleichteilnummer] = value['fingerprint']

    # In jedem Fall wird die Gleichteilnummer zu dem Teil gespeichert
    # Suche die Gleichteilnummer zu dem Fingerprint und speichere diese zu dem Teil
    _tempNr = [x for x, y in results.items() if y == value['fingerprint']][0]
    if (_tempNr < 10):
        value['GltNummer'] = '0' + str(_tempNr)
    else:
        value['GltNummer'] = _tempNr
    
    # Speichere das Teil in der neuen Liste
    newItems.append(value)

# Output Datei schreiben
writeFile(args.outputfilename, newItems)