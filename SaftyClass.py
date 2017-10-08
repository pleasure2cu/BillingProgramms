import glob


def alleAdressenVorhanden(path, kundenListe):
    liste = glob.glob(path + '*.txt')
    for i in range(len(liste)):
        liste[i] = liste[i].split('\\')[-1].split('.')[0]

    for kunde in kundenListe:
        if not kunde in liste:
            raise ValueError("Der Kunde " + kunde + " hat keine Text Datei mit seiner Adresse.")
        

def pruefeDatum(datum):
    if datum == '':
        return
    datum = datum.split('.')
    if not len(datum) in [2,3]:
        raise ValueError("ungültiges Datum gegeben")

    if (not len(datum[0]) in [1,2]) or (int(datum[0])>31):
        raise ValueError("Ungültiger Tag im Datum gegeben")

    if (not len(datum[1]) in [1,2]) or (int(datum[1])>12):
        raise ValueError("Ungültiger Monat im Datum gegeben")

    if (len(datum) == 3) and (not len(datum[2]) in [1,2,3,4]):
        raise ValueError("Ungültiges Jahr im Datum gegeben")


def pruefeVerteilung(verteilung, optionen_anzahl):

    v = verteilung.split()

    if (len(v) == 1) and (optionen_anzahl == 1):
        if ',' in verteilung:
            raise ValueError("Bitte keine Kommas verwenden")
        int(verteilung)
        return
    else:
        if len(v)%2 != 0:
            raise ValueError("Sie haben eine ungerade Anzahl von Punkten eingegeben. Bitte zu jeder Menge eine Art angeben.")
        for i in range(len(v)):
            if i%2 == 0:
                int(v[i])
            else:
                if not int(v[i]) in range(optionen_anzahl):
                    raise ValueError("Es wurde eine unzulässige Art entdeckt.")
                
    
    
























