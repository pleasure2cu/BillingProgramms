import xlsxwriter
import time
import datetime
import SaftyClass


input_path = ''
output_path = ''

def getList(path):
    "Liest die Reihenfolge.txt Datei und gibt eine Liste mit den enthaltenen Namen zurück"
    
    output = []
    file = path + '\\Reihenfolge.txt'
    with open(file, 'r') as file:
        while True:
            line = file.readline().rstrip()
            if (line == ""):
                break
            output.append(line)

    return output

def setFonts(wb):
    "erstellt alle notwendigen Fonts für die Kästchen und gibt sie in einer Liste zurück"
    
    normal = wb.add_format({'font_name':'Arial Narrow', 'font_size':12})
    normal_rechts = wb.add_format({'font_name':'Arial Narrow', 'font_size':12, 'align':'right'})
    tabelle_text = wb.add_format({'font_name':'Arial Narrow', 'font_size':12, 'border':1})
    tabelle_text_doppelt_unterstrichen = wb.add_format({'font_name':'Arial Narrow', 'font_size':12, 'border':1})
    tabelle_text_doppelt_unterstrichen.set_bottom(6)
    tabelle_float = wb.add_format({'font_name':'Arial Narrow', 'font_size':12, 'border':1, 'num_format':'0.00'})
    tabelle_float_doppelt_unterstrichen = wb.add_format({'font_name':'Arial Narrow', 'font_size':12, 'border':1, 'num_format':'0.00'})
    tabelle_float_doppelt_unterstrichen.set_bottom(6)
    gross = wb.add_format({'font_name':'Arial Narrow', 'font_size':18})
    gross_unterstrichen = wb.add_format({'font_name':'Arial Narrow', 'font_size':18, 'bottom':2})
    gross_rechts = wb.add_format({'font_name':'Arial Narrow', 'font_size':18, 'align':'right'})
    gross_rechts_unterstrichen = wb.add_format({'font_name':'Arial Narrow', 'font_size':18, 'align':'right', 'bottom':2})
    return {'normal':normal,
            'normal_rechts':normal_rechts,
            'tabelle_text':tabelle_text,
            'tabelle_text_doppelt_unterstrichen':tabelle_text_doppelt_unterstrichen,
            'tabelle_float':tabelle_float,
            'tabelle_float_doppelt_unterstrichen':tabelle_float_doppelt_unterstrichen,
            'gross':gross,
            'gross_unterstrichen':gross_unterstrichen,
            'gross_rechts':gross_rechts,
            'gross_rechts_unterstrichen':gross_rechts_unterstrichen}
    

def getTitle(output_path, rechnungsnummer, kunde, anhaengsel):
    """erstellt den Titel path für die Excel Datei, wichtig ist dass die Rechnungsnummer noch mit Nullen aufgefüllt wird,
       damit man nicht eine 5 bei den 500er Zahlen hat, wenn man die Dateien sortiert"""
    
    if rechnungsnummer < 10:
        string = '00'+str(rechnungsnummer)
    elif rechnungsnummer < 100:
        string = '0'+str(rechnungsnummer)
    else:
        string = str(rechnungsnummer)
        
    return output_path + string + " Eier " + kunde + " " + anhaengsel + ".xlsx"


    
def getInformation(path, kunde):
    """Sucht im gegebenen Ordner (path) eine Textdatei mit dem Namen Kunde und liest aus dieser
    den Namen, die Adresse und die PLZ-Ort des Kunden und gibt sie in dieser Reihenfolge als Liste zurück"""
    output = []
    path = path + '\\' + kunde + '.txt'
    with open(path, 'r') as file:
        for i in range(3):
            output.append(file.readline().rstrip())

    return output;

def setColumnWidths(sheet):
    "setzt die breiten der Kolonnen"
    sheet.set_column('A:A', 12)
    sheet.set_column('B:B', 8)
    sheet.set_column('C:C', 28)
    sheet.set_column('D:D', 15)
    sheet.set_column('E:E', 15)


def fillHeader(worksheet, info, font):
    "schreibt den header inklusive dem Kunden Bereich"
    worksheet.write('A1', 'Name des Betriebs', font['gross'])
    worksheet.write('E1', 'Strasse', font['gross_rechts'])
    worksheet.write('E2', 'PLZ und Ort', font['gross_rechts_unterstrichen'])
    worksheet.write('E4', 'E-Mail: muster@example.com', font['normal_rechts'])
    worksheet.write('A4', 'Telefon: 012 345 76 89', font['normal'])

    for i in range(4):
        worksheet.write(1, i, '', font['gross_unterstrichen'])

    worksheet.write('A8', 'Kunde:', font['normal'])
    worksheet.write('D8', info[0], font['normal'])
    worksheet.write('D9', info[1], font['normal'])
    worksheet.write('D10', info[2], font['normal'])


def fillRechnungDatum(worksheet, rechnungsnummer, periode, font):
    worksheet.write('A12', 'Rechnung', font['gross_unterstrichen'])

    for i in range(1,5):
        worksheet.write(11, i, '', font['gross_unterstrichen'])

    worksheet.write('A15', 'Datum:', font['normal'])
    t = time.strftime("%d.%m.%Y")
    worksheet.write('D15', t, font['normal'])
    worksheet.write('A16', 'Rechnungperiode:', font['normal'])
    worksheet.write('D16', periode, font['normal'])
    worksheet.write('A17', 'Rechnungsnummer:', font['normal'])
    worksheet.write('D17', str(rechnungsnummer), font['normal'])

def startTable(sheet, font):
    sheet.write('A19', 'Datum', font['normal'])
    sheet.write('B19', 'Anzahl', font['normal'])
    sheet.write('C19', 'Beschreibung', font['normal'])
    sheet.write('D19', 'Preis', font['normal'])
    sheet.write('E19', 'Betrag', font['normal'])

def tagSchoenen(tag):
    "Nimmt einen String Tag entgegen und brint ihn in das Format dd und gibt ihn als String (inkl. den Punkt) zurück (kann auch für Monat verwendet werden)"
    
    if len(tag) == 1:
        return '0'+tag+'.'
    else:
        return tag+'.'


def jahrSchoenen(jahr=''):
    "gibt eine schöne Variante des Jahrs in jahr als String zurück"
    

    if len(jahr) == 0:
        return time.strftime("%Y")
    elif len(jahr) == 1:
        return "200"+jahr
    elif len(jahr) == 2:
        return "20"+jahr
    else:
        return jahr
    
def getDatum(datum):
    "Diese Funktion verändert das Datum so, dass es im Format dd.mm.jjjj ist"
    datum = datum.strip().split('.')

    tag = tagSchoenen(datum[0])
    monat = tagSchoenen(datum[1])
    if len(datum) == 2:
        jahr = jahrSchoenen()
    else:
        jahr = jahrSchoenen(datum[2])

    return tag + monat + jahr


def printKundenOptionen(optionen, kunde):
    if len(optionen) == 1:
        print ("Bei " + kunde + " gilt nur '" + optionen[0] + " Rp.'\n")
    else:
        print ("Für " + kunde + " gilt die folgende Optionen Tabelle: ")
        for i in range(len(optionen)):
            print(str(i) + " = " + optionen[i] + " Rp.")
        print('')
    

def getPriceDescriptionOptions(path, kunde):
    """ diese Funktion liest aus der Kunden Text Datei:
        - alle Zeilen mit Beschreibungen und Preisen
        - produziert dann eine liste mit allen Möglichkeiten
        - gibt diese Liste zurück """
    
    path = path + '\\' + kunde + '.txt'
    output = []

    with open(path, 'r') as file:
        # die ersten 3 Zeilen interessieren nicht
        for i in range(3):
            file.readline()

        while True:
            line = file.readline().rstrip()
            if line == '':
                break
            
            info = line.split(',')

            for i in range(len(info)-1):
                output.append(info[0] + ',' + info[i+1])

    if len(output) == 0:
        raise ValueError("Der Kunde "+kunde+" hat keine Preisinformationen gespeichert.")
    
    return output

def writeLine(zeile, datum, anzahl, beschreibung, sheet1, sheet2, font1, font2):
    "schreibt eine Zeile in die Tabelle, je von sheet1 und sheet2 mit den passenden Formaten. Gibt die Kosten zurück, die in dieser Zeile eingetragen werden"
    sheet1.write(zeile, 0, datum, font1['tabelle_text'])
    sheet1.write(zeile, 1, int(anzahl), font1['tabelle_text'])
    sheet1.write(zeile, 2, beschreibung.split(', ')[0], font1['tabelle_text'])
    sheet1.write(zeile, 3, float('0.'+beschreibung.split(', ')[1]), font1['tabelle_float'])
    sheet1.write(zeile, 4, int(anzahl)*float(beschreibung.split(', ')[1])/100, font1['tabelle_float'])

    sheet2.write(zeile, 0, datum, font2['tabelle_text'])
    sheet2.write(zeile, 1, int(anzahl), font2['tabelle_text'])
    sheet2.write(zeile, 2, beschreibung.split(', ')[0], font2['tabelle_text'])
    sheet2.write(zeile, 3, float('0.'+beschreibung.split(', ')[1]), font2['tabelle_float'])
    sheet2.write(zeile, 4, int(anzahl)*float(beschreibung.split(', ')[1])/100, font2['tabelle_float'])
    
    return int(anzahl)*float(beschreibung.split(', ')[1])/100


def addLine(zeile, sheet1, sheet2, datum, optionen, font1, font2):
    """fragt nach den der Menge an Eiern die an diesem Tag geliefert wurden. Falls es mehr als eine Option gibt, gibt der nutzer direkt deren Zusammensetztung ein.
       Danach werden die angegebe(n) Menge(n) in Aufgtrag gegeben in die Tabelle zu schreiben. Zurück gegeben wird eine Liste, deren erstes Element gleich der Anzahl
       von Zeilen ist, die in der Tabelle gefüllt wurden. Das zweite Element sagt, wie viel Wert alle Lieferungen zusammen haben, die in dieser Runde von addLine()
       hinzugefügt wurden."""

    menge = 0
    while True:
        try:
            if len(optionen) == 1:
                mengen = input('Wie viele Eier wurden am ' + datum + ' geliefert? ')
            else:
                mengen = input('Wie viele Eier wurden am ' + datum + ' geliefert (anzahl art ...)? ')
            SaftyClass.pruefeVerteilung(mengen, len(optionen))
            break
        except ValueError as e:
            print(e.args[0])
            print("Versuchen Sie es nocheinmal.")
        
    
    offset = 0
    kosten = 0
    
    if len(optionen) == 1:
        kosten += writeLine(zeile, datum, mengen, optionen[0], sheet1, sheet2, font1, font2)
        offset = 1
    else:
        verteilung = mengen.split(' ')

        kosten += writeLine(zeile, datum, verteilung[0], optionen[int(verteilung[1])], sheet1, sheet2, font1, font2)
        offset += 1
        for i in range(2, len(verteilung), 2):
            kosten += writeLine(zeile+offset, '', verteilung[i], optionen[int(verteilung[i+1])], sheet1, sheet2, font1, font2)
            offset += 1

    return [offset, kosten]
    
def printRest(zeile, sheet, font):
    sheet.write(zeile, 0, 'Konditionen:', font['normal'])
    zeile += 1
    sheet.write(zeile, 0, 'netto zahlbar bis ' +(datetime.date.today()+datetime.timedelta(365/12)).strftime('%d.%m.%Y'), font['normal'])
    zeile += 2
    sheet.write(zeile, 0, 'Herzlichen Dank für den geschätzten Auftrag', font['normal'])
    zeile += 2
    sheet.write(zeile, 0, 'Mit freundlichen Grüssen', font['normal'])
    zeile += 1
    sheet.write(zeile, 0, 'Name des Betriebs', font['normal'])



def updateRechnungsnummernDatei(sheet, rechnungsnummer, kunde, betrag, zeile):
    "Schreibt die Rechnungsnummer, Beschreibung, Kundenname, Rechnungsbetrag und das heutige Datum in die Rechnungsnummer Tabelle"
    sheet.write(zeile, 0, rechnungsnummer)
    sheet.write(zeile, 1, 'Eier')
    sheet.write(zeile, 2, kunde)
    sheet.write(zeile, 3, betrag)
    sheet.write(zeile, 4, time.strftime("%d.%m.%Y"))




    



# Datei mit den Rechnungsnummern erstellen
rechnungsnummern = xlsxwriter.Workbook(output_path + 'Rechnungsnummern_Liste.xlsx')
sheet_rn = rechnungsnummern.add_worksheet()
sheet_rn.write('A1', 'Nr.')
sheet_rn.write('C1', 'Name')
sheet_rn.write('D1', 'Betrag')
sheet_rn.write('E1', 'Rechnungsdatum')

# Datei zum Drucken erstellen
zum_drucken = xlsxwriter.Workbook(output_path + 'zum drucken.xlsx')
font2 = setFonts(zum_drucken)



periode = input("Rechnungsperiode (steht auf den Rechnungen): ")
basis = int(input("Erste Rechnungsnummer die verwendet werden darf: "))
anhaengsel = input("Was soll am Schluss der Excel-Datei Namen stehen (Bspiele: '2-16', 'Mai-Aug 2016'): ")
print('\n')



listeKunden = getList(input_path)
try:
    SaftyClass.alleAdressenVorhanden(input_path, listeKunden)
except ValueError as e:
    print(e.args[0])
    print("Schliessen Sie dieses Programm, fügen Sie die Adresse hinzu und starten Sie das Programm neu.")
    while True:
        continue
    

for i in range(len(listeKunden)):

    kunde = listeKunden[i]
    totalBetrag = 0
    if basis+i > 999:
        rechnungsnummer = (basis + i)%999
    else:
        rechnungsnummer = basis + i
 
    information = getInformation(input_path, kunde)
    try:
        optionen = getPriceDescriptionOptions(input_path, kunde)
    except ValueError as e:
        print(e.args[0])
        print("Dieser Kunde wird übersprungen. Dessen Rechnung sollte manuell erstellt werden.")
        print("ACHTUNG: Die Rechnungsnummer " + str(rechnungsnummer) + " bleibt für " + kunde + " reserviert.\n")
        continue
    printKundenOptionen(optionen, kunde)

    titel = getTitle(output_path, rechnungsnummer, kunde, anhaengsel)
    archiv = xlsxwriter.Workbook(titel)
    archiv_sheet = archiv.add_worksheet()
    font1 = setFonts(archiv)
    
    drucker_sheet = zum_drucken.add_worksheet()

    setColumnWidths(archiv_sheet)
    setColumnWidths(drucker_sheet)
    fillHeader(archiv_sheet, information, font1)
    fillHeader(drucker_sheet, information, font2)
    fillRechnungDatum(archiv_sheet, rechnungsnummer, periode, font1)
    fillRechnungDatum(drucker_sheet, rechnungsnummer, periode, font2)
    startTable(archiv_sheet, font1)
    startTable(drucker_sheet, font2)

    zeile = 19

    while True:
        while True:
            try:
                datum = input('Falls es noch eine Lieferung gibt, dessen Datum eintragen (sonst einfach ENTER drücken): ').strip()
                SaftyClass.pruefeDatum(datum)
                break
            except ValueError as e:
                print(e.args[0])
                print('Versuchen Sie es nochmals.')

        
        if (datum == ''):
            break
        
        # verschönere noch das Datum
        datum = getDatum(datum)
        tmp = addLine(zeile, archiv_sheet, drucker_sheet, datum, optionen, font1, font2)
        totalBetrag += tmp[1]
        zeile += tmp[0]

    archiv_sheet.write(zeile, 3, 'Total', font1['tabelle_text_doppelt_unterstrichen'])
    archiv_sheet.write(zeile, 4, totalBetrag, font1['tabelle_float_doppelt_unterstrichen'])

    drucker_sheet.write(zeile, 3, 'Total', font2['tabelle_text_doppelt_unterstrichen'])
    drucker_sheet.write(zeile, 4, totalBetrag, font2['tabelle_float_doppelt_unterstrichen'])

    zeile += 3
    printRest(zeile, archiv_sheet, font1)
    printRest(zeile, drucker_sheet, font2)

    updateRechnungsnummernDatei(sheet_rn, rechnungsnummer, kunde, totalBetrag, i+1)



    archiv.close()
    print("\n")
    



rechnungsnummern.close()
zum_drucken.close()





















    
