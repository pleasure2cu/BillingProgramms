import xlsxwriter
import datetime
import glob



def findPricePerPiece(anzahl_einheiten):
	""" caputres the price structure """

	preis_pro_einheit = 0
	
	if anzahl_einheiten > 40:
		preis_pro_einheit = 28
	elif anzahl_einheiten > 35:
		preis_pro_einheit = 29
	elif anzahl_einheiten > 30:
		preis_pro_einheit = 30
	elif anzahl_einheiten > 25:
		preis_pro_einheit = 31
	elif anzahl_einheiten > 20:
		preis_pro_einheit = 32
	elif anzahl_einheiten > 10:
		preis_pro_einheit = 33
	else:
		preis_pro_einheit = 34

	return preis_pro_einheit


def getBillNumber(offset, basis):
	rechnungsbetrage = ''
	mod_nummer = (offset+basis)%999
	if mod_nummer == 0:
		rechnungsnummer = '999'
	elif mod_nummer < 100:
		if mod_nummer < 10:
			rechnungsnummer = '00' + str(mod_nummer)
		else:
			rechnungsnummer = '0' + str(mod_nummer)
	else:
		rechnungsnummer = str(mod_nummer)
	return rechnungsnummer


def setColumnWidths(worksheet, sheet):
	worksheet.set_column('A:A', 9.86)
	worksheet.set_column('B:B', 8)
	worksheet.set_column('C:C', 33.57)
	worksheet.set_column('D:D', 14.29)
	worksheet.set_column('E:E', 16)
	worksheet.set_column('F:F', 5)
	
	sheet.set_column('A:A', 9.86)
	sheet.set_column('B:B', 8)
	sheet.set_column('C:C', 33.57)
	sheet.set_column('D:D', 14.29)
	sheet.set_column('E:E', 16)
	sheet.set_column('F:F', 5)




### alle grundsätzlichen Dinge

path_for_input = ''
path_for_output = ''

# Zeiten
t = datetime.date.today()
today = t.strftime('%d.%m.%Y')
in_one_month = (t + datetime.timedelta(365/12)).strftime('%d.%m.%Y')    

# Kunden
kunden = glob.glob(path+'*.txt')

# bringe die Rechungsperiode in Erfahrung
rechnungsperiode = input('Rechnungsperiode (wird auf den Rechnungen stehen): ')

# bringe die erste verwendbare Rechnungsnummer in Erfahrung
basis = int(input('Erste Rechnungsnummer die verwendet werden darft: '))
print('')

# offset der Rechnungsnummer
offset = 0


# erstellen der Datei mit den Rechnungsnummern
rechnungsnummern_liste = xlsxwriter.Workbook(path_for_output+'Rechnungsnummern.xlsx')
tatsachliche_liste = rechnungsnummern_liste.add_worksheet()
tatsachliche_liste_counter = 1
tatsachliche_liste.write('A1', 'Nr.')
tatsachliche_liste.write('B1', 'Beschrieb')
tatsachliche_liste.write('C1', 'Name')
tatsachliche_liste.write('D1', 'Betrag')
tatsachliche_liste.write('E1', 'Rechnungsdatum')
rechnungsbetrage = rechnungsnummern_liste.add_format({'num_format':'0.00'})

# erstellen der Datei zum Drucken und die benötigten Formate
druck_datei = xlsxwriter.Workbook(path_for_output + 'zum_Drucken.xlsx')

### formats
# statischer header
Fett_vierzehn = druck_datei.add_format({'font_size':14, 'bold':True})
Rechts_zwolf = druck_datei.add_format({'align':'right', 'font_size':12})
Rechts_zwolf_doppelt = druck_datei.add_format({'align':'right', 'font_size':12, 'bottom':6})

#kunde
Fett = druck_datei.add_format({'bold':True})

# grosser Schriftzug 'Rechnung'
Gross_rot_doppelt = druck_datei.add_format({'font_size':24, 'font_color':'red', 'bottom':6})

# Datum bis Rechnungsnummer
Umrandet = druck_datei.add_format({'border':1, 'align':'top'})
Umrandet_rot = druck_datei.add_format({'border':1, 'font_color':'red'})

# Tabelle
Integer_tabelle = druck_datei.add_format({'align':'right', 'border':1, 'align':'top'})
Betrage_tabelle = druck_datei.add_format({'align':'right', 'border':1, 'num_format':'0.00', 'align':'top'})
Umrandet_betrage_grauhinterlegt = druck_datei.add_format({'border':1, 'num_format':'0.00', 'align':'right', 'bg_color':'#DBDBDB'})
Umrandet_betrage_grauhinterlegt_fett = druck_datei.add_format({'border':1, 'num_format':'0.00', 'align':'right', 'bg_color':'#DBDBDB', 'bold':True})
Beschreibung = druck_datei.add_format({'border':1, 'text_wrap':True})
        

# Schluss
Rechts_umrandet = druck_datei.add_format({'align':'right', 'border':1})
Umrandet_rot_rechts_grauhinterlegt = druck_datei.add_format({'border':1, 'font_color':'red', 'align':'right', 'bg_color':'#DBDBDB'})
Rot = druck_datei.add_format({'font_color':'red'})




for kunde in kunden:
    print('Neuer Kunde: ')
    
    # den Name vom Pfad zum Kunde.txt heraus filtern
    nickname = kunde.split("\\")
    nickname = nickname[len(nickname)-1].split('.')[0]

    
    # grosse Unterscheidung, ob man diesem Kunden eine Rechnung schreiben muss
    if 'n' in (input('Gibt es eine Rechnung an ' + nickname + '? (n = nein, j = ja) ')):
        pass
    else:
        # ein neues sheet in der Druck Datei erschaffen
        sheet = druck_datei.add_worksheet()
        sheet.set_margins(left = 0.55, right = 0.55)
        
        # man will diesem Kunden eine Rechnung schreiben -> man braucht die Personal Informationen
        name = ''
        adresse = ''
        plz_ort = ''
        with open(kunde, 'r') as file:
            name = file.readline()
            adresse = file.readline()
            plz_ort = file.readline()

        # Rechnungsnummer (String) erstellen
        rechnungsnummer = getBillNumber(offset, basis)


        # Dateinamen erstellen
        titel = rechnungsnummer + ' ' + nickname + ' ' + 'Lohnarbeiten.xlsx'

        # Datei und sheet erstellen
        workbook = xlsxwriter.Workbook(path_for_output + titel)
        worksheet = workbook.add_worksheet()
        worksheet.set_margins(left = 0.55, right = 0.55)
        


        # bringe als erstes die Kolonnen in die gewüste Grösse
        setColumnWidths(worksheet, sheet)



        ### formats
        # statischer header
        fett_vierzehn = workbook.add_format({'font_size':14, 'bold':True})
        rechts_zwolf = workbook.add_format({'align':'right', 'font_size':12})
        rechts_zwolf_doppelt = workbook.add_format({'align':'right', 'font_size':12, 'bottom':6})

        #kunde
        fett = workbook.add_format({'bold':True})

        # grosser Schriftzug 'Rechnung'
        gross_rot_doppelt = workbook.add_format({'font_size':24, 'font_color':'red', 'bottom':6})

        # Datum bis Rechnungsnummer
        umrandet = workbook.add_format({'border':1, 'align':'top'})
        umrandet_rot = workbook.add_format({'border':1, 'font_color':'red'})

        # Tabelle
        integer_tabelle = workbook.add_format({'align':'right', 'border':1, 'align':'top'})
        betrage_tabelle = workbook.add_format({'align':'right', 'border':1, 'num_format':'0.00', 'align':'top'})
        umrandet_betrage_grauhinterlegt = workbook.add_format({'border':1, 'num_format':'0.00', 'align':'right', 'bg_color':'#DBDBDB'})
        umrandet_betrage_grauhinterlegt_fett = workbook.add_format({'border':1, 'num_format':'0.00', 'align':'right', 'bg_color':'#DBDBDB', 'bold':True})
        beschreibung = workbook.add_format({'border':1, 'text_wrap':True})
        

        # Schluss
        rechts_umrandet = workbook.add_format({'align':'right', 'border':1})
        umrandet_rot_rechts_grauhinterlegt = workbook.add_format({'border':1, 'font_color':'red', 'align':'right', 'bg_color':'#DBDBDB'})
        rot = workbook.add_format({'font_color':'red'})



        # statischer header
        worksheet.write('A1', 'Name des Betriebs', fett_vierzehn)
        worksheet.write('F1', 'Musterstrasse 1', rechts_zwolf)
        worksheet.write('F2', 'PLZ und Ort', rechts_zwolf)
        worksheet.write('F3', 'Telefon: 012 345 67 89', rechts_zwolf)
        worksheet.write('F4', 'MwSt: MWST-Nr.', rechts_zwolf_doppelt)
        for x in range(5):
            worksheet.write(3, x, '', rechts_zwolf_doppelt)

        sheet.write('A1', 'Name des Betriebs', Fett_vierzehn)
        sheet.write('F1', 'Musterstrasse 1', Rechts_zwolf)
        sheet.write('F2', 'PLZ und Ort', rechts_zwolf)
        sheet.write('F3', 'Telefon: 012 345 67 89', Rechts_zwolf)
        sheet.write('F4', 'MwSt: MWST-Nr.', Rechts_zwolf_doppelt)
        for x in range(5):
            sheet.write(3, x, '', Rechts_zwolf_doppelt)


        # kunde im excel sheet einschreiben
        worksheet.write('A8', 'Kunde:', fett)
        worksheet.write('D8', name)
        worksheet.write('D9', adresse)
        worksheet.write('D10', plz_ort)

        sheet.write('A8', 'Kunde:', Fett)
        sheet.write('D8', name)
        sheet.write('D9', adresse)
        sheet.write('D10', plz_ort)


        # grosser Schriftzug 'Rechnung'
        worksheet.write('A14', 'Rechnung', gross_rot_doppelt)
        for x in range(1,6):
            worksheet.write(13, x, '', rechts_zwolf_doppelt)

        sheet.write('A14', 'Rechnung', Gross_rot_doppelt)
        for x in range(1,6):
            sheet.write(13, x, '', Rechts_zwolf_doppelt)


        # Datum bis Rechnungsnummer
        worksheet.write('A16', 'Datum', fett)
        worksheet.write('D16', today, umrandet)
        worksheet.write('A18', 'Rechnungsperiode:')
        worksheet.write('D18', rechnungsperiode, umrandet_rot)
        worksheet.write('A19', 'Rechnungsnummer:')
        worksheet.write('D19', rechnungsnummer, umrandet_rot)

        sheet.write('A16', 'Datum', Fett)
        sheet.write('D16', today, Umrandet)
        sheet.write('A18', 'Rechnungsperiode:')
        sheet.write('D18', rechnungsperiode, Umrandet_rot)
        sheet.write('A19', 'Rechnungsnummer:')
        sheet.write('D19', rechnungsnummer, Umrandet_rot)


        ### Tabelle
        # oberste Zeile
        top_zeile = ['Datum', 'Anzahl', 'Beschreibung', 'Preis', 'Betrag', 'MwSt']
        counter = 0
        for inhalt in top_zeile:
            worksheet.write(20, counter, inhalt, umrandet)
            sheet.write(20, counter, inhalt, Umrandet)
            counter +=1

        # Text für die Beschreibung
        anhangsel = {'b':', Folie blau', 'p':', Folie pink', 'd':', doppelt gewickelt', 's':', Siliermittel', 'n':'', ' ':''}
        
        # Preis bereich
        aufpreise = {'b':2, 'p':2, 's':3, 'd':3, ' ':0, 'n':0}            # aufpreisepreise
        preis_transportieren = 145

        # vorbereitungen, damit die MwSt gut ausgerechnet werden kann und die Tabelle funktioniert
        zeilen_counter = 21
        fertig = False
        total_betrage_ohne_grosse_mwst = 0
        total_betrage_ohne_kleine_mwst = 0
        total_betrage_bereits_mit_mwst = 0

        zeile = 21
        verbrauchte_zeilen = 0
        
        while not fertig:
            arbeits_datum = ''
            ordnung = False
            
            if not fertig:
                while not ordnung:
                    arbeits_datum = input('Falls noch eine Arbeit berechnet wird, dessen Datum eintragen (sonst einfach ENTER drücken): ').strip()
                    fertig = (len(arbeits_datum) < 3)

                    # prüfen, ob die Eingabe korrekt ist:
                        # es gibt im Minimum 2 und in Maximum 3 Zahlen
                        # die Tag Zahl ist eine oder zwei Zeichen lang
                        # die Monat Zahl ist eine oder zwei Zeichen lang
                        # die Jahrzahl ist 1, 2, 3 oder 4 Zeichen lang
                    ordnung = True
                    helper = arbeits_datum.split('.')
                    
                    if fertig:
                        break                    
                    if not len(helper) in [2,3]:
                        ordnung = False
                    if not len(helper[0]) in [1,2]:
                        ordnung = False
                    if not len(helper[1]) in [1,2]:
                        ordnung = False
                    if len(helper) == 3:
                        if not len(helper[2]) in [0,1,2,3,4]:
                            ordnung = False
                        if len(helper[2]) == 0:
                            arbeits_datum = arbeits_datum[:len(arbeits_datum)-1]

                    if ordnung:
                        break
                    
                

            if not fertig:
                # das Datum könnte in einer unschönen Darstellung gegeben sein, darauf wird hier reagiert
                # die angeschauten Situationen sind:
                    # 0 fehlt bei Monat und/oder Tag
                    # Jahresangabe ist nur teilweise gegeben (z.B. 16)
                    # Jahresangabe fehlt ganz
                pieces = arbeits_datum.split('.')
                
                if len(pieces[0]) == 1:
                    arbeits_datum = '0'+pieces[0]
                elif len(pieces[0]) == 2:
                    arbeits_datum = pieces[0]
                
                if len(pieces[1]) == 1:
                    arbeits_datum += '.0'+pieces[1]
                elif len(pieces[1]) == 2:
                    arbeits_datum += '.' + pieces[1]

                if len(pieces) == 2:
                    arbeits_datum += '.' + t.strftime('%Y')
                elif len(pieces) == 3:
                    if len(pieces[2]) == 1:
                        arbeits_datum += '.200' + pieces[2]
                    elif len(pieces[2]) == 2:
                        arbeits_datum += '.20' + pieces[2]
                    elif len(pieces[2]) == 3:
                        arbeits_datum += '.2' + pieces[2]
                    elif len(pieces[2]) == 4:
                        arbeits_datum += '.' + pieces[2]

                        
                art = input('Um was handelt sich die Arbeit (0 = Ballen pressen, 1 = Transport, 2 = anderes): ').strip()
                while not art in ['0', '1', '2']:
                    print('Leider kann die Angabe nicht verarbeitet werden (keine passende Zahl eingegeben). Versuchen Sie es nocheinmal.')
                    art = input('Um was handelt sich die Arbeit (0 = Ballen pressen, 1 = Transport, 2 = anderes): ')

                if art == '0':
                    zulassig = False
                    output_preis_pro_einheit = []
                    output_menge = []
                    output_beschreibungen = []
                    anzahl_einheiten = 0
                    
                    while not zulassig:
                        while anzahl_einheiten == 0:
                            try:
                                anzahl_einheiten = int(input('Wie viele Ballen wurden am '+arbeits_datum+' gepresst? '))
                            except ValueError:
                                print('Die Eingabe konnte nicht verarbeitet werden (keine ganze Zahl eingegeben). Versuchen Sie es nochmal.')
                                anzahl_einheiten = 0
                                
                        # hier wird herausgefunden, was der grundsätzliche Preis pro Balle ist
                        preis_pro_einheit = findPricePerPiece(anzahl_einheiten)

                        

                        # nun werden noch die zusätzlichen Programme bestimmt und der schlussendliche Preis. Falls es mehrere 'Untergruppen' an diesem Tag gibt, wird das ebenfalls berücksichtig
                        liste_zusammensetzung = []
                        ready = False
                        while not ready:
                            ready = True
                            zusammensetzung = input('Wie setzen sich die ' + str(anzahl_einheiten) + ' zusammen? (zahl art ...) ')

                            liste_zusammensetzung = zusammensetzung.split()
                            if ',' in zusammensetzung:
                                liste_zusammensetzung_zwei = []
                                for x in liste_zusammensetzung:
                                    if ',' in x:
                                        for y in x.strip().split(','):
                                            if y != '': 
                                                liste_zusammensetzung_zwei.append(y.strip())
                                    else:
                                        liste_zusammensetzung_zwei.append(x)
                                liste_zusammensetzung = liste_zusammensetzung_zwei

                            if len(liste_zusammensetzung) == 1:
                                for x in liste_zusammensetzung[0]:
                                    try:
                                        aufpreise[x]
                                    except KeyError:
                                        ready = False
                                        print('Leider konnte die Eingabe nicht verarbeitet werden (ein Zeichen kann nicht verarbeitet werden). Versuchen Sie es nochmal.')
                                        break
                            elif len(liste_zusammensetzung)%2 == 1:
                                ready = False
                                print('Leider konnte die Eingabe nicht verarbeitet werden (ungerade Anzahl von Angaben). Versuchen Sie es nocheinmals.')
                            else:
                                total_across = 0
                                for x in range(len(liste_zusammensetzung)):
                                    if x%2 == 0:
                                        try:
                                            total_across += int(liste_zusammensetzung[x])
                                        except ValueError:
                                            ready = False
                                            print('Leider konnte die Eingabe nicht verarbeitet werden (Eingabe an Stelle ' + str(x+1) + ' wurde eine ganze Zahle erwartet). Versuchen Sie es nochmal.')
                                            break
                                    else:
                                        for y in liste_zusammensetzung[x]:
                                            try:
                                                aufpreise[y]
                                            except TypeError:
                                                ready = False
                                                print('Leider konnte die Eingabe nicht verarbeitet werden ('+y+' konnte nicht verarbeitet werden). Versuchen Sie es nochmal.')
                                                break
                                            except KeyError:
                                                ready = False
                                                print('Leider konnte die Eingabe nicht verarbeitet werden ('+y+' ist ein unzulässiger Buchstaben). Versuchen Sie es nochmal.')
                                                break
                                            
                                if total_across != anzahl_einheiten:
                                    ready = False
                                    print('Leider konnte die Eingabe nicht verarbeitet werden (die Anzahl totaler Ballen stimmt nicht mit den Teilangaben überein). Versuchen Sie es nochmal.')
                                    anzahl_einheiten = 0
                                    while anzahl_einheiten == 0:
                                        try:
                                            anzahl_einheiten = int(input('Wie viele Ballen wurden am '+arbeits_datum+' gepresst? '))
                                        except ValueError:
                                            print('Die Eingabe konnte nicht verarbeitet werden (keine ganze Zahl eingegeben). Versuchen Sie es nochmal.')
                                            anzahl_einheiten = 0
                                    # hier wird herausgefunden, was der grundsätzliche Preis pro Balle ist
                                    preis_pro_einheit = findPricePerPiece(anzahl_einheiten)
                                                
                        
                        
                        # mit diesem if ... else statement wird nun dem Benutzereintrag die Informationen entzogen und in den Listen output_preis_pro_einheit, output_menge und output_beschreibungen gespeichert
                        # für jede Art von Ballenkombination gibt es je einen Eintrag in 3 Listen
                        # das if statement ist nur hier, damit die Eingabe für den Benutzer einfacher ist, falls alle Ballen die gleiche Konfiguratione haben
                        if len(liste_zusammensetzung) == 1:
                            zulassig = True
                            zwischen_preis = preis_pro_einheit
                            beschreibung_auftrag = 'Ballen pressen'
                            for bit in liste_zusammensetzung[0]:
                                zwischen_preis += aufpreise[bit]
                                beschreibung_auftrag += anhangsel[bit]
                            output_preis_pro_einheit.append(zwischen_preis)
                            output_menge.append(anzahl_einheiten)
                            output_beschreibungen.append(beschreibung_auftrag)
                            verbrauchte_zeilen += 1
                            
                            
                                
                        else:
                            trigger = 0
                            for piece in liste_zusammensetzung: # mit trigger wird erreicht, dass die Einträge der Liste mit Zahlen anders behandelt wird als jene mit den Buchstaben
                                if trigger == 0:
                                    output_menge.append(int(piece))
                                    trigger = 1
                                elif trigger == 1:
                                    preis_pro_untereinheit = preis_pro_einheit
                                    beschreibung_auftrag = 'Ballen pressen'
                                    if len(piece) > 1:
                                        verbrauchte_zeilen += 2
                                    else:
                                        verbrauchte_zeilen += 1
                                    for bit in piece:
                                        preis_pro_untereinheit += aufpreise[bit]
                                        beschreibung_auftrag += anhangsel[bit]
                                    output_preis_pro_einheit.append(preis_pro_untereinheit)
                                    output_beschreibungen.append(beschreibung_auftrag)
                                    trigger = 0

                            counter = 0
                            for x in output_menge:
                                counter += x

                            if counter == anzahl_einheiten:
                                zulassig = True
                            else:
                                print('Die beiden Angaben stimmen nicht überein. Versuchen Sie es nochmals.')

                                    

                    # die drei ouput Listen (output_preis_pro_einheit, output_menge, output_beschreibungen) wurden erstellt, nun müssen diese Informationen genutzt werden um die Tabelle zu
                    # vervollständigen
                    worksheet.write(zeile, 0, arbeits_datum, umrandet)
                    sheet.write(zeile, 0, arbeits_datum, umrandet)
                    for x in range(len(output_preis_pro_einheit)):
                        if x != 0:
                            worksheet.write(zeile, 0, '', umrandet)
                            sheet.write(zeile, 0, '', umrandet)
                            
                        worksheet.write(zeile, 1, output_menge[x], integer_tabelle)
                        worksheet.write(zeile, 2, output_beschreibungen[x], beschreibung)
                        worksheet.write(zeile, 3, output_preis_pro_einheit[x], betrage_tabelle)
                        worksheet.write(zeile, 4, output_menge[x]*output_preis_pro_einheit[x], betrage_tabelle)
                        worksheet.write(zeile, 5, 1, integer_tabelle)

                        sheet.write(zeile, 1, output_menge[x], Integer_tabelle)
                        sheet.write(zeile, 2, output_beschreibungen[x], Beschreibung)
                        sheet.write(zeile, 3, output_preis_pro_einheit[x], Betrage_tabelle)
                        sheet.write(zeile, 4, output_menge[x]*output_preis_pro_einheit[x], Betrage_tabelle)
                        sheet.write(zeile, 5, 1, Integer_tabelle)
                        
                        total_betrage_ohne_kleine_mwst += output_menge[x]*output_preis_pro_einheit[x]
                        zeile +=1

                    
                elif art == '1':    # transport
                    done = False
                    while not done:
                        try:
                            anzahl_stunden = float(input('Wie viele Stunden wurde transportiert: '))
                            worksheet.write(zeile, 0, arbeits_datum, umrandet)
                            worksheet.write(zeile, 1, anzahl_stunden, integer_tabelle)
                            worksheet.write(zeile, 2, 'Std Ballen transportieren', umrandet)
                            worksheet.write(zeile, 3, preis_transportieren, betrage_tabelle)
                            worksheet.write(zeile, 4, preis_transportieren*anzahl_stunden, betrage_tabelle)
                            worksheet.write(zeile, 5, 1, integer_tabelle)

                            sheet.write(zeile, 0, arbeits_datum, Umrandet)
                            sheet.write(zeile, 1, anzahl_stunden, Integer_tabelle)
                            sheet.write(zeile, 2, 'Std Ballen transportieren', Umrandet)
                            sheet.write(zeile, 3, preis_transportieren, Betrage_tabelle)
                            sheet.write(zeile, 4, preis_transportieren*anzahl_stunden, Betrage_tabelle)
                            sheet.write(zeile, 5, 1, Integer_tabelle)
                            
                            total_betrage_ohne_kleine_mwst += preis_transportieren*anzahl_stunden
                            zeile += 1
                            verbrauchte_zeilen += 1
                            done = True
                        except ValueError:
                            print('Die gegebene Zahl konnte nicht gelesen werden. Versuchen Sie es nochmal.')
                            pass

                elif art == '2':    # anderes
                    done = False
                    while not done:
                        try:
                            worksheet.write(zeile, 0, arbeits_datum, umrandet)
                            anzahl_einheiten = float(input('Was soll bei "Anzahl" stehen? '))
                            worksheet.write(zeile, 1, anzahl_einheiten, integer_tabelle)
                            eintrag_beschreibung = input('Was soll bei "Beschreibung" stehen? ')
                            worksheet.write(zeile, 2, eintrag_beschreibung, beschreibung)
                            preis_pro_einheit = float(input('Was soll bei "Preis" stehen? '))
                            worksheet.write(zeile, 3, preis_pro_einheit, betrage_tabelle)
                            worksheet.write(zeile, 4, anzahl_einheiten*preis_pro_einheit, betrage_tabelle)
                            mwst_kategorie = int(input('Zu welcher MwSt.-Kategorie gehört der Eintrag (0 = MwSt. bereits enthalten, 1 = 2.5 prozent, 2 = 8 prozent)? '))
                            worksheet.write(zeile, 5, mwst_kategorie, integer_tabelle)

                            sheet.write(zeile, 0, arbeits_datum, Umrandet)
                            sheet.write(zeile, 1, anzahl_einheiten, Integer_tabelle)
                            sheet.write(zeile, 2, eintrag_beschreibung, Beschreibung)
                            sheet.write(zeile, 3, preis_pro_einheit, Betrage_tabelle)
                            sheet.write(zeile, 4, anzahl_einheiten*preis_pro_einheit, Betrage_tabelle)
                            sheet.write(zeile, 5, mwst_kategorie, Integer_tabelle)
                            
                            zeile += 1
                            verbrauchte_zeilen += 1

                            # Es werden nur Type mismatchs durch den except-Block unten abgefangen und
                            # eine falsche MwSt. Kategorie
                            if (mwst_kategorie in [0,1,2]):
                                done = True
                            else:
                                done = False
                                print('Die MwSt. wurde falsch gesetzt. Versuchen Sie es nochmal.')

                            if mwst_kategorie == 1:
                                total_betrage_ohne_kleine_mwst += (anzahl_einheiten*preis_pro_einheit)
                            elif mwst_kategorie == 2:
                                total_betrage_ohne_grosse_mwst += (anzahl_einheiten*preis_pro_einheit)
                            elif mwst_kategorie == 0:
                                total_betrage_bereits_mit_mwst += (anzahl_einheiten*preis_pro_einheit)
                            
                        except ValueError:
                            print('Ein Eintrag konnte nicht interpretiert werden. Versuchen Sie es nochmal.')
                            pass
                print('')

        # es wurde getrackt wie viele Zeilen verbraucht wurden, falls dies nur wenige sind, werden aus optischen Gründen weitere leere Zeilen hinzugefügt
        if verbrauchte_zeilen < 14:
            for x in range(14-verbrauchte_zeilen):
                for y in range(6):
                    worksheet.write(zeile, y, '', umrandet)
                    sheet.write(zeile, y, '', umrandet)
                zeile += 1
            
                    
        # Zeile 'Total Netto'
        worksheet.write(zeile, 2, 'Total Netto', umrandet)
        worksheet.write(zeile, 3, '', umrandet)
        worksheet.write(zeile, 4, '=SUM(E22:E'+str(zeile)+')', umrandet_betrage_grauhinterlegt_fett)
        worksheet.write(zeile, 5, '', umrandet)

        sheet.write(zeile, 2, 'Total Netto', Umrandet)
        sheet.write(zeile, 3, '', Umrandet)
        sheet.write(zeile, 4, '=SUM(E22:E'+str(zeile)+')', Umrandet_betrage_grauhinterlegt_fett)
        sheet.write(zeile, 5, '', Umrandet)
        
        zeile += 1

        # Zeile '8% MwSt'
        worksheet.write(zeile, 2, '8   % Mehrwertsteuer von', umrandet)
        sheet.write(zeile, 2, '8   % Mehrwertsteuer von', umrandet)
        if total_betrage_ohne_grosse_mwst > 0:
            worksheet.write(zeile, 3, total_betrage_ohne_grosse_mwst, umrandet_betrage_grauhinterlegt)
            worksheet.write(zeile, 4, '=D'+str(zeile+1)+'*0.08', umrandet_betrage_grauhinterlegt)

            sheet.write(zeile, 3, total_betrage_ohne_grosse_mwst, Umrandet_betrage_grauhinterlegt)
            sheet.write(zeile, 4, '=D'+str(zeile+1)+'*0.08', Umrandet_betrage_grauhinterlegt)
            
        else:
            worksheet.write(zeile, 3, '', umrandet)
            worksheet.write(zeile, 4, '', umrandet)

            sheet.write(zeile, 3, '', Umrandet)
            sheet.write(zeile, 4, '', Umrandet)
            
        worksheet.write(zeile, 5, 2, integer_tabelle)
        sheet.write(zeile, 5, 2, Integer_tabelle)
        zeile += 1

        
        # Zeile '2.5% MwSt'
        worksheet.write(zeile, 2, '2.5% Mehrwertsteuer von', umrandet)
        sheet.write(zeile, 2, '2.5% Mehrwertsteuer von', umrandet)
        if total_betrage_ohne_kleine_mwst > 0:
            worksheet.write(zeile, 3, total_betrage_ohne_kleine_mwst, umrandet_betrage_grauhinterlegt)
            worksheet.write(zeile, 4, '=D'+str(zeile+1)+'*0.025', umrandet_betrage_grauhinterlegt)

            sheet.write(zeile, 3, total_betrage_ohne_kleine_mwst, Umrandet_betrage_grauhinterlegt)
            sheet.write(zeile, 4, '=D'+str(zeile+1)+'*0.025', Umrandet_betrage_grauhinterlegt)
        else:
            worksheet.write(zeile, 3, '', umrandet)
            worksheet.write(zeile, 4, '', umrandet)

            sheet.write(zeile, 3, '', Umrandet)
            sheet.write(zeile, 4, '', Umrandet)
            
        worksheet.write(zeile, 5, 1, integer_tabelle)
        sheet.write(zeile, 5, 1, Integer_tabelle)
        zeile += 1


        # Zeile 'Total Rechnung'
        worksheet.write(zeile, 2, 'Total Rechnung', fett)
        worksheet.write(zeile, 4, '=ROUND((D'+str(zeile)+'*1.025+D'+str(zeile-1)+'*1.08)*20, 0)/20+'+str(total_betrage_bereits_mit_mwst), umrandet_betrage_grauhinterlegt_fett)

        sheet.write(zeile, 2, 'Total Rechnung', Fett)
        sheet.write(zeile, 4, '=ROUND((D'+str(zeile)+'*1.025+D'+str(zeile-1)+'*1.08)*20, 0)/20+'+str(total_betrage_bereits_mit_mwst), umrandet_betrage_grauhinterlegt_fett)
        
        zeile += 1



        # Schluss
        zeile += 2
        worksheet.write(zeile, 0, 'Konditionen:', umrandet)
        worksheet.write(zeile, 1, '', umrandet)
        worksheet.write(zeile, 2, 'netto', rechts_umrandet)

        sheet.write(zeile, 0, 'Konditionen:', Umrandet)
        sheet.write(zeile, 1, '', Umrandet)
        sheet.write(zeile, 2, 'netto', Rechts_umrandet)
        
        zeile += 1
        worksheet.write(zeile, 0, 'zahlbar bis:  ', umrandet)
        worksheet.write(zeile, 1, '', umrandet)
        worksheet.write(zeile, 2, in_one_month, umrandet_rot_rechts_grauhinterlegt)

        sheet.write(zeile, 0, 'zahlbar bis:  ', Umrandet)
        sheet.write(zeile, 1, '', Umrandet)
        sheet.write(zeile, 2, in_one_month, Umrandet_rot_rechts_grauhinterlegt)
        
        zeile += 2
        worksheet.write(zeile, 0, 'Herzlichen Dank für den geschätzten Auftrag')

        sheet.write(zeile, 0, 'Herzlichen Dank für den geschätzten Auftrag')
        
        zeile += 2
        worksheet.write(zeile, 0, 'Mit freundlichen Grüssen')

        sheet.write(zeile, 0, 'Mit freundlichen Grüssen')
        
        zeile += 1
        worksheet.write(zeile, 0, 'Name des Betriebs', rot)

        sheet.write(zeile, 0, 'Name des Betriebs', Rot)
        
        

        # schliessen der Archivdatei
        workbook.close()

        # bookkeeping
        offset += 1
        print('\n')

        # aufnehmen in die Rechnungnummern Liste
        for y in range(6):
            tatsachliche_liste.write(tatsachliche_liste_counter, 0, rechnungsnummer)
            tatsachliche_liste.write(tatsachliche_liste_counter, 1, 'Lohnarbeiten')
            tatsachliche_liste.write(tatsachliche_liste_counter, 2, name)
            tatsachliche_liste.write(tatsachliche_liste_counter, 3, int((total_betrage_ohne_kleine_mwst*1.025+total_betrage_ohne_grosse_mwst*1.08)*20.0+0.5)/20.0,rechnungsbetrage)
            tatsachliche_liste.write(tatsachliche_liste_counter, 4, today)
        tatsachliche_liste_counter += 1


rechnungsnummern_liste.close()
druck_datei.close()

