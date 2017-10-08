import xlsxwriter
import time
import datetime
import glob
price = 0.6 
path = ''
var = 0


periode = input("Rechnungsperiode (steht auf den Rechnungen): ")
basis = int(input("Erste Rechnungsnummer die verwendet werden darf: "))
anhaengsel = input("Was soll am Schluss der Exel-Datei Namen stehen (Bspiele: '2-16', 'Mai-Aug 2016'): ")


# die Liste mit den Rechnungsnummern erstellen, sowie header und ein format
rechnungsnummern = xlsxwriter.Workbook(path+'Rechnungsnummern_Liste.xlsx')
worksheet_rn = rechnungsnummern.add_worksheet()
worksheet_rn.write('A1', 'Nr.')
worksheet_rn.write('C1', 'Name')
worksheet_rn.write('D1', 'Betrag')
worksheet_rn.write('E1', 'Rechnungsdatum')
zahlen_format = rechnungsnummern.add_format({'num_format':'0.00'})


# to print Dokument erstellen
zu_drucken = xlsxwriter.Workbook(path+'zum drucken.xlsx')

# das Dokument mit der Reihenfolge lesen und die Reihenfolge als Liste speichern
reihenfolge = []
file = path + '\\Reihenfolge.txt'
with open(file, 'r') as file:
    while True:
        line = file.readline().rstrip()
        if not line:
            break
        reihenfolge.append(line)
        print(line)



for name in reihenfolge:

    # some definitions at the beginning (momentaner Kunde, name des Kunden, Titel der xlsx Datei und die Rechnungsnummer)
    path_zu_adresse = path + '\\' + name + '.txt'
    familie = ''
    Adresse = ''
    PLZ_Ort = ''
    try:
        with open(path_zu_adresse, 'r') as file:
            familie = file.readline()
            Adresse = file.readline()
            PLZ_Ort = file.readline()
    except FileNotFoundError:
        print('Es gibt keine Adresse zur Person ' + name + '. Es wird mit der nächsten Person weiter gefahren.')
        continue

    menge_eier = int(input("Wie viele Eier bezog "+name+": "))
    
    
    mom_rechnungsnr = ''
    if (var+basis) == 999:
        mom_rechnungsnr = "999"
    elif((var+basis)%999) < 100:
        if ((var+basis)%999) < 10:
            mom_rechnungsnr = "00"+str((var+basis)%999)
        else:
            mom_rechnungsnr = "0"+str((var+basis)%999)
    else:
       mom_rechnungsnr = str(var+basis)

    title = mom_rechnungsnr + " Eier " + name + " " + anhaengsel + ".xlsx"

    
    ################################################# original #############################################################################################################
    
    workbook = xlsxwriter.Workbook(path+title)

    ### fonts
    #default
    d = workbook.add_format({'font_name':'Verdana', 'font_size':12})
    dd = workbook.add_format({'font_name':'Verdana', 'font_size':12, 'bottom':6})

    # Betriebsge...
    bg = workbook.add_format({'bold':True, 'font_name':'Verdana', 'font_size':12})
    bgu = workbook.add_format({'bold':True, 'font_name':'Verdana', 'font_size':12, 'border':1})

    # rechtsbuendig
    rechts = workbook.add_format({'font_name':'Verdana', 'font_size':12, 'align':'right'})
    rechtsd = workbook.add_format({'font_name':'Verdana', 'font_size':12, 'align':'right', 'bottom':6})

    #Rechnung
    rechnung = workbook.add_format({'font_name':'Verdana', 'font_size':24, 'font_color':'red', 'bottom':6})

    # Periode und Rechnungsnummer
    pr = workbook.add_format({'font_name':'Verdana', 'font_size':12, 'font_color':'red'})

    # Tabelle
    leer = workbook.add_format({'border':1, 'font_name':'Verdana', 'font_size':12})
    tt = workbook.add_format({'font_name':'Verdana', 'font_size':12, 'border':1})
    ttr = workbook.add_format({'font_name':'Verdana', 'font_size':12, 'border':1, 'align':'right', 'num_format':'0.00'})
    ttr_rot = workbook.add_format({'font_name':'Verdana', 'font_size':12, 'border':1, 'align':'right', 'num_format':'0.00', 'font_color':'red'})
    anzahl = workbook.add_format({'font_name':'Verdana', 'font_size':12, 'border':1, 'align':'right'})
    total = workbook.add_format({'bold':True, 'font_name':'Verdana', 'font_size':12, 'border':1, 'align':'right', 'num_format':'0.00'})



    worksheet = workbook.add_worksheet()

    # set widths
    worksheet.set_column('A:A', 9)
    worksheet.set_column('B:B', 36)
    worksheet.set_column('C:C', 18)
    worksheet.set_column('D:D', 18)


    ###static text
    # oberhalb erster doppellinie und Doppellinie
    worksheet.write('A1', 'Name des Betriebs', bg)
    worksheet.write('D1', 'Musterstrasse 1', rechts)
    worksheet.write('D2', 'PLZ und Ort', rechts)
    worksheet.write('D3', 'E-Mail: muster@example.ch', rechtsd)
    worksheet.write('A3', 'Telefon: 012 345 67 89', dd)
    worksheet.write('B3', '', dd)
    worksheet.write('C3', '', dd)

    # Mittelstück
    worksheet.write('A7', 'Kunde:', bg)
    worksheet.write('C7', familie, d)
    worksheet.write('C8', Adresse, d)
    worksheet.write('C9', PLZ_Ort, d)

    # Rechnung
    worksheet.write('A13', 'Rechnung', rechnung)
    worksheet.write('B13', '', dd)
    worksheet.write('C13', '', dd)
    worksheet.write('D13', '', dd)

    # Einleitung zu Tabelle
    worksheet.write('A16', 'Datum:', bg)
    t = time.strftime("%d.%m.%Y")
    worksheet.write('C16', t, d)
    worksheet.write('A18', 'Rechnungperiode:',d)
    worksheet.write('C18', periode, pr)
    worksheet.write('A19', 'Rechnungsnummer:', d)
    worksheet.write('C19', mom_rechnungsnr, pr)

    # Tabelle
    worksheet.write('A21', 'Anzahl', tt)
    worksheet.write('B21', 'Beschreibung', tt)
    worksheet.write('C21', 'Preis', tt)
    worksheet.write('D21', 'Betrag', tt)
    worksheet.write('A22', '', leer)
    worksheet.write('B22', 'Eierlieferungen', tt),
    worksheet.write('C22', '', leer)
    worksheet.write('D22', '', leer)

    
    worksheet.write('A23', menge_eier, anzahl)
    worksheet.write('B23', 'Stück', tt)
    worksheet.write('C23', price, ttr)
    worksheet.write('D23', '=A23*C23', ttr)

    for x in range(23, 29):
        for y in range(0,4):
            worksheet.write(x,y,'', leer)

    worksheet.write('C30', 'Total', bgu)
    worksheet.write('D30', '=SUM(D23:D29)', total)

    #Schluss
    worksheet.write('A35', 'Konditionen:', tt)
    worksheet.write('B35', '', leer)
    worksheet.write('C35', 'netto', ttr)
    worksheet.write('A36', 'zahlbar bis:', tt)
    worksheet.write('B36', '', leer)
    worksheet.write('C36', (datetime.date.today()+datetime.timedelta(365/12)).strftime('%d.%m.%Y'), ttr_rot)
    worksheet.write('A38', 'Herzlichen Dank für den geschätzten Auftrag', d)
    worksheet.write('A40', 'Mit freundlichen Grüssen', d)
    worksheet.write('A41', 'Name des Betriebs', pr)

    
    workbook.close()


    ################################################################# Version für Drucksheet ####################################################################################################


    ### fonts
    #default
    d = zu_drucken.add_format({'font_name':'Verdana', 'font_size':12})
    dd = zu_drucken.add_format({'font_name':'Verdana', 'font_size':12, 'bottom':6})

    # Betriebsge...
    bg = zu_drucken.add_format({'bold':True, 'font_name':'Verdana', 'font_size':12})
    bgu = zu_drucken.add_format({'bold':True, 'font_name':'Verdana', 'font_size':12, 'border':1})

    # rechtsbuendig
    rechts = zu_drucken.add_format({'font_name':'Verdana', 'font_size':12, 'align':'right'})
    rechtsd = zu_drucken.add_format({'font_name':'Verdana', 'font_size':12, 'align':'right', 'bottom':6})

    #Rechnung
    rechnung = zu_drucken.add_format({'font_name':'Verdana', 'font_size':24, 'font_color':'red', 'bottom':6})

    # Periode und Rechnungsnummer
    pr = zu_drucken.add_format({'font_name':'Verdana', 'font_size':12, 'font_color':'red'})

    # Tabelle
    leer = zu_drucken.add_format({'border':1, 'font_name':'Verdana', 'font_size':12})
    tt = zu_drucken.add_format({'font_name':'Verdana', 'font_size':12, 'border':1})
    ttr = zu_drucken.add_format({'font_name':'Verdana', 'font_size':12, 'border':1, 'align':'right', 'num_format':'0.00'})
    ttr_rot = zu_drucken.add_format({'font_name':'Verdana', 'font_size':12, 'border':1, 'align':'right', 'num_format':'0.00', 'font_color':'red'})
    anzahl = zu_drucken.add_format({'font_name':'Verdana', 'font_size':12, 'border':1, 'align':'right'})
    total = zu_drucken.add_format({'bold':True, 'font_name':'Verdana', 'font_size':12, 'border':1, 'align':'right', 'num_format':'0.00'})



    worksheet = zu_drucken.add_worksheet()

    # set widths
    worksheet.set_column('A:A', 9)
    worksheet.set_column('B:B', 36)
    worksheet.set_column('C:C', 18)
    worksheet.set_column('D:D', 18)


    ###static text
    # oberhalb erster doppellinie und Doppellinie
    worksheet.write('A1', 'Name des Betriebs', bg)
    worksheet.write('D1', 'Musterstrasse 1', rechts)
    worksheet.write('D2', 'PLZ und Ort', rechts)
    worksheet.write('D3', 'E-Mail: muster@example.com', rechtsd)
    worksheet.write('A3', 'Telefon: 012 345 67 89', dd)
    worksheet.write('B3', '', dd)
    worksheet.write('C3', '', dd)

    # Mittelstück
    worksheet.write('A7', 'Kunde:', bg)
    worksheet.write('C7', familie, d)
    worksheet.write('C8', Adresse, d)
    worksheet.write('C9', PLZ_Ort, d)

    # Rechnung
    worksheet.write('A13', 'Rechnung', rechnung)
    worksheet.write('B13', '', dd)
    worksheet.write('C13', '', dd)
    worksheet.write('D13', '', dd)

    # Einleitung zu Tabelle
    worksheet.write('A16', 'Datum:', bg)
    t = time.strftime("%d.%m.%Y")
    worksheet.write('C16', t, d)
    worksheet.write('A18', 'Rechnungperiode:',d)
    worksheet.write('C18', periode, pr)
    worksheet.write('A19', 'Rechnungsnummer:', d)
    worksheet.write('C19', mom_rechnungsnr, pr)

    # Tabelle
    worksheet.write('A21', 'Anzahl', tt)
    worksheet.write('B21', 'Beschreibung', tt)
    worksheet.write('C21', 'Preis', tt)
    worksheet.write('D21', 'Betrag', tt)
    worksheet.write('A22', '', leer)
    worksheet.write('B22', 'Eierlieferungen', tt),
    worksheet.write('C22', '', leer)
    worksheet.write('D22', '', leer)

    worksheet.write('A23', menge_eier, anzahl)
    worksheet.write('B23', 'Stück', tt)
    worksheet.write('C23', price, ttr)
    worksheet.write('D23', '=A23*C23', ttr)

    for x in range(23, 29):
        for y in range(0,4):
            worksheet.write(x,y,'', leer)

    worksheet.write('C30', 'Total', bgu)
    worksheet.write('D30', '=SUM(D23:D29)', total)

    #Schluss
    worksheet.write('A35', 'Konditionen:', tt)
    worksheet.write('B35', '', leer)
    worksheet.write('C35', 'netto', ttr)
    worksheet.write('A36', 'zahlbar bis:', tt)
    worksheet.write('B36', '', leer)
    worksheet.write('C36', (datetime.date.today()+datetime.timedelta(365/12)).strftime('%d.%m.%Y'), ttr_rot)
    worksheet.write('A38', 'Herzlichen Dank für den geschätzten Auftrag', d)
    worksheet.write('A40', 'Mit freundlichen Grüssen', d)
    worksheet.write('A41', 'Name des Betriebs', pr)



    ####################################################################### Ende beider sheets #####################################################################################3

    #Rechnungsnummer Liste updaten
    worksheet_rn.write(var+1, 0, mom_rechnungsnr)
    worksheet_rn.write(var+1, 1, 'Eier')
    worksheet_rn.write(var+1, 2, name)
    worksheet_rn.write(var+1, 3, menge_eier*price, zahlen_format)
    worksheet_rn.write(var+1, 4, t)
    
    
    #bookkeeping
    var +=1

# Rechnungsnummern Datei schöner darstellen
schieben = rechnungsnummern.add_format({'align':'right'})
worksheet_rn.set_column('E:E', None, schieben)

rechnungsnummern.close()
zu_drucken.close()



