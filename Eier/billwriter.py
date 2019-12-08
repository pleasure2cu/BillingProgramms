import datetime
import time

from xlsxwriter.worksheet import Worksheet

from bill import Bill
from fonts import Fonts


class BillWriter:
    fonts: Fonts

    def __init__(self, fonts: Fonts):
        self.fonts = fonts

    def write_bill(self, worksheet: Worksheet, bill: Bill):
        self._static_upper_part(worksheet)
        self._write_postal_information(worksheet, bill)
        self._write_bill_header(worksheet)
        self._write_bill_meta_info(worksheet, bill)
        self._write_table(worksheet, bill)
        self._write_end(worksheet)

    def _static_upper_part(self, worksheet: Worksheet):
        # set widths
        worksheet.set_column('A:A', 9)
        worksheet.set_column('B:B', 36)
        worksheet.set_column('C:C', 18)
        worksheet.set_column('D:D', 18)

        worksheet.write('A1', 'Name des Betriebs', self.fonts.bg)
        worksheet.write('D1', 'Musterstrasse 1', self.fonts.rechts)
        worksheet.write('D2', 'PLZ und Ort', self.fonts.rechts)
        worksheet.write('D3', 'E-Mail: muster@example.ch', self.fonts.rechtsd)
        worksheet.write('A3', 'Telefon: 012 345 67 89', self.fonts.dd)
        worksheet.write('B3', '', self.fonts.dd)
        worksheet.write('C3', '', self.fonts.dd)

    def _write_postal_information(self, worksheet: Worksheet, bill: Bill):
        worksheet.write('A7', 'Kunde:', self.fonts.bg)
        worksheet.write('C7', bill.customer_name, self.fonts.d)
        worksheet.write('C8', bill.customer_address, self.fonts.d)
        worksheet.write('C9', bill.customer_zip_place, self.fonts.d)

    def _write_bill_header(self, worksheet: Worksheet):
        worksheet.write('A13', 'Rechnung', self.fonts.rechnung)
        worksheet.write('B13', '', self.fonts.dd)
        worksheet.write('C13', '', self.fonts.dd)
        worksheet.write('D13', '', self.fonts.dd)

    def _write_bill_meta_info(self, worksheet: Worksheet, bill: Bill):
        worksheet.write('A16', 'Datum:', self.fonts.bg)
        t = time.strftime("%d.%m.%Y")
        worksheet.write('C16', t, self.fonts.d)
        worksheet.write('A18', 'Rechnungperiode:', self.fonts.d)
        worksheet.write('C18', bill.period_str, self.fonts.pr)
        worksheet.write('A19', 'Rechnungsnummer:', self.fonts.d)
        worksheet.write('C19', bill.get_bill_nbr_string(), self.fonts.pr)

    def _write_table(self, worksheet: Worksheet, bill: Bill):
        worksheet.write('A21', 'Anzahl', self.fonts.tt)
        worksheet.write('B21', 'Beschreibung', self.fonts.tt)
        worksheet.write('C21', 'Preis', self.fonts.tt)
        worksheet.write('D21', 'Betrag', self.fonts.tt)
        worksheet.write('A22', '', self.fonts.leer)
        worksheet.write('B22', 'Eierlieferungen', self.fonts.tt),
        worksheet.write('C22', '', self.fonts.leer)
        worksheet.write('D22', '', self.fonts.leer)

        worksheet.write('A23', bill.nbr_of_eggs, self.fonts.anzahl)
        worksheet.write('B23', 'Stück', self.fonts.tt)
        worksheet.write('C23', Bill.price_per_egg, self.fonts.ttr)
        worksheet.write('D23', '=A23*C23', self.fonts.ttr)

        for x in range(23, 29):
            for y in range(0, 4):
                worksheet.write(x, y, '', self.fonts.leer)

        worksheet.write('C30', 'Total', self.fonts.bgu)
        worksheet.write('D30', '=SUM(D23:D29)', self.fonts.total)

    def _write_end(self, worksheet: Worksheet):
        worksheet.write('A35', 'Konditionen:', self.fonts.tt)
        worksheet.write('B35', '', self.fonts.leer)
        worksheet.write('C35', 'netto', self.fonts.ttr)
        worksheet.write('A36', 'zahlbar bis:', self.fonts.tt)
        worksheet.write('B36', '', self.fonts.leer)
        worksheet.write('C36', (datetime.date.today() + datetime.timedelta(365 / 12)).strftime('%d.%m.%Y'), self.fonts.ttr_rot)
        worksheet.write('A38', 'Herzlichen Dank für den geschätzten Auftrag', self.fonts.d)
        worksheet.write('A40', 'Mit freundlichen Grüssen', self.fonts.d)
        worksheet.write('A41', 'Name des Betriebs', self.fonts.pr)

