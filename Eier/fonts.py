import xlsxwriter
from xlsxwriter.format import Format


class Fonts:
    d: Format
    dd: Format
    bg: Format
    bgu: Format
    rechts: Format
    rechtsd: Format
    rechnung: Format
    pr: Format
    leer: Format
    tt: Format
    ttr: Format
    ttr_rot: Format
    anzahl: Format
    total: Format

    def __init__(self, workbook: xlsxwriter.Workbook):
        # fonts
        # default
        self.d = workbook.add_format({'font_name': 'Verdana', 'font_size': 12})
        self.dd = workbook.add_format({'font_name': 'Verdana', 'font_size': 12, 'bottom': 6})

        # Betriebsge...
        self.bg = workbook.add_format({'bold': True, 'font_name': 'Verdana', 'font_size': 12})
        self.bgu = workbook.add_format({'bold': True, 'font_name': 'Verdana', 'font_size': 12, 'border': 1})

        # rechtsbuendig
        self.rechts = workbook.add_format({'font_name': 'Verdana', 'font_size': 12, 'align': 'right'})
        self.rechtsd = workbook.add_format({'font_name': 'Verdana', 'font_size': 12, 'align': 'right', 'bottom': 6})

        # Rechnung
        self.rechnung = workbook.add_format({'font_name': 'Verdana', 'font_size': 24, 'font_color': 'red', 'bottom': 6})

        # Periode und Rechnungsnummer
        self.pr = workbook.add_format({'font_name': 'Verdana', 'font_size': 12, 'font_color': 'red'})

        # Tabelle
        self.leer = workbook.add_format({'border': 1, 'font_name': 'Verdana', 'font_size': 12})
        self.tt = workbook.add_format({'font_name': 'Verdana', 'font_size': 12, 'border': 1})
        self.ttr = workbook.add_format(
            {'font_name': 'Verdana', 'font_size': 12, 'border': 1, 'align': 'right', 'num_format': '0.00'})
        self.ttr_rot = workbook.add_format(
            {'font_name': 'Verdana', 'font_size': 12, 'border': 1, 'align': 'right', 'num_format': '0.00',
             'font_color': 'red'})
        self.anzahl = workbook.add_format({'font_name': 'Verdana', 'font_size': 12, 'border': 1, 'align': 'right'})
        self.total = workbook.add_format(
            {'bold': True, 'font_name': 'Verdana', 'font_size': 12, 'border': 1, 'align': 'right',
             'num_format': '0.00'})

