import time
from typing import List

import xlsxwriter
import PySimpleGUI as sg

from bill import Bill
from billwriter import BillWriter
from checks import get_invalid_nicknames
from fonts import Fonts

information_path = "C:/Users/Adrian/Desktop/ground/"  # must end with a dash
output_path = "C:/Users/Adrian/Desktop/ground/"  # must end with a dash
root_file_name = "root_file.txt"


def get_all_nick_names() -> List[str]:
    file_name = information_path + root_file_name
    nicknames = []
    with open(file_name, 'r', encoding='utf-8') as f:
        current_line = f.readline().strip()
        while current_line not in [None, '']:
            nicknames.append(current_line)
            current_line = f.readline().strip()
    return nicknames


def get_amounts(nicknames: List[str]) -> List[int]:
    text_width = max(map(len, nicknames))
    layout = [
        [sg.Text(nickname, (text_width, 1)), sg.InputText()] for nickname in nicknames
    ]
    layout += [[sg.Button("Okay")]]
    window = sg.Window('Eier Rechnungen', layout)
    while True:
        event, values = window.read()
        if event is None:
            exit()
        elif event == 'Okay':
            break
    window.close()
    nbr_of_eggs = []
    for i in range(len(nicknames)):
        nbr_of_eggs.append(
            0 if values[i].strip() == '' else int(values[i])
        )
    return nbr_of_eggs


def main():
    period_str = input("Rechnungsperiode (steht auf den Rechnungen): ")
    bill_number = int(input("Erste Rechnungsnummer die verwendet werden darf: "))
    title_additive = input("Was soll am Schluss der Exel-Datei Namen stehen (z.B. '2-16' oder 'Mai-Aug 2016'): ")

    nicknames = get_all_nick_names()
    invalid_nicknames = get_invalid_nicknames(information_path, nicknames)
    if len(invalid_nicknames) != 0:
        print("Folgende Personen sind im " + root_file_name + " genannt, aber haben keine Adressinformationen:")
        print(", ".join(invalid_nicknames))
        input("\nDrücken Sie ENTER um ohne die genannten Kunden fortzufahren. Schliessen Sie das Fenster und "
              "korrigieren Sie das Problem, falls die genannten Kunden wichtig sind.")
    nbr_of_eggs = get_amounts(nicknames)
    bills = []
    for nickname, amount_of_eggs in zip(nicknames, nbr_of_eggs):
        if amount_of_eggs == 0:
            continue
        bills.append(Bill(period_str, bill_number, nickname, information_path, amount_of_eggs))
        bill_number += 1
        if bill_number == 1000:
            bill_number = 1

    write_individual_bill_excels(bills, title_additive)
    write_to_print_excel(bills)
    write_overview_excel(bills)


def write_overview_excel(bills: List[Bill]):
    overview_workbook = xlsxwriter.Workbook(output_path + "Rechnungsnummern_Liste.xlsx")
    worksheet = overview_workbook.add_worksheet()
    worksheet.write('A1', 'Nr.')
    worksheet.write('C1', 'Name')
    worksheet.write('D1', 'Betrag')
    worksheet.write('E1', 'Rechnungsdatum (geschrieben)')
    numbers_format = overview_workbook.add_format({'num_format': '0.00'})
    for i, bill in enumerate(bills, 1):
        worksheet.write(i, 0, bill.bill_nbr)
        worksheet.write(i, 1, 'Eier')
        worksheet.write(i, 2, bill.customer_nickname)
        worksheet.write(i, 3, bill.nbr_of_eggs * Bill.price_per_egg, numbers_format)
        worksheet.write(i, 4, time.strftime("%d.%m.%Y"))
    overview_workbook.close()


def write_to_print_excel(bills: List[Bill]):
    to_print_workbook = xlsxwriter.Workbook(output_path + "zum_drucken.xlsx")
    fonts = Fonts(to_print_workbook)
    bill_writer = BillWriter(fonts)
    for bill in bills:
        worksheet = to_print_workbook.add_worksheet()
        bill_writer.write_bill(worksheet, bill)
    to_print_workbook.close()


def write_individual_bill_excels(bills: List[Bill], title_additive: str):
    for bill in bills:
        excel_title = bill.get_bill_nbr_string() + " Eier " + bill.customer_nickname + " " + title_additive + ".xlsx"
        workbook = xlsxwriter.Workbook(output_path + excel_title)
        fonts = Fonts(workbook)
        bill_writer = BillWriter(fonts)
        worksheet = workbook.add_worksheet()
        bill_writer.write_bill(worksheet, bill)
        workbook.close()


if __name__ == "__main__":
    main()
