import os
import shutil
import pandas as pd
import glob
import codecs
from openpyxl import load_workbook
from datetime import datetime



bw_statuses = {'R1': 'Returned', 'C1': 'Cancelled',
                   '20': 'Unprocessed E-invoice', '4': 'Cancelled Invoices',
                   '3': 'Transfered to GFIS', '2': 'Ready to transfer',
                   '1': 'Sent for approval', '0': 'Unprocessed'}


gfis_data = {}
inv_basware_status = {}
path_file_gfis = glob.glob('gfis/*.xlsx')



def decoding():
    for file in glob.glob('basware/*.csv'):
        with open(file, encoding='UTF-16LE') as inp_f:
            with codecs.open(f'basware/{file.split("/")[1].split(".")[0]}'+'dec.csv', 'w', encoding='utf-8') as out_f:
                shutil.copyfileobj(inp_f, out_f)

def combine_to_excel():
    all_data = pd.DataFrame()
    for f in glob.glob('basware/*dec.csv'):
        df = pd.read_csv(f, sep='\t')
        all_data = all_data.append(df, ignore_index=True)
        os.remove(f)
    all_data.to_excel('all_invoices.xlsx', index=False)


def remove_row():
    for file in glob.glob('gfis/*.xlsx'):
        gfis_excel = load_workbook(filename=f'{file}')
        gfis_sheet = gfis_excel.active
        for col in gfis_sheet.iter_rows(max_row=1, values_only=True):
            print('Checking GFIS files...')
            if None in col:
                print(f'File: {file} - 1st row has been deleted')
                gfis_sheet.delete_rows(1)
                gfis_excel.save(file)
            else:
                print(f'No changes in {file}')
    print('GFIS files are OK.')

def pdate_column(file):
    gfis_excel = load_workbook(filename=f'{file}')
    gfis_sheet = gfis_excel.active
    for col in gfis_sheet.iter_rows(max_row=1):
        return len(col)

def retrieve_invoice_data():
    for file in glob.glob('gfis/*.xlsx'):
        try:
            gfis_excel = load_workbook(filename=f'{file}')
            gfis_sheet = gfis_excel.active
        except FileNotFoundError:
            print('File *.xlsx in <gfis> not found')
        else:
            gfis_invoices = [str(invoice[0]) for invoice in
                             gfis_sheet.iter_rows(min_row=2, min_col=2, max_col=2, values_only=True)]
            gfis_schedule = [datetime.strftime(schedule[0], '%Y-%m-%d') if schedule[0] else 'no data'
                             for schedule in gfis_sheet.iter_rows(min_row=2, min_col=13, max_col=13, values_only=True)]
            gfis_payment = [pdate[0] for pdate in gfis_sheet.iter_rows(min_row=2, min_col=pdate_column(file),
                                                                       max_col=pdate_column(file), values_only=True)]
            pmt_dates = [datetime.strftime(dt, '%Y-%m-%d') if dt is not None else 'not paid' for dt in gfis_payment]

            for inv, sch, dat in zip(gfis_invoices, gfis_schedule, pmt_dates):
                gfis_data[inv] = sch, dat


def decoding_flow():
    for file in glob.glob('flow/*.csv'):
        with open(file, encoding='UTF-16LE') as inp_f:
            with codecs.open(f'flow/{file.split("/")[1].split(".")[0]}'+'dec.csv', 'w', encoding='utf-8') as out_f:
                shutil.copyfileobj(inp_f, out_f)


def combine_to_excel_flow():
    all_data_flow = pd.DataFrame()
    for f in glob.glob('flow/*dec.csv'):
        df = pd.read_csv(f, sep='\t')
        all_data_flow = all_data_flow.append(df, ignore_index=True)
        os.remove(f)
    all_data_flow.to_excel('flow/flow_invoices.xlsx', index=False)


def get_approver(invoice):
    for i in enumerate(flow_inv):
        if str(invoice) == str(i[1]):
            return approvers[i[0]], date_sent[i[0]]

def inv_status():
    for inv in requested_invoices:
        if inv in gfis_data.keys():
            inv_basware_status[inv] = f'Scheduled due {gfis_data[inv][0]}, payment: {gfis_data[inv][1]}'
        elif inv in all_inv.keys():
            inv_basware_status[inv] = bw_statuses[str(all_inv[inv])]
        else:
            inv_basware_status[inv] = 'Missing'



def write_status():
    for i, (k, v) in enumerate(inv_basware_status.items()):
        try:
            if v == bw_statuses['1']:
                statcheck_sheet[f'B{i + 2}'] = f'{v}'
                statcheck_sheet[f'C{i + 2}'] = f'to {get_approver(k)[0]}, on {get_approver(k)[1]}'
                check_invoices.save(filename='check_invoices.xlsx')
            elif v == bw_statuses['3']:
                statcheck_sheet[f'B{i + 2}'] = f'{v}'
                statcheck_sheet[f'C{i + 2}'] = f'scheduled due {gfis_data[k][0]}'
                statcheck_sheet[f'D{i + 2}'] = gfis_data[k][1]
                check_invoices.save(filename='check_invoices.xlsx')
            else:
                statcheck_sheet[f'B{i + 2}'] = f'{v}'
                check_invoices.save(filename='check_invoices.xlsx')
        except TypeError and KeyError:
            statcheck_sheet[f'B{i + 2}'] = f'{v}'
            statcheck_sheet[f'C{i + 2}'] = 'Data is missing. Please check manually. '

if __name__ == '__main__':
    print('################### Basware/GFIS invoice status checking.##########################')
    print('============== Copyright © 2020 Dmytro Zimin. All Rights Reserved. ============')
    print('################################################################################\n\n')
    print('***Please make sure you have read the instruction manual before using this program***')
    run = input('To proceed press <y> or any other button to exit\n')
    if run.lower()[0] == 'y':
        try:
            check_invoices = load_workbook(filename='check_invoices.xlsx')
            statcheck_sheet = check_invoices.active
        except FileNotFoundError:
            print('Unable to find <check_invoices.xlsx> in root directory')
            input('Press any key to exit')
            exit()
        else:
            requested_invoices = [str(inum[0]) for inum in
                                  statcheck_sheet.iter_rows(min_row=2, min_col=1, max_col=1, values_only=True)]
        print('Encoding files...')
        try:
            decoding()
        except FileNotFoundError:
            print('Unable to find *.csv files in <basware> directory')
            input('Press any key to exit')
            exit()
        else:
            print('Encoding completed...')
        print('Merging files...')
        combine_to_excel()
        print('Merging files completed!')
        remove_row()
        print('Encoding flow files')
        decoding_flow()
        print('Encoding completed')
        print('Rewriting <flow> to .xlsx')
        combine_to_excel_flow()
        try:
            path_file_flow = glob.glob('flow/*.xlsx')[0]
            inflow_invoices = load_workbook(filename=f'{path_file_flow}')
            inflow_sheet = inflow_invoices.active
        except:
            print('Retrieving invoice data. Please wait...')
            print('Error code 0. Please check if there is a *.xlxs file in <flow> directory')
            input('Press any key to exit')
            exit()
        else:
            approvers = [str(approver[0]) for approver in
                         inflow_sheet.iter_rows(min_row=2, min_col=13, max_col=13, values_only=True)]
            flow_inv = [str(inv[0]) for inv in
                        inflow_sheet.iter_rows(min_row=2, min_col=7, max_col=7, values_only=True)]
            date_sent = [dt[0].split(' ')[0] for dt in
                         inflow_sheet.iter_rows(min_row=2, min_col=15, max_col=15, values_only=True)]
        retrieve_invoice_data()
        try:
            all_invoices = load_workbook(filename='all_invoices.xlsx')
            allinv_sheet = all_invoices.active
        except FileNotFoundError:
            print('Cannot find <check_invoices.xlsx> file. Please read instructions and try again.')
            input('Press any key to exit')
            exit()
        else:
            all_inv_nr = [str(invoice[0]) for invoice in
                          allinv_sheet.iter_rows(min_row=2, min_col=9, max_col=9, values_only=True)]
            all_inv_st = [str(stat[0]) for stat in
                          allinv_sheet.iter_rows(min_row=2, min_col=2, max_col=2, values_only=True)]
            all_inv = {key: value for (key, value) in zip(all_inv_nr, all_inv_st)}
            inv_status()
            write_status()
            input('Statuses have been added to check_invoices.xlsx! Press any key to exit')
            exit()






