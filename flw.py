import os
import shutil
import pandas as pd
import glob
import codecs


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
