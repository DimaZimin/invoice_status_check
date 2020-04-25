import os
import shutil
import pandas as pd
import glob
import codecs


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
