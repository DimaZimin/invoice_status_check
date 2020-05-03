import os
import shutil
import pandas as pd
import glob
import codecs


def decode_basware_files():  # TODO: delete if not needed
    for path in glob.glob('basware/*.csv'):
        with open(path, encoding='UTF-16LE') as original_file:
            new_name = path.replace('.csv', 'dec.csv')
            with codecs.open(new_name, 'w', encoding='utf-8') as new_file:
                shutil.copyfileobj(original_file, new_file)


def read_file(path):
    with open(path, encoding='UTF-16LE') as f:
        return pd.read_csv(f, sep='\t')


def combine_to_excel(input_directory: str, output_file: str) -> None:
    parsed = [read_file(path) for path in glob.glob(f'{input_directory}/*.csv')]
    merged = pd.concat(parsed)
    merged.to_excel(output_file, index=False)
