from functions import last_column, retrieve_gfis_data, data_from_combined, data_from_flow,\
                       get_inv_status, write_status
from datafile import DataFile
import glob

if __name__ == '__main__':
    print('################### Basware/GFIS invoice status checking.##########################')
    print('========================= Created by Dmytro Zimin =============================')
    print('################################################################################\n\n')
    print('***Please make sure you have read the instruction manual before using this program***')
    run = input('To proceed press <y> or any other button to exit\n')
    if run.lower()[0] == 'y':
        print('combining basware *.csv files to excel ')
        for file in glob.glob('basware/*.csv'):
            basware_file = DataFile(file)
            basware_file.combine_to_excel('basware', 'basware/combined.xlsx')
        print('combining flow *.csv files to excel ')
        for file in glob.glob('flow/*.csv'):
            flow_file = DataFile(file)
            flow_file.combine_to_excel('flow', 'flow/flow.xlsx')
        print('retrieving data...')
        retrieve_gfis_data('gfis/*.xlsx')
        data_from_combined('basware/combined.xlsx')
        data_from_flow('flow/flow.xlsx')
        get_inv_status('check_invoices.xlsx')
        print('writing statuses...')
        write_status('check_invoices.xlsx')
        print('deleting temporary files...')
        DataFile.remove_temporary_files('basware/combined.xlsx')
        DataFile.remove_temporary_files('flow/flow.xlsx')
        input('Statuses have been added to check_invoices.xlsx! Press any key to exit')
        exit()
    else:
        exit()



