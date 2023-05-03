import time
import PyPDF2
import glob
import os
import re
import pickle

import pandas as pd

from datetime import datetime
from random import randint


# currently for 1112 only. Will create ad-hoc if demand
class InvMergerBot:

    def __init__(self):

        self.pickle_fpath = r'C:\Users\V Song\Desktop\_PYTHONSCRIPTS\_picklenv\data\envs.pickle'
        self.envar = self.unpickle_envar()

        self.y_per = 2023
        self.m_per = 1
        
        # default selected prog
        # 1168 and 1112 
        self.selected_prog = '1112'

        self.main_dir = fr'{self.envar["FROM_HOST"]}\_pdfMERGER\{self.selected_prog}\{str(self.y_per)}{str(self.m_per).zfill(2)}'
        
        
        # for 1112 merger  Invoice + TS + Expense pre-approval support
        self.member_list = [os.path.basename(x) for x in glob.glob(fr'{self.main_dir}\TS\*')]
        
        self.attach_type = [
                            'TS',
                            'EXPENSES',
                            ]
        
        self.menu_choices = {
            'QUIT': exit,
            'CHANGE PROG': self.change_prog,
            'CHANGE PER' : self.change_per,
            'START MERGE 1112' : self.start_merge_elevtwel,
            'START MERGE 1168' : self.start_merge_elevsixty,
 
        }
        
        self.main_menu()
        
    def main_menu(self):

        while True:
            drawn_num = randint(0,1)
            choose_message = ''

            if drawn_num == 1:
                choose_message = 'It\'s time to choose..'
            else:
                choose_message = 'TIme...too TCHooOOoOSee...'
                
            self.welcome_message()

            print('---------MAIN MENU ---------')
            for i, v in enumerate(self.menu_choices.keys()):
                print(f'{i} - {v}')

            user_input = input(f'{choose_message}\n')

            try:
                self.menu_choices[list(self.menu_choices.keys())[int(user_input)]]()

            except Exception as e:
                print(e)
                
    def unpickle_envar(self):
        with open(self.pickle_fpath, 'rb') as p:
            return pickle.load(p)
        
    def welcome_message(self):

        print('------------------------------------')
        print(f'Hello I am Invoice Merger')
        print(f'Current selected prog/proj: {self.selected_prog}')
        print(f'Today is: {datetime.now()}')
        print(f'Current PER set to: {self.y_per} {self.m_per}\n')
        print(f'Current dir set to: {self.main_dir}')

    def start_merge_elevsixty(self):
        wo_folderlist = list()
        
        for wo in os.listdir(self.main_dir):
            wo_folderlist.append(wo)
        
        for folder in wo_folderlist:
            print(f'Starting wo for: {folder}\non period {self.m_per}{self.y_per}')
            time.sleep(0.6)
                        
            # -1, hardcoded, usually the approved pdf support will be listed last 
            # as long as it keeps following the naming convention. ex: "Aaron Maez_APPROVED"
            # returns a list based on each folder dir 
            # merge with 
            for x in os.listdir(os.path.join(self.main_dir, folder)):
                merger = PyPDF2.PdfMerger(strict=False)
                approved_dir = os.path.join(self.main_dir, folder, os.listdir(os.path.join(self.main_dir, folder))[-1])
                approved_file = PyPDF2.PdfReader(open(approved_dir, 'rb'))
                
                if re.match('\d{7}.pdf', x):
                    inv_file = PyPDF2.PdfReader(open(os.path.join(self.main_dir, folder, x), 'rb'))
                    merger.append(inv_file)
                    time.sleep(0.5)
                    print(f'{approved_dir} will append to invoice: {x}')
                    merger.append(approved_file)
                    time.sleep(0.5)
                    merger.write(os.path.join(self.main_dir, x))
                    time.sleep(1)
                    del merger
                    del inv_file   
        
        return

    def start_merge_elevtwel(self):
        self.eleventwelve_create_workorder()
        
        time.sleep(1)
        
        self.eleventwelve_start_wo()
        return

    def eleventwelve_create_workorder(self):

        def prepare_regex_list():
            # gets the reject pattern based on self.member_list in TS dir
            result = ''
            for n in enumerate(self.member_list):
                if n[0] != len(self.member_list)-1:
                    result += f'{n[1]}|'
                else:
                    result += f'{n[1]}'

            return result

        def pdf_to_str(apdf):
            # gets opened pdf obj from passed argument
            # uses PyPDF2 to parse strings
            # and loop pdfdoc into a reader to get str values
            # returns result
            pdfdoc = PyPDF2.PdfReader(apdf)
            result = ''

            for i in range(len(pdfdoc.pages)):
                current_page = pdfdoc.pages[i]
                result += current_page.extract_text()

            return result


        work_order = {
            'PRONUM': [],
            'INVNUM': [],
            'EMPNAME': []
        }

        regex_member_list_pattern = prepare_regex_list()

        regex_project_pattern = '1112.\d{3}.\d{3}|' \
                                '1112.\d{3}.00R|' \
                                '1112.00E.\d{3}|' \
                                '1112.000.00M'    

        regex_invoice_pattern = 'Invoice No:\s*\d{7}'

        print(self.member_list)

        for invoice_dir in glob.glob(fr'{self.main_dir}\INV\*.pdf'):

            pdf_invoice = open(invoice_dir, mode='rb')
            invoice_parsed = pdf_to_str(pdf_invoice)

            project_number_found = re.search(regex_project_pattern, invoice_parsed).group(0)
            invoice_number_found = re.search(regex_invoice_pattern, invoice_parsed).group(0)[-7:]

            if re.search(regex_member_list_pattern, invoice_parsed) is None:
                print(f'invoice {os.path.basename(invoice_dir)} not in Brenton\'s Group')
                
            else:
                temp_dup_checker = []
                
                for empname in re.finditer(regex_member_list_pattern, invoice_parsed):
                    # temp_dup_checker.append(empname.group(0)) 
                    # if empname.group(0) not in temp_dup_checker else False
                    if empname.group(0) not in temp_dup_checker:
                        temp_dup_checker.append(empname.group(0))
                        
                # after iterating each regex/ and checked for duplicates
                # append the member project/inv number in dict work_order
                for name in temp_dup_checker:
                    work_order['EMPNAME'].append(name)
                    work_order['PRONUM'].append(project_number_found)
                    work_order['INVNUM'].append(invoice_number_found)

        # converts dictionary into dataframe
        df = pd.DataFrame.from_dict(work_order)
        print('workorder output')
        print(df)
        
        # save as csv
        df.to_csv(fr'{self.main_dir}\wo.csv', index=False)
        time.sleep(0.5)
        print(fr'Work Order created at: {self.main_dir}')
    
    def eleventwelve_start_wo(self):
        # main() will pick up the invoices in INV dir,
        # use the wo.csv as a mapping to direct which PBC TS to attach
        # the function overwrites the appended INVOICES in the INV dir
        # make sure you have a backup on the 1112 invoices before runnning this function

        wo_cols = (
                   'PRONUM',
                   'INVNUM',
                   'EMPNAME',
                   )

        wo = pd.read_csv(fr'{self.main_dir}\wo.csv', names=wo_cols, skiprows=1)

        # creates a temp list of invoices into invoice_wo
        # invoice_wo does not contain duplicated invoice number
        # to avoid including the same TS group twice or more
        invoice_wo = wo['INVNUM'].drop_duplicates()

        # REMINDER: this is the function that merges invoices to signed timesheets
        # the loop is based on the invoice_wo
        # which contains unique in num values
        for invnum in invoice_wo:
            time.sleep(1)
            inv_file = PyPDF2.PdfReader(
                open(fr'{self.main_dir}\INV\{invnum}.pdf',
                     'rb'),  strict=False)

            t = wo.loc[wo["INVNUM"] == int(invnum)]
            print(f'Work order for invoice {invnum}:\n')
            print(t)
            print(f'\n')

            time.sleep(1)
            
            for tsname in t['EMPNAME']:
                merger = PyPDF2.PdfMerger(strict=False)
                merger.append(inv_file)
                print(f'starting {tsname} merging')

                for att_type in self.attach_type:
                    
                    for fdir in glob.glob(fr'{self.main_dir}\{att_type}\{tsname}\*.pdf'):
                        print(f'merging {invnum} with {os.path.basename(fdir)}')
                        ts_file = PyPDF2.PdfReader(open(fdir, 'rb'))
                        merger.append(ts_file)
                        del ts_file
                        time.sleep(1)

                # for fdir in glob.glob(fr'{self.main_dir}\EXPENSES\{tsname}\*.pdf'):
                #     print(f'merging {invnum} with {os.path.basename(fdir)}')
                #     expense_file = PyPDF2.PdfFileReader(open(fdir, 'rb'))
                #     merger.append(expense_file)
                #     del expense_file
                #     time.sleep(1)

                merger.write(fr'{self.main_dir}\INV\{invnum}.pdf')
                time.sleep(1)
                print(f'Merged invoice {invnum} saved')
                print(f'\n')
                print(f'\n')
                time.sleep(1)

            del merger
            del inv_file

    def change_prog(self):
        self.selected_prog = input(f'Please enter new prog num in XXXX format:\n')
        return
    
    def change_per(self):
        self.m_per = input(f'Please enter new month in MM format:\n')
        print(f'Month period set to: {self.m_per}\n')
        time.sleep(0.5)
        
        self.y_per = input(f'Please enter new year in YYYY format:\n')
        print(f'Year period set to: {self.y_per}\n')
        time.sleep(0.5)
        print('Returning to main menu...')
        
        # refreshes main dir with new m and y per 
        self.main_dir = fr'{self.envar["FROM_HOST"]}\_pdfMERGER\{self.selected_prog}\{str(self.y_per)}{str(self.m_per).zfill(2)}'
        
        return

def main():
    INVMERGER = InvMergerBot()


if __name__ == '__main__':
    main()
