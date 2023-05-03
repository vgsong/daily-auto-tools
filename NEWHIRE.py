import os
import pandas as pd
import re

from datetime import datetime, timedelta
from win32com.client import Dispatch


class NewHire:
    def __init__(self):
        
        # CONSTANTS
        self.upload_fpath = fr'C:\Users\V Song\Documents'
        
        # to capture relevant data from Cherie's new hire email form
        self.regex_dict = {
            'emp_name': '(?<=Employee Name: )((.|\n)*)(?=Organization:)',
            'emp_number': '(?<=Employee Number: )\d{4}',
            'on_board': '(?<=PTO Designations(.))((.|\n)*)(?=Billable Projects:)',
            'org' :  '(?<=Organization: Cordoba)[\s\S]*[A-Z]{2}'

        }

        # mappings for cordoba org
        self.org_map = {
            'CW': 10,
            'LA': 15,
            'SA': 20, 
            'SD': 25,
            'SF': 35,
            'SA': 40,
            'ON': 45,
        }

        # mappings for cordoba sector
        self.sector_mapp = {
            'Corporate Services': 10,
            'Education': 15,
            'Energy': 20,
            'Transportation': 25,
            'Water': 30,
        }

        # wbs1 onboarding projects
        self.onboarding_fulltime = {
            
            'Training/Safety': '0000.000.105',
            'Vacation': '0000.000.200',
            'Sick Leave': '0000.000.205',
            'Holiday': '0000.000.210',
            'Bereavement': '0000.000.215',
            'Jury Duty': '0000.000.220',
            # 'Business Proposal': '0000.000.105',
            # 'Energy Proposals 2021': '0000.000.105',
            # 'Business Development': '0000.000.105',
            # 'Recruiting': '0000.000.105',
            # 'Emergency-PTO': '0000.000.105',
            # 'Business Development': '0000.000.105',
            # 'Admin': '0000.000.105',
            # 'Admin-WorkCare': '0000.000.105',
            
        }

        # col names for upload
        self.df_colnames = [
                            'WBS1',
                            'WBS2',
                            'WBS3',
                            'EMPID',
                            ]
        
        # MUTABLES
        self.mail_items = None # obtains mail items from outlook dispatch last 5 days 
        self.wo = None # var for cleaned data as workorder(clean) format
        
        self.output_fpath = r'C:\Users\V Song\Documents'
        
        self.class_main()

    def class_main(self):
        self.mail_items = self.get_filtered_outlook_items(2) # obtains mail items from outlook dispatch last 5 days 
        self.wo = self.get_newhire_details() # var for cleaned data as workorder(clean) format

    def get_filtered_outlook_items(self, n_days):

        olapp = Dispatch('outlook.application')
        mapi = olapp.GetNamespace('MAPI').Folders('first.last@company.com').Folders('Inbox').Folders('_NEWHIRE').Items
        received_dt = datetime.now() - timedelta(days=n_days)
        received_dt = received_dt.strftime('%m/%d/%Y %H:%M %p')
        messages = mapi.Restrict("[ReceivedTime] >= '" + received_dt + "'")
        newhire_mailitems = list()

        for x in messages:
            if x.Unread:
                newhire_mailitems.append(x)

        return newhire_mailitems

    def get_newhire_details(self):

        wo = dict()

        for mail in self.mail_items:
            
            try:
                emp_name = re.search(self.regex_dict['emp_name'], mail.body).group(0).strip()
                emp_number = re.search(self.regex_dict['emp_number'], mail.body).group(0).strip()
                org_desc = re.search(self.regex_dict['org'], mail.body).group(0).strip()[:-2]
                sector = re.search(self.regex_dict['org'], mail.body).group(0).strip()[-2:]
                
                wo[emp_number] = list()
                wo[emp_number].append(emp_name.strip())
                wo[emp_number].append(org_desc.strip())
                wo[emp_number].append(sector.strip())
            
            except Exception as e:
                print(f'ERROR FOUND: {e}')
                print(f'NEW HIRE IGNORED:')

        
        print(wo)
        
        return wo
    
    def wo_to_csv(self):
        
        df = pd.DataFrame(columns=self.df_colnames)
        print(f'Total wo count: {len(self.wo)}')
        
        for i, va in enumerate(self.wo.items()):
            # print(i, va)
            
            to_append = {
                    'WBS1' : list(),
                    'WBS2' : list(),
                    'WBS3' : list(),
                    'EMPID': list(),
                     }
            
            df2 = pd.DataFrame(columns=self.df_colnames)
            
            for projname, projnum in self.onboarding_fulltime.items():
            
                try:
                    sector_number = self.sector_mapp[va[1][1]]
                    proj_sector_org = str(projnum[-3:]) + str(self.sector_mapp[va[1][1]]) + str(self.org_map[va[1][2]])
                    emp_id = va[0]
                         
                    to_append['WBS1'].append(projnum)
                    to_append['WBS2'].append(sector_number)
                    to_append['WBS3'].append(proj_sector_org)
                    to_append['EMPID'].append(emp_id)
                    
                except Exception as e:
                    print(f'ERROR FOUND: {e}')
                    to_append['WBS1'].append('ERROR!')
                    to_append['WBS2'].append('ERROR!')
                    to_append['WBS3'].append('ERROR!')
                    to_append['EMPID'].append('ERROR!')
                    

            df2 = pd.DataFrame(to_append, columns=self.df_colnames)
            # print(df.append(to_append, ignore_index=True))
            df = pd.concat([df, df2])
     
     
        tstamp = datetime.now().strftime('%Y%m%d_%M%S')
        print(df)
        
        df.to_csv(fr'{self.upload_fpath}\NEWHIRE_{tstamp}.csv', index=False, header=None)
        
        return
    
def main():
    nh = NewHire()
    nh.wo_to_csv()

if __name__ == '__main__':
    main()
