import pandas as pd
import pickle
import os
import numpy as np


pd.set_option('display.max_rows', 1000)


class DeltekReporter:
    def __init__(self):
        self.pickle_fpath = r'C:\Users\V Song\Desktop\_PYTHONSCRIPTS\_picklenv\data\envs.pickle'
        self.envar = self.unpickle_envar()

        self.proj_filename = 'PROJECT_LIST_EXPORT.csv'
        self.fromhost_fpath = self.envar['FROM_HOST']
        
        self.main_fpath = os.path.join(self.envar['FROM_HOST'], '_FROMOUTLOOK', self.proj_filename)
        
        self.csv_data = pd.read_csv(self.main_fpath)
        
        self.type_reports = {
            'general' : '',
            
        }

    def unpickle_envar(self):
        with open(self.pickle_fpath, 'rb') as p:
            return pickle.load(p)
        
    def load_proj_df(self):
        print(self.main_fpath)
        df =  self.csv_data
        # df = df.dropna(axis=0, subset=['detail_Phase'])
        df = df[df['detail_Phase'].str.contains(' ')]
        
        
        print(df)
        

        
        # filter1 = df['detail_Phase'].dropna(how='all')
        # clean_df = df[filter1]
        # print(clean_df)

def main():
    dr = DeltekReporter()
    dr.load_proj_df()
   

if __name__ == '__main__':
    main()
