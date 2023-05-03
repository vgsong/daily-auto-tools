import pickle
import pandas as pd

from win32com.client import Dispatch
from vs_classes.itd_reader import ITDReader


class UnbilledCheck(ITDReader):
    def __init__(self, **kwargs):
        super(UnbilledCheck, self).__init__()
        self.DIR_DICT = dict()
        self.unpickle()
        self.sector_mapp = [
                            [fr'{self.DIR_DICT["FROM_HOST"]}\_FROMOUTLOOK\Unbilled Detail and Aging.csv', 'ENERGY'],
                            [fr'{self.DIR_DICT["FROM_HOST"]}\_FROMOUTLOOK\Unbilled_EDU.csv', 'EDUCATION'],
                            ]

        self.pm_mapp = pd.read_csv(r'C:\Users\V Song\Desktop\_PYTHONSCRIPTS\_mapping\PROJECTNUMBER_MANAGER.csv',
                                   index_col='PROJ',)

        self.df_tosend = list()
        self.filter_list = [
                        '1406.000.001',
                        '1406.000.002',
                        '1406.000.003',
                        '1383.000.001',
                        '1360.000.001',
                        '1319.\d{3}.\d{3}',
                    ]
        
        self.cc_list = [ 
                        'janel.toregozhina@cordobacorp.com;',
                        'vincent.tran@cordobacorp.com',
        ]

    def unpickle(self):
        with open(r"C:\Users\V Song\Desktop\_PYTHONSCRIPTS\_picklenv\data\envs.pickle", "rb") as pfile:
            self.DIR_DICT = pickle.load(pfile)
        return

    def run_unbilled(self):

        drop_col = [
                    'detail_billStatus', 'detail_laborCode',
                    'groupFooter6_Age1Amount', 'groupFooter6_Age2Amount',
                    'groupFooter6_Age3Amount', 'groupFooter6_Age4Amount',
                    'groupFooter6_Age5Amount', 'groupFooter6_Age6Amount',
                    'groupFooter5_GroupColumn', 'groupFooter5_hoursQty',
                    'groupFooter5_billAmount', 'groupFooter5_Age1Amount',
                    'groupFooter5_Age2Amount', 'groupFooter5_Age3Amount',
                    'groupFooter5_Age4Amount', 'groupFooter5_Age5Amount',
                    'groupFooter5_Age6Amount', 'groupFooter4_GroupColumn',
                    'groupFooter4_hoursQty', 'groupFooter4_billAmount',
                    'groupFooter4_Age1Amount', 'groupFooter4_Age2Amount',
                    'groupFooter4_Age3Amount', 'groupFooter4_Age4Amount',
                    'groupFooter4_Age5Amount', 'groupFooter4_Age6Amount',
                    'groupFooter3_GroupColumn', 'groupFooter3_hoursQty',
                    'groupFooter3_billAmount', 'groupFooter3_Age1Amount',
                    'groupFooter3_Age2Amount', 'groupFooter3_Age3Amount',
                    'groupFooter3_Age4Amount', 'groupFooter3_Age5Amount',
                    'groupFooter3_Age6Amount', 'groupFooter2_GroupColumn',
                    'groupFooter2_hoursQty', 'groupFooter2_billAmount',
                    'groupFooter2_Age1Amount', 'groupFooter2_Age2Amount',
                    'groupFooter2_Age3Amount', 'groupFooter2_Age4Amount',
                    'groupFooter2_Age5Amount', 'groupFooter2_Age6Amount',
                    'groupFooter1_GroupColumn', 'groupFooter1_hoursQty',
                    'groupFooter1_billAmount', 'groupFooter1_Age1Amount',
                    'groupFooter1_Age2Amount', 'groupFooter1_Age3Amount',
                    'groupFooter1_Age4Amount', 'groupFooter1_Age5Amount',
                    'groupFooter1_Age6Amount', 'textbox74', 'cLabel', 'textbox55',
                    'Labor_Billable', 'Consultant_Billable', 'Expense_Billable',
                    'Unit_Billable', 'textbox54', 'Labor_Deleted', 'Consultant_Deleted',
                    'Expense_Deleted', 'Unit_Deleted', 'textbox50', 'Labor_Held',
                    'Consultant_Held', 'Expense_Held', 'Unit_Held', 'textbox52',
                    'Labor_WriteOff', 'Consultant_WriteOff', 'Expense_WriteOff',
                    'Unit_WriteOff', 'textbox51', 'Labor_Total', 'Consultant_Total',
                    'Expense_Total', 'Unit_Total',
                    'detail_Age1Amount', 'detail_Age2Amount', 'detail_Age3Amount',
                    'detail_Age4Amount', 'detail_Age5Amount', 'detail_Age6Amount',
                    'groupFooter6_GroupColumn', 'groupFooter6_hoursQty',
                    'groupFooter6_billAmount'
                    ]

        rename_col = {
                    'groupHeader1_GroupColumn': "PROJ",
                    'groupHeader2_GroupColumn': "TASK",
                    'groupHeader3_GroupColumn': "SUBTASK",
                    'groupHeader4_GroupColumn': "LABORTYPE",
                    'detail_transDate': "DATE",
                    'detail_employee': "EMPID",
                    'detail_Description': "EMPNAME",
                    'detail_hoursQty': "HOURS",
                    'detail_billRate': "RATE",
                    'detail_billAmount': "TOTALAMT",
                    }

        for fpath in self.sector_mapp:
            df = pd.read_csv(fpath[0], skiprows=3)

            # DROP/Rename Cols
            df = df.drop(columns=drop_col)
            df = df.rename(columns=rename_col)

            # SLICE DF
            df['PROJDESC'] = df['PROJ'].map(lambda x: x[29:])
            df['PROJ'] = df['PROJ'].map(lambda x: x[16:28])

            # Filter Clean DF
            filter1 = ~df['PROJ'].str.contains('.999')
            filter2 = df['LABORTYPE'].str.contains('   Labor:')
            filter3 = df['TOTALAMT'].isnull()
            filter4 = ~df['PROJ'].str.contains('|'.join(self.filter_list))

            # Apply all filters into raw csv
            clean_df = df[filter1 & filter2 & filter3 & filter4]

            clean_df = clean_df.drop(columns=[
                                        'TASK', 'SUBTASK', 'LABORTYPE', 'DATE',
                                        'HOURS', 'RATE', 'TOTALAMT', 'EMPID',
                                    ])

            clean_df = clean_df[['PROJ', 'PROJDESC', 'EMPNAME']]
            clean_df = clean_df.set_index('PROJ')
            clean_df = clean_df.drop_duplicates()

            print(clean_df.dtypes, clean_df)
            print(self.pm_mapp.dtypes, self.pm_mapp)

            self.df_tosend.append(pd.merge(clean_df, self.pm_mapp, left_index=True, right_index=True))

        return

    def send_unbilled_email(self):

        filter_string = '<ul>'

        for x in self.filter_list:
            filter_string += f'<li>{x}</li>'

        filter_string += '</ul>'

        for index, sector in enumerate(self.sector_mapp):

            olapp = Dispatch('Outlook.Application')
            olmail = olapp.CreateItem(0)

            olmail.CC = ';'.join(self.cc_list)

            olmail.Subject = f'DELTEK ZERO RATE - {sector[1]}'

            olmail.HTMLbody = f'{sector[1]} - <br><br>' \
                              f'The following projects have unbilled zero rates: <br><br>' \
                              f'Please reach out to your PM' \
                              f'<br><br>' \
                              f'{self.df_tosend[index].to_html(border=None)}' \
                              f'<br><br>' \
                              f'Fee-based projects filtered out:' \
                              f'{filter_string}<br>'  \
                              f'Thanks'

            olmail.display(True)

        return


def main():
    ub = UnbilledCheck()
    ub.run_unbilled()
    ub.send_unbilled_email()
    # ub.filter_itd()


if __name__ == '__main__':
    main()
