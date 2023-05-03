import pandas as pd
from win32com.client import Dispatch


class FloorChecker:
    # this script just distributes emails
    # you have to prime the excel file located in self.main_filepath before working with the script
    # how to prime the excel file: please refer to INSTRUCTION sheet in the file
    def __init__(self):
        self.main_filepath = r'C:\Users\V Song\Box\Invoicing\_FINANCE\_REPORTS\2022\FLOORCHECK'
        self.sheet_4 = 'PER4'
        self.sheet_3 = 'PER3'
        self.sheet_week_range = 'WEEKRANGE'
        
        self.distr_cc_list = [
                        'cc.listfrom@company.com'
                                    ]
        
        self.energy_directors = (
            'director.email@company.com'
        )

        self.table_columns = (
            'PERRANGE', 'Employee', 'Timesheet Status', 'Time Group',
            'Employee Name', 'Expected Hours', 'Hours Entered', 'Supervisor',
            'DELTEKAPPROVERID', 'SECTOR', 'SECTORCLEAN', 'DELTEKAPPROVERNAME',
            'DELTEKAPPROVEREMAIL'
        )

        self.corp_sector = (
            'Energy', 'Education',
            'Water', 'Transportation'
        )

        self.send_list = {
            'Energy': [],
            'Education': [],
            'Water': [],
            'Transportation': [],
        }

        self.send_table = pd.DataFrame()
        self.df_per_four = pd.DataFrame(columns=self.table_columns)
        self.df_per_three = pd.DataFrame(columns=self.table_columns)
        self.week_range = pd.DataFrame()

    def load_work_order(self):

        self.df_per_four = pd.read_excel(fr'{self.main_filepath}\MASTER_FLOOR_CHECK.xlsx',
                                         sheet_name=self.sheet_4,
                                         index_col=False)

        self.week_range = pd.read_excel(fr'{self.main_filepath}\MASTER_FLOOR_CHECK.xlsx',
                                         sheet_name=self.sheet_week_range,
                                         index_col=False, header=None)
        # print(self.df_per_four.info())
        return None

    def send_email(self, sector):

        olapp = Dispatch('Outlook.Application')
        olmail = olapp.CreateItem(0)

        if sector == 'Energy':
            for contact in self.energy_directors:
                if contact not in self.send_list[sector]:
                    self.send_list[sector].append(contact)
            olmail.Bcc = ';'.join(self.send_list[sector])

        else:
            olmail.Bcc = ';'.join(self.send_list[sector])

        olmail.CC = ';'.join(self.distr_cc_list)

        olmail.Subject = f'DELTEK {sector} PENDING APPROVAL PERIOD: {self.week_range[0].to_string(header=False, index=False)} to {self.week_range[1].to_string(header=False, index=False)}'
        olmail.HTMLbody = f'{sector} - <br><br>' \
                          f'The following timesheet approvals are overdue <br><br>' \
                          f'Please approve the following timesheets in DELTEK at your earliest ' \
                          f'convenience to avoid billing delays <br><br>' \
                          f'{self.send_table.to_html(border=None)}' \
                          f'<br><br>' \
                          f'Thanks'

        olmail.display(True)
        # olmail.display

        del olmail
        del olapp

        return None

    def load_to_list(self):
        # creates the list of approve emails by sector
        # the dict is stored as self.send_list
        # where key = sector
        # and value = approve email

        for n, x in enumerate(self.corp_sector):
            result_to_list = self.df_per_four[self.df_per_four[self.df_per_four.columns[9]] == x]
            # print(result_to_list)
            self.send_table = self.df_per_four[self.df_per_four[self.df_per_four.columns[9]] == x]
            result_to_list = result_to_list[result_to_list.columns[11]]
            self.send_list[x] = result_to_list.drop_duplicates().reset_index(drop=True).to_list()
            self.send_table = self.send_table.drop([
                                                    'Supervisor',
                                                    'DELTEKAPPROVERID',
                                                    'SECTOR',
                                                    'SECTORCLEAN',
                                                    'DELTEKAPPROVEREMAIL'
                                                        ], axis=1)
            print(self.send_table)
            self.week_range = self.week_range.loc[lambda df: df[3] == 'a']
            self.week_range = self.week_range.reset_index(drop=True)
            # print(self.week_range[0].to_string(header=False, index=False))
            self.send_email(x)

        return

    def start_wo(self):
        self.load_work_order()
        self.load_to_list()


def main():
    a = FloorChecker()
    a.start_wo()


if __name__ == '__main__':
    main()
