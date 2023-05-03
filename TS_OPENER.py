import pandas as pd
from datetime import date, timedelta
import webbrowser


class EmpSearcher:
    def __init__(self, omit_info=True):
        # main path, from report_saver
        self.main_fpath = fr'C:\Users\V Song\Documents\FromHost\_FROMOUTLOOK\DTEK_DAILY_EMP_LIST.csv'

        # deltek url for emphub
        self.url = 'https://cordobacorp.deltekfirst.com/cordobacorp/app/#!Employees/view/0/experience_web/{}/hybrid/'

        # omit info bol
        self.omit_info = omit_info

        # remove sensitive cols also not needed info
        self.columns_to_omit = [
            'detail_JobCostRate',
            'detail_JCOvtPct',
            'detail_TargetRatio',
            'detail_UtilizationRatio',
            'detail_RaiseDate',
        ]

    def get_emp_info(self):

        def get_filtered_emp():
            while True:
                # detail_FullName = col name from original csv data,
                # contains First Last name
                # print(df['detail_FullName'])
                emp_input = input(f'Please enter first or last name to search:\n')

                # search in df using emp_input from end user
                result = df.loc[df['detail_FullName'].str.contains(emp_input, case=False)]

                print(result['detail_FullName'])

                if emp_input == 'q':
                    exit()
                elif len(result) > 0:
                    return result
                else:
                    print('No Results were found. Please try again')
                    continue

        def omit_filtered():
            if self.omit_info:
                return filtered_df.drop(self.columns_to_omit, axis=1)
            return

        def print_filtered_df():
            for i in filtered_df.index:
                print(filtered_df.loc[i])  # iloc
                print(''.ljust(20, '-'))  # spacer

        def print_filtered_df_choices():
            print('Please select result:\n')
            count = 0
            for i in filtered_df.index:
                print(count, '-', filtered_df.loc[i][0])  # iloc
                count += 1
            return

        # GENERAL simple method: it received df from emp_data csv
        # filter based on str obtained from user input()
        # shows output based on close search
        # gives input choices to load in deltel

        # loads df
        df = pd.read_csv(self.main_fpath, skiprows=3, index_col=0)

        filtered_df = get_filtered_emp()  # filters main df from user input and return filtered_df
        filtered_df = omit_filtered()  # drop not needed cols self.omit_info

        # print user found results in loc format
        print_filtered_df()

        # print choices for launching
        print_filtered_df_choices()

        # user input launches deltek emphub
        # user input index return userID
        while True:
            user_input = input('press q to QUIT\nor r to start another search.')

            if user_input == 'q':
                exit()

            elif user_input == 'r':
                self.get_emp_info()

            else:
                try:
                    print(filtered_df.index[0])
                    return filtered_df.index[0]

                except IndexError:
                    print('invalid index. please try again ')

        return


class TsOpener(EmpSearcher):
    def __init__(self):
        super().__init__(self)
        self.input_empid = self.get_emp_info()
        self.yper = 2023
        self.saturday_list = pd.date_range(start=str(self.yper), end=str(self.yper + 1),
                                           freq='W-SAT').strftime('%Y-%m-%d').tolist()

    def open_timesheet(self):

        while True:
            mper_input = input(f'Please enter month (1-12) for year: {self.yper}\n')
            url_month_string = f'-{str(mper_input.zfill(2))}-'
            week_choice_list = list()

            for item in self.saturday_list:
                if url_month_string in item:
                    week_choice_list.append(item)


def main():
    to = TsOpener()
    to.open_timesheet()


if __name__ == '__main__':
    main()
