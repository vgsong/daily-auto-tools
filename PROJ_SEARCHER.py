import os
import pandas as pd

pd.set_option('display.max_rows', None)


class PROJSearcher:
    def __init__(self):
        self.df_fpath = r'C:\Users\V Song\Documents\FromHost\_FROMOUTLOOK\PROJECT_LIST_EXPORT.csv'

    def remove_df_duplicates(self):
        final_df = pd.DataFrame(columns=['PROG', 'PROJ', 'DESC'])
        df = pd.read_csv(self.df_fpath)
        df = df[df['detail_Phase'].str.contains(' ')]

        # print(df)

        df.to_csv(r'output.csv')

        final_df['PROJ'] = df['detail_Project']
        final_df['PROG'] = df['PROJ'].map(lambda x: x[:5])
        print(final_df)


def main():
    ps = PROJSearcher()
    ps.remove_df_duplicates()


if __name__ == '__main__':
    main()
