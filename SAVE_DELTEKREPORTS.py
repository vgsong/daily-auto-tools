import pickle

from win32com.client import Dispatch
from datetime import datetime, timedelta


class DeltekReportSaver:
    def __init__(self):
        self.outlook_maininbox = 'my.email@company.com'

        self.DIR_DICT = dict()
        self.unpickle()

        self.dtekmailinfo = {
            # 'DTEK_DAILY_Unbilled Detail and Aging Report': ['UNBILLED_DETAILS.txt', 'UNBILLED_DETAILS.csv'],
            'DTEK_DAILY_Unbilled Detail and Aging Report': ['Unbilled Detail and Aging.txt',
                                                            'Unbilled Detail and Aging.csv',
                                                            True,
                                                            ],

            'Project List Export Report': ['PROJECT_LIST_EXPORT.txt',
                                           'PROJECT_LIST_EXPORT.csv',
                                           True,
                                           ],

            'DTEK_DAILY_EMP_LIST Report': ['DTEK_DAILY_EMP_LIST.txt',
                                           'DTEK_DAILY_EMP_LIST.csv',
                                           True,
                                           ],

            'DTEK_DAILY_Unposted Labor Report': ['UNPOSTED_LABOR_DETAIL.txt',
                                                 'UNPOSTED_LABOR_DETAIL.csv',
                                                 False,
                                                 ],

            'DAILY_DTEK_AR Aged Report': ['AR Aged.txt',
                                          'AR_Aged.csv',
                                          True
                                          ],

            '1112_Project List Export Report': ['1112_Project List Export.xlsx',
                                                '1112_Project List Export.xlsx',
                                                False
                                                ],

            '1168_Project List Export Report': ['1168_Project List Export.xlsx',
                                                '1168_Project List Export.xlsx',
                                                False
                                                ],

            'Unbilled Detail and Aging Report EDU': ['Unbilled_EDU.txt',
                                                     'Unbilled_EDU.csv',
                                                     True
                                                     ],
        }


        self.qs_mail= self.outlook_load_mailItems()
        self.outlook_get_dailyreports()

    def unpickle(self):
        with open(r"C:\Users\V Song\Desktop\_PYTHONSCRIPTS\_picklenv\data\envs.pickle", "rb") as pfile:
            self.DIR_DICT = pickle.load(pfile)
        return

    def outlook_load_mailItems(self):

        def outlook_get_inbox():
            olapp = Dispatch('Outlook.Application')
            result = olapp.GetNameSpace('MAPI').Folders(self.outlook_maininbox).Folders('Inbox').Items
            return result

        def outlook_filter_inbox(aitems, hdiff):
            # Sets the filter for inbox by Received Time
            # uses datetime.now() minus hours (keep the scope to minimum)
            result_list = []
            received_date = datetime.now() - timedelta(hours=hdiff)
            received_date = received_date.strftime('%m/%d/%Y %H:%M %p')
            result = aitems.Restrict("[ReceivedTime] >= '" + received_date + "'")

            for item in result:
                result_list.append(item)

            return result_list

        time_diff = int(24)
        mail_items = outlook_get_inbox()
        filtered_mail = outlook_filter_inbox(mail_items, time_diff)

        return filtered_mail

    def outlook_get_dailyreports(self):

        def dtek_reports_saver(aitem):

            # Procedure that saves attachments for Deltek scheduled distribution
            # dict contains email {SUBJECT : Attachment}

            ATTACHFPATH = fr'{self.DIR_DICT["FROM_HOST"]}\_FROMOUTLOOK'  # local drive
            ATTACHFPATH2 = fr'{self.DIR_DICT["BOX_MAIN"]}\_DTEKREPORTS'  # box network drive remove comment once QA is done

            if aitem.Subject in self.dtekmailinfo:
                # Error handling if I am not logged in the Finance Team Network Drive
                # it tries to save ATTACHFPATH and ATTACHFPATH2, if not logged it saves only on local
                # print(f'{DTEKMAILINFO[aitem.Subject][0]}')
                aitem.Attachments.Item(1).SaveAsFile(rf'{ATTACHFPATH}\{self.dtekmailinfo[aitem.Subject][0]}')
                aitem.Attachments.Item(1).SaveAsFile(rf'{ATTACHFPATH}\{self.dtekmailinfo[aitem.Subject][1]}')
                aitem.Attachments.Item(1).SaveAsFile(rf'{ATTACHFPATH2}\{self.dtekmailinfo[aitem.Subject][0]}')
                aitem.Attachments.Item(1).SaveAsFile(rf'{ATTACHFPATH2}\{self.dtekmailinfo[aitem.Subject][1]}')

                print(f'{self.dtekmailinfo[aitem.Subject]} saved successfully')

        mail_items = self.qs_mail

        for mail in mail_items:
            if mail.Unread:
                dtek_reports_saver(mail)
                mail.Unread = False


def main():
    drs = DeltekReportSaver()
    drs.outlook_load_mailItems()


if __name__== '__main__':
    main()
