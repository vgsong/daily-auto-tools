from win32com.client import Dispatch


class POSender:
    def __init__(self):

        self.att_fpath = fr'C:\Users\V Song\Desktop\_PYTHONSCRIPTS\_temp'

        self.cc_list = [
                    'contracts@cordobacorp.com',
                    'nissagholian@cordobacorp.com',
                    'janel.Toregozhina@cordobacorp.com',
                    'tromo@cordobacorp.com',
                    'salmanza@cordobacorp.com',
                    'michael.avina@cordobacorp.com',
                    'stenorio@cordobacorp.com',
                    'lmurguia@cordobacorp.com',
                    'Elizabeth.Lopez@cordobacorp.com',
        ]
        
        self.test_list = [
            'lmurguia@cordobacorp.com',
            'victor.song@cordobacorp.com',
        ]

        self.wo_dict = {
            'Electric': [
                        'Electric',
                        'Brad Carter',
                        'brad.carter@cordobacorp.com',
                        ],

            'Gas': [
                        'Gas',
                        'Peter Yu',
                        'pyu@cordobacorp.com',
                        ],

            'Energy': [
                        'Energy',
                        'Sam Tenorio',
                        'stenorio@cordobacorp.com',
                    ],
        }

    def send_pos(self, mmm, yyyy, Test_mail=False):

        mper = str(mmm)
        yper = str(yyyy)

        olapp = Dispatch('Outlook.Application')
        
        html_bodytext = f'Hi -- <br><br>' \
                        f'Attached please find Contract Management report. Spent to date is through {mper} {yper} billing cycle. <br>' \
                        f'In addition, contract amendments/CO’s are updated in Deltek once FULLY executed.<br>' \
                        f'Note: Critical contracts meeting this criteria are highlighted in the Senior Management agenda - FINANCE.' \
                        f'<br><br>' \
                        f'Thank you. <br><br> -v'

        # k is sector
        # v[0] is sector (redundant)
        # v[1] is director name = unused just for label/readability
        # v[2] is director email address

        if Test_mail:
            
            for k, v in self.wo_dict.items():
                olmail = olapp.CreateItem(0)
                olmail.To = ';'.join(self.test_list)

                olmail.Subject = f'TESTMAIL#3: {k} Contract Management Report - {mper} {yper}'
                olmail.Attachments.Add(fr'{self.att_fpath}\{v[0]}_MAIN_PO_REPORT_DISTR.xlsx')
                olmail.HTMLbody = html_bodytext

                olmail.display(True)
    
        else:

            for k, v in self.wo_dict.items():
                olmail = olapp.CreateItem(0)
                olmail.To = v[2]

                if k == 'Energy' or k == 'Gas':
                    olmail.Cc = f'{";".join(self.cc_list)}; Hugo.Mejia@cordobacorp.com'
                
                elif k == 'Electric':
                    olmail.Cc = f'{";".join(self.cc_list)}; Jennifer.Yamaguchi@cordobacorp.com; Andrew.Domask@cordobacorp.com; Emily.Newcomer@cordobacorp.com'

                olmail.Subject = f'{k}: Contract Management Report - {mper} {yper}'
                olmail.Attachments.Add(fr'{self.att_fpath}\{v[0]}_MAIN_PO_REPORT_DISTR.xlsx')
                olmail.HTMLbody = html_bodytext
                olmail.display(True)

            return


def main():
    pos = POSender()
    pos.send_pos('MAR', 2023,)

if __name__ == '__main__':
    main()
