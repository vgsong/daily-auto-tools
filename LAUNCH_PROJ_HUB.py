import csv
import webbrowser


class ProjLauncher:
    def __init__(self):
        self.list_fpath = fr'C:\Users\V Song\Desktop\_PYTHONSCRIPTS\_temp\TS_PROJECT_ADD_LIST_COMMON.csv'
        self.proj_csv = self.get_proj_list()
        self.url = 'https://webappurl.com'

    def get_proj_list(self):
        result = list()
        with open(self.list_fpath, 'r', newline='') as pf:
            data = csv.reader(pf, delimiter=',')

            for x in data:
                result.append(x[0])

        print('Proj list loaded!')

        return result

    def start_launcher(self):

        for i, proj in enumerate(self.proj_csv):
            print(f'{i} - {proj}')

        while True:
            user_input = input('Please enter index to launch in Proj Hub:\n')

            if user_input == 'q':
                exit()

            elif len(user_input) == 12:
                print(f'launching {user_input}')
                webbrowser.open(self.url.format(str(user_input)))
                break

            else:
                try:
                    webbrowser.open(self.url.format(self.proj_csv[int(user_input)]))
                    return

                except IndexError as e:
                    print(f'ERROR: {e}')
                    self.start_launcher()

                except ValueError as e:
                    print(f'ERROR: {e}\n')
                    self.start_launcher()


def main():
    pj = ProjLauncher()
    pj.start_launcher()


if __name__ == '__main__':
    main()
