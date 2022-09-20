import pdfplumber, os
import openpyxl, re


def parsing_header(header, file):
    head = [ ''.join(h.split('\n')[::-1]) if h is not None else h for h in header ]
    if not [i for i in head[-5:] if i is not None]:
        last5 = [ 
                    'Total Of Valid Votes',
                    'No. Of Rejected Votes',
                    'NOTA',
                    'Total',
                    'No. Of Tendered Votes',
                ]

        head[-5:] = last5
    
    elif not [i for i in head[-4:] if i is not None]:
        last4 = [ 
                    'Total Of Valid Votes',
                    'No. Of Rejected Votes',
                    'NOTA',
                    'Total',
                ]

        head[-4:] = last4

    head[0] = 'SI NO.'
    head[1] = 'Polling Station No.'

    return head

chk_header = []
def filter_datas(data_list, file):
    line, indx = [], 0
    for ind, data in enumerate(data_list):
        if data[0] is not None:
            if data[0].isdigit():
                indx = ind
                break

    if indx == 0:
        final_out = []
        for row in data_list[indx:]:
            if row[0] is not None:
                h = str(row[0]).lower().replace(' ', '').replace('\n', '')
                remove_head = re.findall('s(i|l)\s?no.', h)
                if not remove_head:
                    final_out.append(row)
        return final_out

    if indx:
        if file not in chk_header:
            chk_header.append(file)
            head = parsing_header(data_list[indx-1], file)
            data_list[indx-1] = head

            final_out = [ [file] + i for i in  data_list[indx-1:]]
            return final_out
        else:

            final_out = [ [file] + i for i in  data_list[indx:]]
            return final_out

def create_workbook(filename):
    wb = openpyxl.Workbook()
    f = f'{filename.split(".")[0]}.xlsx'
    print(f)
    wb.save(f)

def voterList(path):

    for ind, file in enumerate(os.listdir(path), 1):
        if file.endswith('.pdf'):
            pdf = pdfplumber.open(f"{path}\\{file}")
            os.makedirs('voters excel', exist_ok=True)
            create_workbook(f'voters excel\\{file}')

            wb = openpyxl.load_workbook(f'voters excel\\{file.split(".")[0]}.xlsx')
            ws = wb.worksheets[0]

            for page in pdf.pages:
                try:
                    data_list = page.extract_tables()[0]
                    actual_data = filter_datas(data_list, file)
                    
                except Exception as e:
                    print(f'Error: {file}', e)
                    with open('corrupted.txt', 'a') as corrupt_file:
                        corrupt_file.write(f'\n{file}')
                        corrupt_file.close()
                    break
                
                for data in actual_data:
                    ws.append(data)
                    
            wb.save(f'voters excel\\{file.split(".")[0]}.xlsx')

            # if ind == 25:
            #     break

if __name__ == '__main__':
    path = 'constituency wise'
    voterList(path)