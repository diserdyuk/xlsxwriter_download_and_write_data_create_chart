import gdown, xlsxwriter, csv


def download_csv():
    url = 'https://drive.google.com/uc?export=download&id=1f4Z89sdriZhEdVSVtC2zQz5qS-J3R2Py' 
    # foreign-holdings-in-brazil-chart.csv
    output = 'foreign-holdings-in-brazil-chart.csv' 
    gdown.download(url, output, quiet=False)


def create_file_xlsx():
    workbook = xlsxwriter.Workbook('./Foreign Holdings.xlsx')
    worksheet_data = workbook.add_worksheet('Foreign_Holdings_in_Brazil')
    
    font_size_16 = workbook.add_format(
        {'font_size': 16, 
         'bold': True
         })

    worksheet_data.write('B1', 'Foreign Holdings in Brazil, USD bn', font_size_16)

    font_bold = workbook.add_format({'bold': True})
    
    worksheet_data.write('A3', 'Date', font_bold)
    worksheet_data.write('B3', 'Brazil', font_bold)
    worksheet_data.freeze_panes(3, 2)

    data = read_csv()
    start_row, start_column = 3, 0
    start_row_2, start_column_2 = 3, 1 
    
    while True:
        for i in data: 
            if i[0][0].isdigit() == False:
                continue
            worksheet_data.write(start_row, start_column, i[0].replace('-','.'))
            worksheet_data.write(start_row_2, start_column_2, float(i[1]))

            start_row += 1
            start_row_2 += 1
        break

    worksheet_data.write_formula('B175', "=SUM(B4:B173)", font_bold)

    # new list, for chart
    worksheet_chart = workbook.add_worksheet('Chart_(Bazil)')

    chart = workbook.add_chart({'type': 'column'})

    chart.set_size({'width': 1000, 'height': 360})
    chart.set_title({'name': 'Brazil, USD bn',
                     'name_font': {'size': 15, 'bold': False}
                     })
    chart.set_legend({'position': 'none'})

    chart.add_series({'values': '=Foreign_Holdings_in_Brazil!$A$4:$A$171',
                      'values': '=Foreign_Holdings_in_Brazil!$B$4:$B$171',
                      'column': {'color':'blue'}
                      })

    worksheet_chart.insert_chart('A1', chart)

    worksheet_data.write_url('M1', 'internal:Chart_(Bazil)!A1', string='Go to chart')

    workbook.close()


def read_csv():
    with open('foreign-holdings-in-brazil-chart.csv', 'r') as f:
        reader = csv.reader(f)
        data_read = [row for row in reader]

        return data_read


def main():
    # download_csv()
    create_file_xlsx()
    # read_csv()


if __name__ == "__main__":
    main()

