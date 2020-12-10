import xlrd
import pandas as pd

### Year should be in form a matrix, even if a single value, please put it inside a matrix
year = ['2011', '2012']
### filename should be with the path, if it is local file, no need of path
### IMPORTANT [FILE SHOULD BE IN xls format only], this script will not work with xlsx format
# filename_with_path = 'myb3-2015-ch.xls'
filename_with_path = 'C:\\Users\\Arjun Janamatti\\Downloads\\Bolivia_2015.xls'

### SHEET NAME, in some examples it was 'Table1' and in others it was 'Table 1', hence this variable
sheet_name = 'Table 1'
### UNIT name default is mentioend as 'metric_tons'
unit_name = 'metric_tons'
### THIS is the row number from where the original table starts, in china example the values start at 8, hence this value is 7
row_number_start_data = 7
### Below is for filename
country_name = 'Bolivia'
### In some instances the unit column number is 'B', and in some sheets this column was 'D', if the column is 'B', the value should be 1
unit_column_number = 1

def get_dataframe(year, filename_with_path, sheet_name, unit_name, row_number_start_data, country_name, unit_column_number):

    global c_unit_name, b_unit_name
    book = xlrd.open_workbook(filename_with_path, formatting_info=True)

    sheet = book.sheet_by_name(sheet_name)
    num_rows = sheet.nrows
    values_row = (sheet.row_values(row_number_start_data - 2))[unit_column_number + 2:]
    year_row_list = []
    output_file_direc = [_ for _ in filename_with_path.split('\\')]
    output_file_direc = ('/'.join(output_file_direc[:-1]))

    commdity_names  = []

    unit_name_list = []

    for i in range(row_number_start_data,num_rows):
        cell = sheet.cell(i,0)
        unit_cell = sheet.cell(i,unit_column_number)
        fmt = book.xf_list[cell.xf_index]
        if fmt.alignment.indent_level == 0:
            a = cell.value
            if (':' not in cell.value) & (len(a) > 1) & ('white' not in cell.value) & ('includes' not in cell.value) & ('Commodity' not in cell.value) & ('Estimated' not in cell.value) & ('available' not in cell.value) & ('Table' not in cell.value) & ('Reported' not in cell.value) & ('Includes' not in cell.value) & ('addition' not in cell.value) & ('Ditto' not in cell.value) & ('Commodity' not in cell.value) & ('Continued' not in cell.value) & ('footnotes' not in cell.value) & ('unless' not in cell.value):
                if not cell.value.isupper():
                    # print('Intend:0: ', cell.value)
                    commdity_names.append(cell.value)
                    if len(unit_cell.value) > 1:
                        if unit_cell.value != 'do.':
                            print('Intend:0: ', cell.value, unit_cell.value, sheet.row_values(i)[3:])
                            unit_name_list.append(unit_cell.value)
                            a_unit_name = unit_cell.value
                            year_row_list.append(sheet.row_values(i)[3:])
                        elif unit_cell.value == 'do.':
                            print('Intend:0: ', cell.value, '\t Unit name: ', a_unit_name, sheet.row_values(i)[3:])
                            unit_name_list.append(a_unit_name)
                            year_row_list.append(sheet.row_values(i)[3:])
                    else:
                        print('Intend:0: ', cell.value, '\t Unit name: ', unit_name, sheet.row_values(i)[3:])
                        unit_name_list.append(unit_name)
                        year_row_list.append(sheet.row_values(i)[3:])

        if fmt.alignment.indent_level == 1:
            if (':' not in cell.value):
                # print('Intend:1: ', a + cell.value)
                commdity_names.append(a + cell.value)
                if len(unit_cell.value) > 1:
                    if unit_cell.value != 'do.':
                        print('Intend:1: ', a + cell.value, '\t Unit name: ', unit_cell.value, sheet.row_values(i)[3:])
                        unit_name_list.append(unit_cell.value)
                        b_unit_name = unit_cell.value
                        year_row_list.append(sheet.row_values(i)[3:])
                    elif unit_cell.value == 'do.':
                        print('Intend:1: ', a + cell.value, '\t Unit name: ', b_unit_name, sheet.row_values(i)[3:])
                        unit_name_list.append(b_unit_name)
                        year_row_list.append(sheet.row_values(i)[3:])
                else:
                    print('Intend:1: ', a + cell.value, '\t Unit name: ', unit_name, sheet.row_values(i)[3:])
                    unit_name_list.append(unit_name)
                    year_row_list.append(sheet.row_values(i)[3:])
            else:
                b = cell.value
        if fmt.alignment.indent_level == 2:
            if ('which' not in cell.value) & (':' not in cell.value) & ('Total' not in cell.value):
                # print('Intend:2: ',a + b + cell.value)
                commdity_names.append(a+ b +cell.value)
                if len(unit_cell.value) > 1:
                    if unit_cell.value != 'do.':
                        print('Intend:2 cell value: ',a + b + cell.value, '\t Unit name: ', unit_cell.value, sheet.row_values(i)[3:])
                        unit_name_list.append(unit_cell.value)
                        year_row_list.append(sheet.row_values(i)[3:])
                        # globals(c_unit_name)
                        c_unit_name = unit_cell.value
                    elif unit_cell.value == 'do.':
                        print('Intend:2: ',a + b + cell.value, '\t Unit name: ', b_unit_name, sheet.row_values(i)[3:])
                        unit_name_list.append(b_unit_name)
                        year_row_list.append(sheet.row_values(i)[3:])

                else:
                    print('Intend:2: ',a + b + cell.value, '\t Unit name: ', unit_name, sheet.row_values(i)[3:])
                    unit_name_list.append(unit_name)
                    year_row_list.append(sheet.row_values(i)[3:])

            if (':' in cell.value):
                c = cell.value
        if fmt.alignment.indent_level == 3:
            if ('Total' not in cell.value):
                # print('Intend 3: ', a + b + c + cell.value)
                commdity_names.append(a + b + c + cell.value)
                if len(unit_cell.value) > 1:
                    if unit_cell.value != 'do.':
                        print('Intend 3: ', a + b + c + cell.value, '\t Unit name: ', unit_cell.value, sheet.row_values(i)[3:])
                        unit_name_list.append(unit_cell.value)
                        year_row_list.append(sheet.row_values(i)[3:])
                        # globals(c_unit_name)
                        c_unit_name = unit_cell.value
                    elif unit_cell.value == 'do.':
                        # if b_unit_name:
                        #     print('Intend 3: ', a + b + c + cell.value, '\t Unit name: ', b_unit_name, sheet.row_values(i)[3:])
                        #     unit_name_list.append(b_unit_name)
                        #     year_row_list.append(sheet.row_values(i)[3:])
                        # else:
                        try:
                            print('Intend 3: ', a + b + c + cell.value, '\t Unit name: ', b_unit_name,
                                  sheet.row_values(i)[3:])
                            unit_name_list.append(b_unit_name)
                            year_row_list.append(sheet.row_values(i)[3:])

                        except NameError:
                            print("well, it WASN'T defined after all!")
                            print('Intend 3: ', a + b + c + cell.value, '\t Unit name: ', c_unit_name, sheet.row_values(i)[3:])
                            unit_name_list.append(c_unit_name)
                            year_row_list.append(sheet.row_values(i)[3:])
                else:
                    print('Intend 3: ', a + b + c + cell.value, '\t Unit name: ', unit_name, sheet.row_values(i)[3:])
                    unit_name_list.append(unit_name)
                    year_row_list.append(sheet.row_values(i)[3:])

    df = pd.DataFrame()
    df['commdity_names'] = pd.Series(commdity_names)
    df['unit_names'] = pd.Series(unit_name_list)

    index_year_input = []
    year_list_to_df = []
    for index,i in enumerate(values_row):
        if i in year:
            index_year_input.append(index)
    for k in index_year_input:
        temp_list = []
        for j in year_row_list:
            temp_list.append(j[k])
        year_list_to_df.append(temp_list)
    print('Length of year list: ', len(index_year_input))

    print('Length of year list: ', len(year_list_to_df))
    for index, m in enumerate(year):
        df[m] = pd.Series(year_list_to_df[index])

    # print(len(df))
    if len(year) > 1:
        df.to_csv('{}'.format(output_file_direc)+'/'+'{}{}-{}.csv'.format(country_name,
                                      year[0],
                                      year[-1]))
    elif len(year) == 1:
        df.to_csv('{}{}.csv'.format(country_name,
                                    year[0]))


    return df

get_dataframe(year, filename_with_path, sheet_name, unit_name, row_number_start_data, country_name, unit_column_number)
