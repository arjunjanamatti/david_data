import xlrd
import numbers
import pandas as pd


book = xlrd.open_workbook('myb3-2015-ch.xls', formatting_info=True)
first_sheet = book.sheet_by_name('Table1')

sheet = book.sheet_by_name("Table1")
num_rows = sheet.nrows
commodity_row = (sheet.row_values(5))
values_row = (sheet.row_values(5))[3:]
# index_years = []
# years_value = []
# print([(index_years.append(index), years_value.append(value)) for index, value in enumerate(commodity_row)
#        if isinstance(value, numbers.Number) ])

year_row_list = []

raw_dict = {}

commdity_names  = []

unit_name = 'metric_tons'

unit_name_list = []


for i in range(6,num_rows):
    cell = sheet.cell(i,0)
    unit_cell = sheet.cell(i,1)
    fmt = book.xf_list[cell.xf_index]
    if fmt.alignment.indent_level == 0:
        a = cell.value
        if (':' not in cell.value) & (len(a) > 1) & ('Estimated' not in cell.value) & ('available' not in cell.value) & ('Table' not in cell.value) & ('Reported' not in cell.value) & ('Includes' not in cell.value) & ('addition' not in cell.value) & ('Ditto' not in cell.value) & ('Commodity' not in cell.value) & ('Continued' not in cell.value) & ('footnotes' not in cell.value) & ('unless' not in cell.value):
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
                    print('Intend:2: ',a + b + cell.value, '\t Unit name: ', unit_cell.value, sheet.row_values(i)[3:])
                    unit_name_list.append(unit_cell.value)
                    year_row_list.append(sheet.row_values(i)[3:])
                    # globals(c_unit_name)
                    # c_unit_name = unit_cell.value
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
                    # c_unit_name = unit_cell.value
                elif unit_cell.value == 'do.':
                    print('Intend 3: ', a + b + c + cell.value, '\t Unit name: ', b_unit_name, sheet.row_values(i)[3:])
                    unit_name_list.append(b_unit_name)
                    year_row_list.append(sheet.row_values(i)[3:])

            else:
                print('Intend 3: ', a + b + c + cell.value, '\t Unit name: ', unit_name, sheet.row_values(i)[3:])
                unit_name_list.append(unit_name)
                year_row_list.append(sheet.row_values(i)[3:])

df = pd.DataFrame()
df['commdity_names'] = pd.Series(commdity_names)
df['unit_names'] = pd.Series(unit_name_list)
df.to_csv('df.csv')

year = [2012, 2013]
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

print(len(year_list_to_df))
for index, m in enumerate(year):
    df[m] = pd.Series(year_list_to_df[index])

print(len(df))
