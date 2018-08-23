import xlrd

SHEET1 = 'PO Uploaded'
SHEET2 = 'User Active'
SHEET3 = 'Production List User'

def x(SheetName) :

    workbook = xlrd.open_workbook('xx.xlsx')
    sheet = workbook.sheet_by_name(SheetName)

    curr_row = 0
    row_list = []
    new_row = []

    while curr_row < (sheet.nrows - 1):
        curr_row += 1
        row = sheet.row_values(curr_row)
        row_list.append(row)
        # new_row.append([[ele.value for ele in each] for each in row_list])
    return row_list


def xz(SheetName) :

    workbook = xlrd.open_workbook('xx.xlsx')
    sheet = workbook.sheet_by_name(SheetName)

    curr_row = 0
    row_list = []
    new_row = []

    row = list(filter(None, sheet.row_values(1)))

    for i in row :
        new_row.append(i)

    # for z in row :
    #     filt = filter(None, z)
    #     for i in filt :
    #         new_row.append(filt)
        # new_row.append([[ele.value for ele in each] for each in row_list])
    return row

def ss() :

    workbook = xlrd.open_workbook('xx.xlsx')
    sheet = workbook.sheet_by_name(SHEET2)

    sheet.conditional_format(2, 8, 100, 100, {
                'type': 'cell',
                'criteria': 'equal to',
                'value': '"Yes"',
                'format': orange})

    orange = workbook.add_format({'bg_color': '#F5D76E'})



    # for z in row :
    #     filt = filter(None, z)
    #     for i in filt :
    #         new_row.append(filt)
        # new_row.append([[ele.value for ele in each] for each in row_list])
    return row

ss()

# print(xz(SHEET2))

