from collections import defaultdict
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# params
year = '2017'
month = '10'
verbose = True

# connect to spreadsheet
if verbose:
    print('connect to spreadsheet')
scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_name('d:/python/AutoExpance/cred.json', scope)
client = gspread.authorize(creds)
bus_sheet = client.open('הוצאות').worksheet('עסקים')
exp_sheet = client.open('הוצאות').worksheet('הוצאות')
year_sheet = client.open('הוצאות').worksheet(year)

# read expenses categories line numbers
if verbose:
    print('read expenses categories line numbers')
exp_col = year_sheet.col_values(1)
start = False
exp_cat_col_ind = {}
j = 0
for i in range(0,len(exp_col)):
    if exp_col[i] == 'הוצאות':
        cat_line_start_ind = i+2
        start = True
    else:
        if start:
            exp_cat_col_ind[exp_col[i]] = j
            j = j + 1
            if exp_col[i] == 'אחר':
                cat_line_end_ind = i+1
                break


# read month column numbers
if verbose:
    print('read month column numbers')
months_row = year_sheet.row_values(1)
for i in range(0,len(exp_col)):
    if months_row[i] == month:
        month_col_ind = i+1
        break

# read businesses table
if verbose:
    print('read businesses table')
bus_table = defaultdict(list)
cat_list = bus_sheet.row_values(1)
for i in range(0,len(cat_list)):
    cat_col = bus_sheet.col_values(i+1)
    for bus in cat_col[1:]:
        bus_table[bus] = cat_list[i]

# read this month expenses list
if verbose:
    print('read this month expenses list')
exp_bus = exp_sheet.col_values(1)
exp_debit = exp_sheet.col_values(2)

# read current month debit
if verbose:
    print('read current month debit')
cur_month_col = year_sheet.col_values(month_col_ind)
new_debit_list = cur_month_col[cat_line_start_ind-1:cat_line_end_ind]
# convert to number
for i in range(0,len(new_debit_list)):
    if new_debit_list[i] == '':
        new_debit_list[i] = 0
    else:
        new_debit_list[i] = float(new_debit_list[i])

# main loop
#######################################
if verbose:
    print('main loop')
unknown_cat_bus_list = []
for i in range(0,len(exp_bus)):
    cur_bus = exp_bus[i]
    if cur_bus in bus_table:
        cur_cat = bus_table[cur_bus]
        cur_cat_ind = exp_cat_col_ind[cur_cat]
        new_debit_list[cur_cat_ind] = new_debit_list[cur_cat_ind] + float(exp_debit[i])
        if verbose:
            print('add ' + cur_bus + ' to cat: ' + cur_cat + ', debit: ' + exp_debit[i] + ', total debit: ' + str(new_debit_list[cur_cat_ind]))
    else:
       unknown_cat_bus_list.append(cur_bus)

if unknown_cat_bus_list:
    # write unknown category businesses to businesses table
    print('unknown categories found, fix them and run again')
    for i in range(0, len(unknown_cat_bus_list)):
        bus_sheet.update_cell(i+1, len(cat_list)+1, unknown_cat_bus_list[i])
else:
    # update debit cells
    if verbose:
        print('update debit cells')
    for i in range(0,len(new_debit_list)):
        year_sheet.update_cell(i+cat_line_start_ind, month_col_ind, str(new_debit_list[i]))

print('The End.')