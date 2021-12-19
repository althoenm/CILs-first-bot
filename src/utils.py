import re

# Check time format, throw error if not in HH:MM am/pm
def check_time_format(time_string):
    time_re = re.compile(r'\b((1[0-2]|[1-9]):([0-5][0-9]) ([ap][m]))')
    if time_re.match(time_string):
        return True
    else:
        raise Exception("-E- An entered time is not in HH:MM am/pm format")

# Read excel rows into a list of lists
def read_excel_rows_to_list(sheet_name, list_name):
    for rows in sheet_name.iter_rows():
        row_cells = []
        for cell in rows:
            if cell.value is not None:
                row_cells.append(cell.value)
            else:
                break
        list_name.append(list(row_cells))
    return list_name

# Remove empty rows from a list of lists
def remove_empty_rows(list_name):
    list_name = [x for x in list_name if x != []]
    return list_name

