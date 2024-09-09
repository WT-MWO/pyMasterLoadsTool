from openpyxl.utils import get_column_letter


def clear_range(ws, range):
    cell_range = ws[range]
    for row in cell_range:
        for cell in row:
            cell.value = None


def get_key(dict, val):
    for key, value in dict.items():
        if val == value:
            return key


def clear_range2(ws, min_row=1, max_row=1, min_col=1, max_col=1):
    "Clear range much faster"
    for row in ws.iter_rows(min_row, max_row, min_col, max_col):
        for cell in row:
            cell.value = None


def max_row_index(ws, start_row=1, max_row=1, min_column=1, max_column=1):
    """Returns last row index containing data."""
    if start_row == 1:
        count = 1
    else:
        count = start_row - 1
    for row in ws.iter_rows(min_row=start_row, max_row=ws.max_row, min_col=min_column, max_col=max_column):
        for cell in row:
            if cell.value is not None:
                count += 1
    return count


def max_column_index(ws, start_col=1, min_row=1, max_row=1):
    """Returns last column index containing data."""
    if start_col == 1:
        count = 1
    else:
        count = start_col - 1
    for col in ws.iter_cols(min_row=min_row, max_col=ws.max_column, max_row=max_row):
        for cell in col:
            if cell.value is not None:
                count += 1
    return count


def write_list_to_file(list, path, filename):
    """Writes each item in list to new line in .txt file.
    Parameters:
        list(list): list with items
        path(string):  path for txt file
        filename(string): name for file
    """
    # Fix this function to write plain text without
    # parantesis always the same if its a set or list
    with open(path + filename, "w") as file:
        for item in list:
            file.write("{}\n".format(str(item)))
