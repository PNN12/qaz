
def get_values(listbook: str = None, text: str = None):
    list_value, dict_value = [], {}
    max_row = listbook.max_row
    max_column = listbook.max_column

    for i in range(1, max_row + 1):
        for j in range(1, max_column + 1):
            text_value = str(listbook.cell(row=i, column=j).value)
            if text_value.find(text) >= 0:
                for col in range(1, max_column + 1):
                    if col == max_column:
                        dict_value[i] = list_value
                    list_value.append(listbook.cell(row=i, column=col).value)
            list_value = []
    return dict_value


def get_categories(listbook: str = None, text: str = None, category: bool = False):
    list_category, list_value = [], []
    for i in range(1, listbook.max_column + 1):
        list_category.append(listbook.cell(row=1, column=i).value)
    if category:
        return list(i for i in list_category if i is not None)
    else:
        for i, col in enumerate(list_category):
            if col == text:
                for j in listbook.rows:
                    list_value.append(j[i].value)
    return list_value
