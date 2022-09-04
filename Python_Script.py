import pandas as pd

# config for section name and limit
SECTIONS = [
    {"name": 'Q1', "limit": 16},
    {"name": 'Q2', "limit": 16},
    {"name": 'Q3', "limit": 6},
    {"name": 'Q4', "limit": 6},
]

order = []  # output order item list, initially empty
item_list_df = None  # initialize item list dataframe


def isNan(term):
    """Returns True if string is NaN otherwise False"""
    return term != term


def check_predecessor_added(pred_items):
    """Returns True if all predecessors are already added otherwise False"""
    for index_pred_item, pred_item in enumerate(pred_items):
        for index, item in item_list_df.iterrows():
            if item["Item Code"] == pred_item and item["Initial Status"] == 'No':
                status = False
                for section_index, section in enumerate(order):
                    if pred_item in section['items']:
                        status = True
                        break
                if status is False:
                    return False
                break
    return True


def get_simultaneous_units(simultaneous, units):
    """Returns all simultaneous items and sum of their units"""
    items = []
    for index_simul_item, simul_item in enumerate(simultaneous):
        for index, item in item_list_df.iterrows():
            if item["Item Code"] == simul_item and item["Initial Status"] == 'No':
                units += item["Units"]
                items.append(item["Item Code"])

    return units, items


def update_simultaneous_units(simultaneous):
    """Updates Initial Status of simultaneous items to Yes"""
    for index_simul_item, simul_item in enumerate(simultaneous):
        for index, item in item_list_df.iterrows():
            if item["Item Code"] == simul_item and item["Initial Status"] == 'No' and item["WIP Status"] == 'No':
                item_list_df.at[index, 'Initial Status'] = 'Yes'


def add_item(item_index, item):
    """Try to add an item to order section"""
    predecessor = [] if isNan(
        item['Predecessor']) else item['Predecessor'].split(",")
    simultaneous = [] if isNan(
        item['Simultaneous']) else item['Simultaneous'].split(",")

    # check predecessor
    if len(predecessor) > 0:
        if not check_predecessor_added(predecessor):
            return "predecessor not added"

    units = item['Units']

    # check simultaneous
    if len(simultaneous) > 0:
        units, items = get_simultaneous_units(simultaneous, units)

    for index, section in enumerate(order):
        if (section['total'] + units <= section['limit']):

            # add item in section item list
            section['items'].append(item["Item Code"])

            # update section total units
            section['total'] += units

            # update Initial Status of item
            item_list_df.at[item_index, 'Initial Status'] = 'Yes'

            # if simultaneous update add to order and update its status
            if len(simultaneous) > 0:
                update_simultaneous_units(simultaneous)
                section['items'] += items
                pass

            return 'added'

    return 'not added'


def calculate_order():
    """iterate over items until Initial Status is No and WIP Status is No """
    for index, item in item_list_df.iterrows():
        # print(Initial_Status, WIP Status, Item_code, units, predecessor, simultaneous)
        status = 'not added'
        while status == 'not added' and item_list_df['Initial Status'][index] == 'No' and item_list_df['WIP Status'][index] == 'No':
            status = add_item(index, item)

            # if status is not added append new section in order dict
            if status == 'not added':
                section = SECTIONS[len(order) % len(SECTIONS)]

                # adding new section in order
                order.append(dict(
                    name=section['name'],
                    limit=section['limit'],
                    total=0,
                    items=[]
                ))



def check_pending_item():
    """check if there is item with Initial Status as 'No' and WIP Status is No """
    for index, item in item_list_df.iterrows():
        if item['Initial Status'] == 'No' and item['WIP Status'] == 'No':
            return True
    return False


def format_output():
    """format output dict in required format"""
    maxSectionItems = 0 if len(order) == 0 else max(
        [len(x['items']) for x in order])
    temp_dict = dict()
    temp_dict[0] = ["Item"]*maxSectionItems
    headers = [""]

    for index, section in enumerate(order):
        headers.append(section["name"])
        temp_dict[index+1] = []

        for item_index, item in enumerate(section["items"]):
            temp_dict[index+1].append(item)

        # padding empty cells
        temp_dict[index+1] += [''] * \
            (maxSectionItems - len(temp_dict[index+1]))

        temp_dict[index+1].append(section["total"])

    return headers, temp_dict


if __name__ == '__main__':
    # read input excel file into pandas dataframe
    item_list_df = pd.read_excel('Test Database.xlsx')

    # while there is item pending for order, do the calculation
    while check_pending_item():
        calculate_order()

    # format the output dict to be in required format
    headers, output = format_output()

    # convert order dict to pandas dataframe
    df = pd.DataFrame.from_dict(output, orient='index').transpose()

    # write output order to excel
    df.to_excel("output.xlsx", index=False, header=headers)

    # write updated item list in excel
    item_list_df.to_excel("updated_item_list.xlsx", index=False)

    print("output saved to output.xlsx")
    print("updated item list saved to updated_item_list.xlsx")
