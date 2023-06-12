import openpyxl
import json
import string

items = {}
items_found = {}
sorted_item_found = {}
column = "A"
letters = string.ascii_uppercase


def get_parameters():
    while True:
        try:
            amount_sku = int(input("Podaj ilość SKU: "))
            break
        except:
            continue
    for _ in range(amount_sku):
        items[input("Podaj SKU: ")] = 0


def search_boxes():
    wb = openpyxl.load_workbook('test.xlsx')
    name_worksheet = input("Wprowadź nazwę arkusza który przeliczyć: ")
    ws = wb[name_worksheet]
    for i in range(1, ws.max_row + 1):
        if ws[f"{column}{i}"].value in items:
            if items_found.get(ws[f'A{i}'].value):
                items_found[ws[f"{column}{i}"].value] += [{ws[f"D{i}"].value: int(ws[f"C{i}"].value)}]
            else:
                items_found[ws[f"{column}{i}"].value] = [{ws[f"D{i}"].value: int(ws[f"C{i}"].value)}]

    return {k: sorted(v, key=lambda x: list(x.values())[0], reverse=True) for k, v in
            items_found.items()}


def appropriate_amount(sorted_items_found):
    item_amount_need = {}
    output_json = dict()
    while True:
        try:
            for item in items:
                item_amount = int(input(f"Podaj ilość do wyjęcia dla {item}: "))
                item_amount_need[item] = item_amount
            break
        except:
            continue
    for item in items:
        output_json[item] = []
        for item_found in sorted_items_found[item]:
            for box, amount in item_found.items():
                if item_amount_need[item] <= amount:
                    output_json[item] += [box, amount]
                    break
                else:
                    output_json[item] += [box, amount]
                    item_amount_need[item] -= amount
                    continue
            else:
                continue
            break
    with open("test.json", "w") as f:
        json.dump(output_json, f, indent=4)

get_parameters()
output_json = appropriate_amount(search_boxes())

