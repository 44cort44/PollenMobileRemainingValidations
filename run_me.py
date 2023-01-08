import requests
import json
import xlsxwriter
import os
import urllib.parse
from requests_auth import HeaderApiKey

def get_daily_reward_status(device, date):
    r = requests.get('https://api.pollenmobile.io/explorer/daily-reward-status?date='+date+'&device='+device, headers=headers, auth=auth)
    return json.dumps(r.json())

parent_dir = os.getcwd()
headers = {'Accept': 'application/json'}
api_key = ""
auth = HeaderApiKey(api_key,'X-API-KEY')
date="2023-01-06"

bumblebees=[
    "WideSupremeBumblebee",
    "MinorElegantBumblebee",
    ]

flowers=[
    "SordidEvasiveDandelion",
    "CrowdedGabbyMoonflower",
    "MindlessVigorousButtercup",
    "WoozyWateryButtercup",
    "SassyVacuousButtercup",
    "TangyBlushingMoonflower"
    ]

output_file = os.path.join(parent_dir, "weekly_remaining.xlsx")
workbook = xlsxwriter.Workbook(output_file)
worksheet = workbook.add_worksheet()

# Light red fill with dark red text.
red_format = workbook.add_format({'bg_color':   '#FFC7CE',
                               'font_color': '#9C0006'})

row = 0

header = ('device', 'flower', 'hexes', 'hexes_remaining')
worksheet.write_row(row,0,header)
row = row + 1

bee_dictionary = {}
for bee in bumblebees:
    bee_dictionary[bee] = {}
    for flower in flowers:
        bee_dictionary[bee][flower] = {}
    returned_from_pollen_api = get_daily_reward_status(bee, date)
    out = json.loads(returned_from_pollen_api)
    for item in out["items"]:
        for validation_reward in item["validation_rewards"]:
            flower = validation_reward["client"]
            for h3_hex in validation_reward["h3_hex"]:
                bee_dictionary[bee][flower][h3_hex] = None
    for flower in bee_dictionary[bee]:
        validated_hex_count = len(bee_dictionary[bee][flower].items())
        max_hexes = 0
        if ("Dandelion" in flower):
            max_hexes = 5
        elif ("Camelia" in flower):
            max_hexes = 5
        elif ("Mosobonzai" in flower):
            max_hexes = 6
        elif ("Elderflower" in flower):
            max_hexes = 10
        elif ("Mosoflower" in flower):
            max_hexes = 12
        elif ("Sunflower" in flower):
            max_hexes = 16
        elif ("Buttercup" in flower):
            max_hexes = 25
        elif ("Moonflower" in flower):
            max_hexes = 30
        line = (bee,flower,validated_hex_count,max_hexes - validated_hex_count)
        worksheet.write_row(row,0,line)
        worksheet.conditional_format(row,3,row,3,
            {'type':    'cell',
            'criteria': 'greater than',
            'value':    0,
            'format':   red_format})
        row = row + 1
worksheet.autofit()
workbook.close()