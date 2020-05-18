import requests
import json
import xlrd

loc = input("Enter Excel file path: ")
col = int(input("Enter Excel column number from which to read links: "))
apiKey = input("Enter CrowdTangle API key: ").strip()
wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(0)
#sheet.cell_value(0, 0)
for i in range(sheet.nrows):
    link = sheet.cell_value(i, col)
    parameters = {'link': link, "sortBy": "total_interactions"}
    headers = {"x-api-token": apiKey}
    apiURL = 'https://api.crowdtangle.com/links'

    print(f'Getting CrowdTangle data for: {link}')

    response = requests.get(apiURL, params=parameters, headers=headers)
    data = response.json()
    # print(data)
    split = link.split('/')[4]
#     print(split[4])
    filename = split + '.json'
    writeFile = open(filename, 'w')
    json.dump(data, writeFile, indent=4)
    writeFile.close()
    print(f'Writing JSON to {filename}')

print('done.')
