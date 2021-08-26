import requests
import json


def getPackageTypeApi():
    url = "http://52.74.32.248:8080/api/vod/getPackageTypeApi"

    payload = {}
    headers = {}

    response = requests.request("GET", url, headers=headers, data=payload)
    data = json.loads(response.text)
    return data['packageData']


# res = getPackageTypeApi()
# for item in res['packageType']:
    # if item['packageName'] == 'Highlight':
        # print(item['packageId'])


