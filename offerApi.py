import requests
import json


def offerApi():
    url = "http://52.74.32.248:8080/api/vod/offer"

    payload = {}
    headers = {}

    response = requests.request("GET", url, headers=headers, data=payload)
    data = json.loads(response.text)
    return data['offer']


# res = offerApi()
# for item in res['packageType']:
    # if item['packageName'] == 'Highlight':
        # print(item['packageId'])


