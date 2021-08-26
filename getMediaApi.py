import requests
import json


def getMediaApi():
    url = "http://52.74.32.248:8080/api/vod/getMediaApi"

    payload = {}
    headers = {}

    response = requests.request("GET", url, headers=headers, data=payload)
    data = json.loads(response.text)
    return data['media']


#res = getMediaApi()
#print(res)
# for item in res['packageType']:
    # if item['packageName'] == 'Highlight':
        # print(item['packageId'])


