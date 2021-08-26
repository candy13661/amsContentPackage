import requests
import json
import datetime


def saveContentapi(contentid, systemtitle, episodic, recordtype, yearofrelease, originallanguage, contentsynopsis, contenttitle, contentgenre, contenttalent):
    # contentid = 12919
    # contenttitle
    # episodic
    # recordtype
    # yearofrelease
    # originallanguage
    createddate = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    url = "http://52.74.32.248:8080/api/v3/epg/saveContent"

    payload = json.dumps({
        "contentCoreData": [
            {
                "contentId": contentid,
                "systemTitle": systemtitle,
                "episodic": episodic,
                "recordType": recordtype,
                "yearOfRelease": yearofrelease,
                "endYear": "",
                "countryOfOrigin": "",
                "originalLanguage": originallanguage,
                "certification": "",
                "visibleStartDate": "",
                "visibleEndDate": "",
                "createdBy": "274",
                "updatedBy": "274",
                "createdDate": createddate,
                "contentTitle": contenttitle,
                "contentSynopsis": contentsynopsis,
                "contentGenre": contentgenre,
                "contentTalent": contenttalent,
                "contentKeyword": [],
                "contentImage": [],
                "contentVideo": [],
                "contentAstroRef": [],
                "contentExtRef": [],
                "contentSeason": [],
                "contentEpisode": [],
                "contentCurated": [],
                "contentExtRefDoc": [],
                "enrichMetadataEpgExport": True,
                "enrichMetadataPriority": False,
                "seriesLinkExport": False,
                "vodEnrich": False,
                "enrichMetadataRequiredLanguage": []
            }
        ]
    })
    headers = {
        'Content-Type': 'application/json'
    }
    response = requests.request("POST", url, headers=headers, data=payload)
    data = json.loads(response.content)
    return data


# res = api(12919, 'Test', 0, 'Content', '2021', 'English')
# print(res['status'], res['responseMessage'])
