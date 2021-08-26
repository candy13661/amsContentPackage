import requests
import json
import datetime


def savePackageapi(packagename, pkgid, pkgtypeid, autopublish, channelowner, pkgclassification, boxsettypeid, sortorder, acqstart, acqend, sponsored, supplierid, contentid, pkgoffers):
    # packagename = 'Test ABC'

    url = "http://52.74.32.248:8080/api/vod/savePackageData"

    payload = json.dumps({
        "vod_data": {
            "packageData": {
                "userId": "274",
                "packageName": packagename,
                "liveEventId": "",
                "packageId": "",
                "pkId": pkgid,
                "inactive": False,
                "export_status": "1",
                "packageTypeId": pkgtypeid,
                "classification": pkgclassification,
                "boxsetTypeId": boxsettypeid,
                "level": False,
                "audienceId": "1",
                "clipId": "HLFCL",
                "cdnId": "1",
                "brandpage": "",
                "brandPackage": [],
                "excludeBrand": False,
                "isSponsoredBrandPage": sponsored,
                "sortOrder": sortorder,
                "supplierId": supplierid,
                "contentId": contentid,
                "acquisitionStart": acqstart,
                "acquisitionEnd": acqend,
                "autoPublish": autopublish,
                "ownerChannel": channelowner,
                "legacyndsId": "",
                "autoUpdateMetadata": False
            },
            "offers": pkgoffers,
            "media": {
                "imagesRibbon": [],
                "mediafull": {
                    "fullAssetClipId": "HLFCL",
                    "assetTypeId": "2",
                    "formatId": "2",
                    "duration": "00:00:30:00",
                    "mediaForm": 1,
                    "logoBurnedIn": False,
                    "logoId": "",
                    "logoPositionId": "",
                    "tracks": [
                        {
                            "trackId": "1",
                            "languageId": "1"
                        },
                        {
                            "trackId": "2",
                            "languageId": "1"
                        }
                    ]
                },
                "subtitle": [],
                "preview": [],
                "images": [
                    {
                        "useParentImage": False,
                        "line1": "",
                        "line2": "",
                        "fileName": "Ironman3_land.jpg",
                        "usageTypeId": "3",
                        "logoId": None,
                        "fileurl": "https://s3-ap-southeast-1.amazonaws.com/ams-astro-tvp-nonprod/staging/images/Ironman3_land.jpg",
                        "dimensions": "1920:1080",
                        "originalChecksum": "994c6410ce7c6640ce7bd7e4e2f69909",
                        "previewChecksum": "994c6410ce7c6640ce7bd7e4e2f69909",
                        "channelId": None,
                        "profileId": [
                            "1"
                        ],
                        "serviceId": [
                            "34"
                        ],
                        "previewImages": [
                            "https://s3-ap-southeast-1.amazonaws.com/ams-astro-tvp-nonprod/staging/previewImage/SOTT(OTT_STV)_Ironman3_land.jpg"
                        ]
                    },
                    {
                        "useParentImage": False,
                        "line1": "",
                        "line2": "",
                        "fileName": "IronMan3_port.jpg",
                        "usageTypeId": "3",
                        "logoId": None,
                        "fileurl": "https://s3-ap-southeast-1.amazonaws.com/ams-astro-tvp-nonprod/staging/images/IronMan3_port.jpg",
                        "dimensions": "1080:1620",
                        "originalChecksum": "15f089642c57eb9bbe2f32c499ecdcd2",
                        "previewChecksum": "15f089642c57eb9bbe2f32c499ecdcd2",
                        "channelId": None,
                        "profileId": [
                            "1"
                        ],
                        "serviceId": [
                            "34"
                        ],
                        "previewImages": [
                            "https://s3-ap-southeast-1.amazonaws.com/ams-astro-tvp-nonprod/staging/previewImage/SOTT(OTT_STV)_IronMan3_port.jpg"
                        ]
                    }
                ],
                "subtitlteMapping": [],
                "audioMapping": [
                    {
                        "languageId": "1",
                        "trackId": "1",
                        "serviceId": [
                            "34"
                        ]
                    }
                ]
            },
            "linearSchedule": {
                "linearChargeCodeTypeId": "",
                "vodChargeCodePricecId": "",
                "epgLink": []
            },
            "metaData": {
                "englishMetadata": {
                    "title": "Tokyo 2020 Olympic: Daily Highlights Day -2 Part 1",
                    "certification": "3",
                    "audioLanguage": "5",
                    "yearOfRealease": "2021",
                    "filter": "8",
                    "subFilter": [
                        "57",
                        "58"
                    ],
                    "genre": "24",
                    "subGenre": "209",
                    "shortSynopsis": "Catch up on Olympics Tokyo 2020 daily Highlights, actions and best moments - Day -02 Part 1",
                    "longSynopsis": "Catch up on Olympics Tokyo 2020 daily Highlights, actions and best moments - Day -02 Part 1",
                    "contentId": "12931"
                },
                "vernacular": [
                    {
                        "languageId": "1"
                    },
                    {
                        "languageId": "2"
                    },
                    {
                        "languageId": "3"
                    },
                    {
                        "languageId": "4"
                    },
                    {
                        "languageId": "5"
                    }
                ]
            },
            "hierarchyParents": [
                {
                    "parentId": "144284",
                    "title": "L1 - Dummy L1",
                    "typeId": 1,
                    "parentTypeId": "1"
                }
            ]
        }
    })
    headers = {
        'Content-Type': 'application/json'
    }
    response = requests.request("POST", url, headers=headers, data=payload)
    data = json.loads(response.content)
    return data


# res = savePackageapi(packagename, pkgid)
# print(res['packId'])
