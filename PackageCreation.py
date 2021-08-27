import requests
import json
from openpyxl import load_workbook
from saveContent import *
from savePackageData import *
from getPackageTypeApi import *
from getMediaApi import *
from offerApi import *

wb = load_workbook(r'C:\Users\MTMUNIAN\Desktop\Python\Excel\AMS.xlsx')

sheet = wb.active

i = 3
for row in sheet.iter_rows(min_row=4, max_col=1):
    for cell in row:
        i += 1

        # #################--- IF status = MP ---##################
        if sheet['A' + str(i)].value == 'MP':

            # #################--- (MP) Title ---##################
            if sheet['J' + str(i)].value is None and sheet['K' + str(i)].value is None:

                # #################################--- contentSave API Call Function---##################################

                # #################-- Set contentId 0 for new content or update if ACID exist---##################
                if sheet['C' + str(i)].value is None:
                    contentid = 0
                else:
                    contentid = sheet['C' + str(i)].value

                systemtitle = sheet['E' + str(i)].value

                # #################-- Set Episodic value 0 or 1---##################
                if sheet['F' + str(i)].value == 'Y':
                    episodic = 1
                else:
                    episodic = 0

                if not sheet['G' + str(i)].value is None:
                    recordtype = sheet['G' + str(i)].value
                else:
                    recordtype = ''

                if not sheet['H' + str(i)].value is None:
                    yearofrelease = sheet['H' + str(i)].value
                else:
                    yearofrelease = ''

                if not sheet['I' + str(i)].value is None:
                    originallanguage = sheet['I' + str(i)].value
                else:
                    originallanguage = ''

                # #################-- Set Synopsis---##################
                contentsynopsis = []
                if not sheet['L' + str(i)].value is None:
                    shtSynopEng = {"contentSynopsisId": "0", "synopsisType": "Short", "synopsis": "" + sheet['L' + str(i)].value + "", "synopsisLanguage": "English"}
                    contentsynopsis = [shtSynopEng]
                if not sheet['M' + str(i)].value is None:
                    lngSynopEng = {"contentSynopsisId": "0", "synopsisType": "Long", "synopsis": "" + sheet['M' + str(i)].value + "", "synopsisLanguage": "English"}
                    contentsynopsis = [shtSynopEng, lngSynopEng]
                if not sheet['N' + str(i)].value is None:
                    shtSynopMly = {"contentSynopsisId": "0", "synopsisType": "Short", "synopsis": "" + sheet['N' + str(i)].value + "", "synopsisLanguage": "Malay"}
                    contentsynopsis = [shtSynopEng, lngSynopEng, shtSynopMly]
                if not sheet['O' + str(i)].value is None:
                    lngSynopMly = {"contentSynopsisId": "0", "synopsisType": "Long", "synopsis": "" + sheet['O' + str(i)].value + "", "synopsisLanguage": "Malay"}
                    contentsynopsis = [shtSynopEng, lngSynopEng, shtSynopMly, lngSynopMly]
                if not sheet['P' + str(i)].value is None:
                    shtSynopChi = {"contentSynopsisId": "0", "synopsisType": "Short", "synopsis": "" + sheet['P' + str(i)].value + "", "synopsisLanguage": "Chinese"}
                    contentsynopsis = [shtSynopEng, lngSynopEng, shtSynopMly, lngSynopMly, shtSynopChi]
                if not sheet['Q' + str(i)].value is None:
                    lngSynopChi = {"contentSynopsisId": "0", "synopsisType": "Long", "synopsis": "" + sheet['Q' + str(i)].value + "", "synopsisLanguage": "Chinese"}
                    contentsynopsis = [shtSynopEng, lngSynopEng, shtSynopMly, lngSynopMly, shtSynopChi, lngSynopChi]
                if not sheet['R' + str(i)].value is None:
                    shtSynopTam = {"contentSynopsisId": "0", "synopsisType": "Short", "synopsis": "" + sheet['R' + str(i)].value + "", "synopsisLanguage": "Tamil"}
                    contentsynopsis = [shtSynopEng, lngSynopEng, shtSynopMly, lngSynopMly, shtSynopChi, lngSynopChi, shtSynopTam]
                if not sheet['S' + str(i)].value is None:
                    lngSynopTam = {"contentSynopsisId": "0", "synopsisType": "Long", "synopsis": "" + sheet['S' + str(i)].value + "", "synopsisLanguage": "Tamil"}
                    contentsynopsis = [shtSynopEng, lngSynopEng, shtSynopMly, lngSynopMly, shtSynopChi, lngSynopChi, shtSynopTam, lngSynopTam]
                if not sheet['T' + str(i)].value is None:
                    shtSynopInd = {"contentSynopsisId": "0", "synopsisType": "Short", "synopsis": "" + sheet['T' + str(i)].value + "", "synopsisLanguage": "Bahasa Indonesia"}
                    contentsynopsis = [shtSynopEng, lngSynopEng, shtSynopMly, lngSynopMly, shtSynopChi, lngSynopChi, shtSynopTam, lngSynopTam, shtSynopInd]
                if not sheet['U' + str(i)].value is None:
                    lngSynopInd = {"contentSynopsisId": "0", "synopsisType": "Long", "synopsis": "" + sheet['U' + str(i)].value + "", "synopsisLanguage": "Bahasa Indonesia"}
                    contentsynopsis = [shtSynopEng, lngSynopEng, shtSynopMly, lngSynopMly, shtSynopChi, lngSynopChi, shtSynopTam, lngSynopTam, shtSynopInd, lngSynopInd]
                if not sheet['V' + str(i)].value is None:
                    shtSynopTh = {"contentSynopsisId": "0", "synopsisType": "Short", "synopsis": "" + sheet['V' + str(i)].value + "", "synopsisLanguage": "Thai"}
                    contentsynopsis = [shtSynopEng, lngSynopEng, shtSynopMly, lngSynopMly, shtSynopChi, lngSynopChi, shtSynopTam, lngSynopTam, shtSynopInd, lngSynopInd, shtSynopTh]
                if not sheet['W' + str(i)].value is None:
                    lngSynopTh = {"contentSynopsisId": "0", "synopsisType": "Long", "synopsis": "" + sheet['W' + str(i)].value + "", "synopsisLanguage": "Thai"}
                    contentsynopsis = [shtSynopEng, lngSynopEng, shtSynopMly, lngSynopMly, shtSynopChi, lngSynopChi, shtSynopTam, lngSynopTam, shtSynopInd, lngSynopInd, shtSynopTh, lngSynopTh]

                # #################-- Content Title---##################
                contenttitle = []
                if not sheet['X' + str(i)].value is None:
                    titleEng = {"contentTitleId": "0", "title": "" + sheet['X' + str(i)].value + "", "titleType": "Standard", "titleLanguage": "English"}
                    contenttitle = [titleEng]
                if not sheet['Y' + str(i)].value is None:
                    titleMly = {"contentTitleId": "0", "title": "" + sheet['Y' + str(i)].value + "", "titleType": "Standard", "titleLanguage": "Malay"}
                    contenttitle = [titleEng, titleMly]
                if not sheet['Z' + str(i)].value is None:
                    titleChi = {"contentTitleId": "0", "title": "" + sheet['Z' + str(i)].value + "", "titleType": "Standard", "titleLanguage": "Chinese"}
                    contenttitle = [titleEng, titleMly, titleChi]
                if not sheet['AA' + str(i)].value is None:
                    titleTam = {"contentTitleId": "0", "title": "" + sheet['AA' + str(i)].value + "", "titleType": "Standard", "titleLanguage": "Tamil"}
                    contenttitle = [titleEng, titleMly, titleChi, titleTam]
                if not sheet['AB' + str(i)].value is None:
                    titleInd = {"contentTitleId": "0", "title": "" + sheet['AB' + str(i)].value + "", "titleType": "Standard", "titleLanguage": "Bahasa Indonesia"}
                    contenttitle = [titleEng, titleMly, titleChi, titleTam, titleInd]
                if not sheet['AC' + str(i)].value is None:
                    titleTh = {"contentTitleId": "0", "title": "" + sheet['AC' + str(i)].value + "", "titleType": "Standard", "titleLanguage": "Thai"}
                    contenttitle = [titleEng, titleMly, titleChi, titleTam, titleInd, titleTh]

                # #################-- Content Genre---##################
                if not sheet['AD' + str(i)].value is None and not sheet['AE' + str(i)].value is None:
                    genreDVB = {"contentGenreId": "0", "genreType": "DVB", "genre": "" + sheet['AD' + str(i)].value + "", "subgenre": "" + sheet['AE' + str(i)].value}
                    contentgenre = [genreDVB]
                else:
                    print('Genre is missing')
                    break

                if not sheet['AF' + str(i)].value is None and not sheet['AG' + str(i)].value is None:
                    filter1 = {"contentGenreId": "0", "genreType": "Category/Filter", "genre": "" + sheet['AF' + str(i)].value + "", "subgenre": "" + sheet['AG' + str(i)].value}
                    contentgenre = [genreDVB, filter1]
                else:
                    print('Filter1 is missing')
                    break

                if not sheet['AH' + str(i)].value is None and not sheet['AI' + str(i)].value is None:
                    filter2 = {"contentGenreId": "0", "genreType": "Category/Filter", "genre": "" + sheet['AH' + str(i)].value + "", "subgenre": "" + sheet['AI' + str(i)].value}
                    contentgenre = [genreDVB, filter1, filter2]
                else:
                    print('Filter2 is missing')
                    break

                # #################-- Content Talent---##################
                contenttalent = []
                if not sheet['AJ' + str(i)].value is None:
                    talActEng = {"contentTalentId": "0", "talentType": "Actor", "talentName": "" + sheet['AJ' + str(i)].value + "", "talentLanguage": "English", "billingOrder": "0"}
                    contenttalent = [talActEng]
                if not sheet['AK' + str(i)].value is None:
                    talActMly = {"contentTalentId": "0", "talentType": "Actor", "talentName": "" + sheet['AK' + str(i)].value + "", "talentLanguage": "Malay", "billingOrder": "0"}
                    contenttalent = [talActEng, talActMly]
                if not sheet['AL' + str(i)].value is None:
                    talActChi = {"contentTalentId": "0", "talentType": "Actor", "talentName": "" + sheet['AL' + str(i)].value + "", "talentLanguage": "Chinese", "billingOrder": "0"}
                    contenttalent = [talActEng, talActMly, talActChi]
                if not sheet['AM' + str(i)].value is None:
                    talActTam = {"contentTalentId": "0", "talentType": "Actor", "talentName": "" + sheet['AM' + str(i)].value + "", "talentLanguage": "Tamil", "billingOrder": "0"}
                    contenttalent = [talActEng, talActMly, talActChi, talActTam]

                # ########## saveContentapi API call function ###############
                contentRes = saveContentapi(contentid, systemtitle, episodic, recordtype, yearofrelease, originallanguage, contentsynopsis, contenttitle, contentgenre, contenttalent)
                print(contentRes['status'], contentRes['responseMessage'])
                contentid = contentRes['status']

                # ################### ----- Save Excel with changes ACID Number------- ####################
                sheet['C' + str(i)] = contentid

                # #################-- Set pkgId for new package or update if pkgId exist---##################
                if sheet['AR' + str(i)].value is None:
                    pkgid = 0
                else:
                    pkgid = sheet['AR' + str(i)].value

                # ################### ----- Prepare package name PackageType + PackageName + ChannelOwner------- ####################
                if not sheet['AS' + str(i)].value is None and not sheet['B' + str(i)].value is None and not sheet['AU' + str(i)].value is None:
                    packagename = sheet['AS' + str(i)].value + ' - ' + sheet['B' + str(i)].value + ' (' + sheet['AU' + str(i)].value + ')'
                    # print(packagename)
                else:
                    print("Package Type | Package Name | ChannelOwner is Empty")
                    break

                # ################### ----- Set packagetype------- ####################
                packageTypeRes = getPackageTypeApi()
                for packageType in packageTypeRes['packageType']:
                    if packageType['packageName'] == sheet['AS' + str(i)].value:
                        pkgtypeid = packageType['packageId']
                        # print(packageType['packageId'])

                # ################### ----- Set channelOwner  (single and multiple value)------- ####################
                if not sheet['AU' + str(i)].value is None:
                    channelOwnerExcel = sheet['AU' + str(i)].value
                    channelOwnerExcelDict = channelOwnerExcel.split("|")
                    channelownerArray = []
                    for channelOwn in channelOwnerExcelDict:
                        # print(channelOwn)
                        channelOwnerRes = getPackageTypeApi()
                        for channelOwner in channelOwnerRes['channels']:
                            if channelOwner['channelName'] == channelOwn:
                                # print(channelOwner['channelId'])
                                channelownerArray.append("" + str(channelOwner['channelId']) + "")
                    channelowner = channelownerArray
                    # print(channelowner)
                else:
                    print('Channel Owner is Empty')
                    break

                # ################### ----- Set autoPublish------ ####################
                if sheet['AT' + str(i)].value == 'Y':
                    autopublish = True
                else:
                    autopublish = False
                # print(autopublish)

                # ################### ----- Set Classification (single and multiple value)------ ####################
                if not sheet['AV' + str(i)].value is None:
                    classExcel = sheet['AV' + str(i)].value
                    classExcelDict = classExcel.split("|")
                    classExcelArray = []
                    for classification in classExcelDict:
                        classificationRes = getPackageTypeApi()
                        for classs in classificationRes['classification']:
                            if classs['classificationName'] == classification:
                                classExcelArray.append("" + str(classs['classificationId']) + "")
                    pkgclassification = classExcelArray
                    # print(pkgclassification)
                else:
                    print('Classification is Empty')
                    break

                # ################### ----- Set BoxsetTypeId ------ ####################
                if not sheet['AW' + str(i)].value is None:
                    boxsetTypeRes = getPackageTypeApi()
                    for boxsetType in boxsetTypeRes['boxsetType']:
                        if boxsetType['boxsetName'] == sheet['AW' + str(i)].value:
                            boxsettypeid = boxsetType['boxsetId']
                            # print(boxsetType['boxsetId'])
                else:
                    boxsettypeid = ''

                # ################### ----- Sort Order ------ ####################
                if not sheet['AX' + str(i)].value is None:
                    sortorder = sheet['AX' + str(i)].value
                else:
                    sortorder = ''

                # ################### ----- Acquisition Start & End ------ ####################
                acqstart = ''
                acqend = ''
                if not sheet['AY' + str(i)].value is None:
                    acqstart = sheet['AY' + str(i)].value.strftime("%Y-%m-%d %H:%M:%S")

                if not sheet['AZ' + str(i)].value is None:
                    acqend = sheet['AZ' + str(i)].value.strftime("%Y-%m-%d %H:%M:%S")

                # ################### ----- Sponsored ------ ####################
                if sheet['BA' + str(i)].value == 'Y':
                    sponsored = True
                else:
                    sponsored = False

                # ################### ----- Supplier ------ ####################
                if not sheet['BB' + str(i)].value is None:
                    supplieridRes = getPackageTypeApi()
                    for supplier in supplieridRes['supplier']:
                        if supplier['supplierName'] == sheet['BB' + str(i)].value:
                            supplierid = supplier['supplierId']
                else:
                    supplierid = ''

                # ################### ----- Offer Creation ------ ####################

                # ################### ----- STB ------ ####################
                pkgoffers = []
                if not sheet['BC' + str(i)].value is None:
                    pkgOfferSTBRes = offerApi()
                    for serviceAPISTBList in pkgOfferSTBRes['service']:
                        if serviceAPISTBList['serviceLabel'] == sheet['BC' + str(i)].value:
                            serviceSTB = serviceAPISTBList['serviceId']

                    offerStartSTB = sheet['BD' + str(i)].value.strftime("%Y-%m-%d %H:%M:%S")

                    offerEndSTB = sheet['BE' + str(i)].value.strftime("%Y-%m-%d %H:%M:%S")

                    pkgOfferSTB = {"serviceId": "" + str(serviceSTB) + "", "offerStart": "" + offerStartSTB + "", "offerEnd": "" + offerEndSTB + ""}
                    #pkgoffers = [pkgOfferSTB]
                    print(pkgOfferSTB)


                # ################### ----- IVP ------ ####################

                # ################### ----- SOTT ------ ####################

                # ########## savePackageapi API call function ###############
                packageRes = savePackageapi(packagename, pkgid, pkgtypeid, autopublish, channelowner, pkgclassification, boxsettypeid, sortorder, acqstart, acqend, sponsored, supplierid, contentid, pkgoffers)
                print(packageRes['packId'])

            # #################--- (MP) Season ---##################
            elif not sheet['J' + str(i)].value is None and sheet['K' + str(i)].value is None:
                print('Season')

            # #################--- (MP) Episode ---##################
            else:
                print('Episode')

        # #################--- IF status = M ---##################
        elif sheet['A' + str(i)].value == 'M':

            # #################--- (M) ACID Empty ---##################
            if sheet['C' + str(i)].value is None:
                print('M ACID Empty')

            # #################--- (M) ACID Not-Empty ---##################
            else:
                print('M ACID not Empty')

        # #################--- IF status = P ---##################
        elif sheet['A' + str(i)].value == 'P':

            # #################--- (P) ACID Empty ---##################
            if sheet['C' + str(i)].value is None:
                print('P ACID Empty')

            # #################--- (P) ACID Not-Empty ---##################
            else:
                print('P ACID not Empty')

        # #################--- Print error if wrong value ---##################
        else:
            print('Wrong value for status')

    wb.save('AMS_Status.xlsx')
