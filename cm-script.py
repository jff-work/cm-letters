import copy
import json
import math
import os
import shutil
import time
import zipfile
import pycountry
import requests
import selenium.common.exceptions
from bs4 import BeautifulSoup
from selenium.webdriver.common.keys import Keys
from seleniumrequests import Firefox
from selenium.webdriver.support.select import Select
from selenium.webdriver.firefox.options import Options
import PyPDF2
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

from _global_parms import *
from post_parms import *


# Split names into 3 lines of 30 chars with proper separation
def split_names(order):
    name = order['shippingAddress']['name']
    newname = ['', '', '']

    #If the name is longer than 90 Chars, shorten to 90 and display warning
    if len(name) > 91:
        newname[0] = name[:30]
        newname[1] = name[30:60]
        newname[2] = name[60:90]
        print('WARNING!! Name in Order Nr: ' + order['idOrder'] + ' longer than 90 charachters, shortened TO:: ' + name[:90] + ' FROM:: ' + name)

    #Try to separate it neatly
    else:
        namelist = name.split()
        nlit = iter(namelist)
        el = next(nlit, 'endlist')
        for i in range(3):
            if newname[i] == '' and len(newname[i] + el) < 31 and el != 'endlist':
                newname[i] = el
                el = next(nlit, 'endlist')
            while len(newname[i] + el) < 30 and el != 'endlist':
                newname[i] = newname[i] + ' ' + el
                el = next(nlit, 'endlist')
            #If we're here it didn't work, so separate it not neatly
            if i == 2 and el != 'endlist':
                newname[0] = name[:30]
                newname[1] = name[30:60]
                newname[2] = name[60:90]

    return (newname)


# Split address + extra into 3 lines of 30 chars
def split_addresses(order):
    address = order['shippingAddress']['street']
    extra = order['shippingAddress']['extra']
    newadd = ['', '', '']
    if extra == None:
        exlen = 0
    else:
        exlen = len(extra)

    #If no extra, we can just use the process used for long names. We do the same if the Address is too long, then the extra is omitted and an error generated.
    if (extra != None and len(address) > 60) or (extra != None and (len(address) + exlen > 90)):
        print('WARNING!! Extra address information: ' + extra + ' was omitted because address was too long!')
    if extra == None or len(address) > 60 or (len(address) + exlen > 90):
        #If the name is longer than 90 Chars, shorten to 90 and display warning
        if len(address) > 91:
            newadd[0] = address[:30]
            newadd[1] = address[30:60]
            newadd[2] = address[60:90]
            print('WARNING!! Street Address in Order Nr: ' + order['idOrder'] + ' longer than 90 charachters, shortened TO:: ' + address[:90] + ' FROM:: ' + address)

        #Try to separate it neatly
        else:
            addresslist = address.split()
            nlit = iter(addresslist)
            el = next(nlit, 'endlist')
            for i in range(3):
                if newadd[i] == '' and len(newadd[i] + el) < 31 and el != 'endlist':
                    newadd[i] = el
                    el = next(nlit, 'endlist')
                while len(newadd[i] + el) < 30 and el != 'endlist':
                    newadd[i] = newadd[i] + ' ' + el
                    el = next(nlit, 'endlist')
                #If we're here it didn't work, so separate it not neatly
                if i == 2 and el != 'endlist':
                    newadd[0] = address[:30]
                    newadd[1] = address[30:60]
                    newadd[2] = address[60:90]
    #We have an extra and there's space to print it
    else:
        #The easy cases where we can divide it on different lines
        #Two One-Liners
        if len(address) < 31 and exlen < 31:
            newadd[0] = address
            newadd[1] = extra
        #Two Address, One Extra
        elif len(address) < 61 and exlen < 31:
            addresslist = address.split()
            nlit = iter(addresslist)
            el = next(nlit, 'endlist')
            for i in range(2):
                if newadd[i] == '' and len(newadd[i] + el) < 31 and el != 'endlist':
                    newadd[i] = el
                    el = next(nlit, 'endlist')
                while len(newadd[i] + el) < 30 and el != 'endlist':
                    newadd[i] = newadd[i] + ' ' + el
                    el = next(nlit, 'endlist')
                # If we're here it didn't work, so separate it not neatly
                if i == 1 and el != 'endlist':
                    newadd[0] = address[:30]
                    newadd[1] = address[30:60]
            newadd[2] = extra
        #One Address, Two Extra
        elif len(address) < 31 and exlen < 61:
            newadd[0] = address
            extralist = extra.split()
            nlit = iter(extralist)
            el = next(nlit, 'endlist')
            for i in range(1,3):
                if newadd[i] == '' and len(newadd[i] + el) < 31 and el != 'endlist':
                    newadd[i] = el
                    el = next(nlit, 'endlist')
                while len(newadd[i] + el) < 30 and el != 'endlist':
                    newadd[i] = newadd[i] + ' ' + el
                    el = next(nlit, 'endlist')
                # If we're here it didn't work, so separate it not neatly
                if i == 1 and el != 'endlist':
                    newadd[1] = address[:30]
                    newadd[2] = address[30:60]
        #Leaves the case where we have space but can't neatly space it
        else:
            newadd[0] = (address + extra)[:30]
            newadd[1] = (address + extra)[30:60]
            newadd[2] = (address + extra)[60:90]
    return(newadd)


#::Get all relevant info on an order and return it as a dictionary.
def get_order_infos(ordernr, state, s, base):#, cookies=None):
    time.sleep(0.3)
    infdict = {}
    infp = s.get(base + u'/Orders/' + str(ordernr))#, cookies=cookies)
    soup = BeautifulSoup(infp.text, features="html.parser")
    retryoinf = 0
    while ((soup.find(id="SummaryRow") == None) and retryoinf < 5):
        print('Connection to MKM failed. Please close all MKM windows in browsers! Retrying in 40 seconds. Retrying ' + str(
            5 - retryoinf) + ' more times. Press Ctrl + C to cancel.')
        time.sleep(40)
        infp = s.get(base + u'/Orders/' + str(ordernr))#, cookies=cookies)
        soup = BeautifulSoup(infp.text, features="html.parser")

    #::Only if order was paid
    if state != 1 and state != 128:
        #::Date + Time Bought
        datb = soup.find(id="Timeline").div.div.next_sibling.span.string
        timb = soup.find(id="Timeline").div.div.next_sibling.span.next_sibling.string
        dattim = time.strptime(datb + '_' + timb, '%d.%m.%Y_%H:%M')
        infdict.update({"datePaid": dattim})
        #::Address - Name
        infdict.update({'AddName': soup.find(id="ShippingAddress").find(class_="Name").string})
        #::Address - Street
        infdict.update({'AddStreet': soup.find(id="ShippingAddress").find(class_="Street").string})
        #::Address - City + PLZ
        citystring = soup.find(id="ShippingAddress").find(class_="City").string
        if citystring.find(' ') > 0:
            infdict.update({'AddPLZ': citystring[:citystring.find(' ')]})
            infdict.update({'AddCity': citystring[citystring.find(' ') + 1:]})
        else:
            infdict.update({'AddPLZ': citystring})
            infdict.update({'AddCity': citystring})
        #::Address - Country
        infdict.update({'AddCountry': pycountry.countries.search_fuzzy(
            soup.find(id="ShippingAddress").find(class_="Country").string)[0].alpha_2})
        #::Address - Extra
        if soup.find(id="ShippingAddress").find(class_="Extra") != None:
            infdict.update({'AddExtra': soup.find(id="ShippingAddress").find(class_="Extra").string})
        else:
            infdict.update({'AddExtra': None})
    else:
        infdict.update({"datePaid": None})
        infdict.update({'AddName': None})
        infdict.update({'AddName': None})
        infdict.update({'AddStreet': None})
        infdict.update({'AddPLZ': None})
        infdict.update({'AddCity': None})
        infdict.update({'AddCountry': None})
        infdict.update({'AddExtra': None})

    #::Costs
    item_eur = soup.find(id="SummaryRow").div.span.strong.string.replace('.', '').replace(',', '.')
    infdict.update({'articleValue': item_eur[1:-3]})
    shipping_eur = soup.find("span", {"class": "shipping-price"}).string.replace('.', '').replace(',', '.')
    infdict.update({'shippingValue': shipping_eur[1:-3]})

    #::Temp Email
    try:
        infdict.update({"tempEmail": soup.find(id="collapsibleOtherInfo").find(text="Email:").parent.next_sibling.span.string})
    except AttributeError:
        infdict.update({"tempEmail": ""})

    #::Trustee Service + Letter Stuff
    if soup.find(id="collapsibleOtherInfo").find(class_="text-danger") != None:
        infdict.update({"isInsured": False})
    else:
        infdict.update({"isInsured": True})
    infdict.update({'shipLetterType': soup.find(id='collapsibleOtherInfo').dl.dd.a.next_sibling.string})
    infdict.update({'shipMaxWeight': soup.find(id='collapsibleOtherInfo').dl.dd.a.next_sibling.next_sibling.string})

    #::Number of Articles
    infdict.update({"articleCount": soup.find(id="collapsibleSellerShipmentSummary").div['data-article-count']})

    #::Is it a presale?
    if soup.find(class_='notification') != None:
        infdict.update({"isPresale": True})
    else:
        infdict.update({"isPresale": False})

    return infdict


#::Create a list with ID of all paid orders.
def create_order_id_list(state, s, base):
    pages = 0
    totpages = 1
    orderlist = []
    ordersort = {
        1: 'Unpaid',
        2: 'Paid',
        4: 'Sent',
        8: 'Arrived',
        32: 'NotArrived',
        128: 'Cancelled',
    }

    if sender['Presale']:
        upresale = u'presaleStatus=2&'
    else:
        upresale = u''
    while pages < totpages:
        sales = s.get(base + u'/Orders/Sales/' + ordersort[state] + u'?' + upresale + u'perSite=30&site=' + str(pages + 1))
        a = sales.text
        anchor = 0
        for i in range(0, a.count('<div data-url="/en/Magic/Orders/')):
            anchor = a.find('<div data-url="/en/Magic/Orders/', anchor) + 1
            ordstrt = a.find('rs/', anchor) + 3
            ordfin = a.find('"', ordstrt)
            orderlist.append(a[ordstrt:ordfin])
        if math.floor(len(orderlist) / 30) - totpages + 1 > 0:
            totpages = totpages + 1
        pages = pages + 1
    return orderlist


#::Build a full dictionary formatted just like the JSON provided by the API
def get_full_order_dict(orderlist, state, s, base):#, cookies=None):
    all_order_dict = {
        'order': [],
    }
    for i in orderlist:
        ordict = get_order_infos(i, state, s, base)#, cookies)
        d = {
            'y': str(ordict["datePaid"].tm_year),
            'm': str(ordict["datePaid"].tm_mon),
            'd': str(ordict["datePaid"].tm_mday),
            'h': str(ordict["datePaid"].tm_hour),
            'mi': str(ordict["datePaid"].tm_min),
        }
        for j in d:
            if len(d[j]) == 1:
                d[j] = '0' + d[j]
        oinfdict = {
            "idOrder": int(i),
            "state": {
                "datePaid": d['y'] + '-' + d['m'] + '-' + d['d'] + 'T' + d['h'] + ':' + d['mi'] + ':00+0100',
            },
            "shippingMethod": {
                "name": ordict['shipLetterType'],
                "maxWeight": ordict['shipMaxWeight'],
                "price": ordict['shippingValue'],
                "isInsured": ordict["isInsured"],
            },
            "isPresale": ordict["isPresale"],
            "temporaryEmail": ordict["tempEmail"],
            "shippingAddress": {
                "name": ordict['AddName'],
                "extra": ordict['AddExtra'],
                "street": ordict['AddStreet'],
                "zip": ordict['AddPLZ'],
                "city": ordict['AddCity'],
                "country": ordict['AddCountry'],
            },
            "articleCount": ordict["articleCount"],
            "articleValue": ordict["articleValue"],
        }
        all_order_dict['order'].append(oinfdict)

    return all_order_dict


# Write the requisite CN22 customs file for all 'Paid' orders
def create_cn22_csv(allord, state, single_artn=0):
    pres = ''
    if sender['Presale']:
        pres = '_presales_'
    filename = ''
    sortby = 'dateBought'
    order_list = {'ch_100rp': [],
                  'ch_200rp': [],
                  'ch_100rp_insured': [],
                  '150rp': [],
                  '260rp': [],
                  '370rp': [],
                  '150rp_insured': [],
                  '260rp_insured': [],
                  '370rp_insured': [],
                  '700rp_insured': [],
                  '1200rp_insured': [],
                  'manual': []}
    if state > 1:
        sortby = 'datePaid'
    csv_position = 0

    for i in sorted(allord, key=lambda item: item['state'][sortby], reverse=False):
        #Skip presales
        if i['isPresale'] and not(sender['Presale']):
            continue

        # Make sure everything is compatible with the requirements for the .csv
        namelist = split_names(i)
        addlist = split_addresses(i)
        plz = i['shippingAddress']['zip'][:12]
        if len(i['shippingAddress']['zip']) > 12:
            print('WARNING!! PLZ of OrderNr: ' + str(i['idOrder']) + ' larger than allowed, shortened from: ' +
                  i['shippingAddress']['zip'] + ' to: ' + plz)
        city = i['shippingAddress']['city'][:30]
        if len(i['shippingAddress']['city']) > 30:
            print('WARNING!! City of OrderNr: ' + str(i['idOrder']) + ' larger than allowed, shortened from: ' +
                  i['shippingAddress']['city'] + ' to: ' + city)
        country = i['shippingAddress']['country']
        if country == 'D':
            country = 'DE'
        if len(country) != 2:
            newcountry = pycountry.countries.search_fuzzy(i['shippingAddress']['country'])[0].alpha_2
            print('WARNING!! Country code changed from: ' + country + ' to: ' + newcountry)
            country = newcountry
        full_name = namelist[0] + namelist[1]
        last_name = full_name[-full_name[::-1].find(' '):]
        first_name = full_name[:-full_name[::-1].find(' ')]
        order_split = -1
        card_quantity = int(i['articleCount'])
        stamp_price_local = 0

        # Process CH/LI orders:
        if (i['shippingAddress']['country'] == ('CH')) or (i['shippingAddress']['country'] == ('LI')):
            # Set A+ to normal orders instead
            if not(i['shippingMethod']['price'] == 6.43) and i['shippingMethod']['isInsured'] and card_quantity < 41:
                # if i['shippingMethod']['name'] == 'Tracked Letter (A-Post Plus)':
                #     print('A+ confusion with order ' + str(i['idOrder']))
                i['shippingMethod']['isInsured'] = False
                i['shippingMethod']['price'] = 1.43

            more_orders = True
            while more_orders:
                if card_quantity < 41:#Standardbrief bis 100g: 1.10 CHF
                    stamp_type = 'ch_100rp'
                    stamp_product_number = 25723
                    stamp_price_local = 100     #Change this in the future
                    more_orders = False
                elif card_quantity > 40 and card_quantity < 201: #Grossbrief bis 500g: 2.00 CHF
                    stamp_type = 'ch_200rp'
                    stamp_product_number = 25729
                    more_orders = False
                else:
                    stamp_type = 'ch_200rp'
                    stamp_product_number = 25729
                    card_quantity -= 200
                if i['shippingMethod']['isInsured']:
                    stamp_type = 'ch_100rp_insured'
                order_split += 1
                order_list[stamp_type].append({'idOrder': i['idOrder'],
                                               'orderSplitNr': order_split,
                                               'isInsured': i['shippingMethod']['isInsured'],
                                               'abroad': False,
                                               'receiverName': namelist[0],
                                               'trackingNR': '',
                                               'stampNR': stamp_product_number,
                                               'localPrice': stamp_price_local,
                                               'articleCount': i['articleCount'],
                                               'articleValue': i['articleValue'],
                                               'csvPosition': -1,
                                               'webstampInfo': {'first_name': first_name,
                                                                'last_name': last_name,
                                                                'street': addlist[0],
                                                                'extra_add': addlist[1],
                                                                'country': country,
                                                                'location': city,
                                                                'plz': plz,
                                                                'email': i['temporaryEmail']},
                                               'orderInfo': i})

        # Abroad insured orders (requires filling in CN22 by WebStamp), over 200 cards need special treatment, thus .CSV entry:
        elif i['shippingMethod']['isInsured'] or card_quantity > 200:
            if card_quantity < 5:   # Standardbrief bis 20g 1.80 CHF
                stamp_type = '150rp_insured'
                stamp_product_number = 26313
            elif card_quantity > 4 and card_quantity < 18:      # Standardbrief bis 50g 2.90 CHF
                stamp_type = '260rp_insured'
                stamp_product_number = 26313
            elif card_quantity > 17 and card_quantity < 41:     # Standardbrief bis 100g 4.00 CHF
                stamp_type = '370rp_insured'
                stamp_product_number = 26313
            elif card_quantity > 40 and card_quantity < 101:    # Grossbrief bis 250g 7.0 CHF
                stamp_type = '700rp_insured'
                stamp_product_number = 26229
            elif card_quantity > 100 and card_quantity < 201:   # Grossbrief bis 500g 12.0 CHF
                stamp_type = '1200rp_insured'
                stamp_product_number = 26235
            order_list[stamp_type].append({'idOrder': i['idOrder'],
                                           'orderSplitNr': 0,
                                           'isInsured': i['shippingMethod']['isInsured'],
                                           'abroad': True,
                                           'receiverName': namelist[0],
                                           'trackingNR': '',
                                           'stampNR': stamp_product_number,
                                           'localPrice': stamp_price_local,
                                           'articleCount': i['articleCount'],
                                           'articleValue': i['articleValue'],
                                           'csvPosition': -1,
                                           'webstampInfo': {'first_name': first_name,
                                                            'last_name': last_name,
                                                            'street': addlist[0],
                                                            'extra_add': addlist[1],
                                                            'country': country,
                                                            'location': city,
                                                            'plz': plz,
                                                            'email': i['temporaryEmail']},
                                           'orderInfo': i})


        # Uninsured abroad, The cases where we use the classic .CSV > CN22 tool.
        else:
            if card_quantity < 5:   # Standardbrief bis 20g 1.80 CHF
                stamp_type = '150rp'
                stamp_product_number = 26205
                stamp_price_local = 150
            elif card_quantity > 4 and card_quantity < 18:      # Standardbrief bis 50g 2.90 CHF
                stamp_type = '260rp'
                stamp_product_number = 26211
                stamp_price_local = 260
            elif card_quantity > 17 and card_quantity < 41:     # Standardbrief bis 100g 4.00 CHF
                stamp_type = '370rp'
                stamp_product_number = 26217
            else:
                stamp_type = 'manual'
                stamp_product_number = 00000
                print('Manual order: ' + str(i['idOrder']))
            if csv_position == 0:
                filename = temp_folder + 'MultiCN22' + sender['UN'] + pres + datetime.datetime.now().strftime("%y%m%d%H%M%S") + '.csv'
                cnfile = open(filename, "w+", encoding="Latin-1")
                cnfile.write('ADDITIONALSERVICE;PRINT PRIORITY LOGO;BARCODE;SENDER NAME 1;SENDER NAME 2;SENDER NAME 3;SENDER ADDRESS 1;SENDER ADDRESS 2;SENDER ADDRESS 3;SENDER POSTCODE;SENDER CITY;SENDER COUNTRY;SENDER CONTACT PERSON;SENDER TELEPHONE;SENDER EMAIL;SENDER VAT NO.;SENDER TAX NO.;RECEIVER NAME 1;RECEIVER NAME 2;RECEIVER NAME 3;RECEIVER ADDRESS 1;RECEIVER ADDRESS 2;RECEIVER ADDRESS 3;RECEIVER POSTCODE;RECEIVER CITY;RECEIVER COUNTRY;RECEIVER CONTACT PERSON;RECEIVER TELEPHONE;RECEIVER EMAIL;RECEIVER VAT NO.;RECEIVER TAX NO.;NATURE OF CONTENT;CUSTOMER TESTIMONIALS;GOODS CURRENCY;ARTICLE 1 DESCRIPTION;ARTICLE 1 QTY;ARTICLE 1 VALUE;ARTICLE 1 ORIGIN;ARTICLE 1 WEIGHT;ARTICLE 1 CUSTOMS TARIFF NO.;ARTICLE 1 INVOICE DESCRIPTION;ARTICLE 1 KEY;ARTICLE 1 EXPORT LICENCE NO.;ARTICLE 1 EXPORT LICENCE DATE;ARTICLE 1 MOVEMENT CERTIFICATE;ARTICLE 1 MOVEMENT CERTIFICATE NO.;ARTICLE 2 DESCRIPTION;ARTICLE 2 QTY;ARTICLE 2 VALUE;ARTICLE 2 ORIGIN;ARTICLE 2 WEIGHT;ARTICLE 2 CUSTOMS TARIFF NO.;ARTICLE 2 INVOICE DESCRIPTION;ARTICLE 2 KEY;ARTICLE 2 EXPORT LICENCE NO.;ARTICLE 2 EXPORT LICENCE DATE;ARTICLE 2 MOVEMENT CERTIFICATE;ARTICLE 2 MOVEMENT CERTIFICATE NO.;ARTICLE 3 DESCRIPTION;ARTICLE 3 QTY;ARTICLE 3 VALUE;ARTICLE 3 ORIGIN;ARTICLE 3 WEIGHT;ARTICLE 3 CUSTOMS TARIFF NO.;ARTICLE 3 INVOICE DESCRIPTION;ARTICLE 3 KEY;ARTICLE 3 EXPORT LICENCE NO.;ARTICLE 3 EXPORT LICENCE DATE;ARTICLE 3 MOVEMENT CERTIFICATE;ARTICLE 3 MOVEMENT CERTIFICATE NO.')

            csv_position += 1
            order_list[stamp_type].append({'idOrder': i['idOrder'],
                                           'orderSplitNr': 0,
                                           'isInsured': i['shippingMethod']['isInsured'],
                                           'abroad': True,
                                           'receiverName': namelist[0],
                                           'trackingNR': '',
                                           'stampNR': stamp_product_number,
                                           'localPrice': stamp_price_local,
                                           'csvPosition': -1 + csv_position,
                                           'articleCount': i['articleCount'],
                                           'articleValue': i['articleValue'],
                                           'webstampInfo': {},
                                           'orderInfo': i})

            # Create the dictionary with all the data to write into the csv
            cndict = {
                'ADDITIONALSERVICE': '0', #str(int(i['shippingMethod']['isInsured'])),  Shouldn't be insured either way here.
                'PRINT PRIORITY LOGO': '0',
                'BARCODE': '',
                'SENDER NAME 1': sender['Name'],
                'SENDER NAME 2': '',
                'SENDER NAME 3': '',
                'SENDER ADDRESS 1': sender['StreetHnr'],
                'SENDER ADDRESS 2': '',
                'SENDER ADDRESS 3': '',
                'SENDER POSTCODE': sender['PLZ'],
                'SENDER CITY': sender['City'],
                'SENDER COUNTRY': sender['Country'],
                'SENDER CONTACT PERSON': sender['Name'],
                'SENDER TELEPHONE': sender['TelNr'],
                'SENDER EMAIL': sender['Email'],
                'SENDER VAT NO.': '',
                'SENDER TAX NO.': '',
                'RECEIVER NAME 1': namelist[0],
                'RECEIVER NAME 2': namelist[1],
                'RECEIVER NAME 3': namelist[2],
                'RECEIVER ADDRESS 1': addlist[0],
                'RECEIVER ADDRESS 2': addlist[1],
                'RECEIVER ADDRESS 3': addlist[2],
                'RECEIVER POSTCODE': plz,
                'RECEIVER CITY': city,
                'RECEIVER COUNTRY': country,
                'RECEIVER CONTACT PERSON': '',
                'RECEIVER TELEPHONE': sender['TelNr'],
                'RECEIVER EMAIL': i['temporaryEmail'],
                'RECEIVER VAT NO.': '',
                'RECEIVER TAX NO.': '',
                'NATURE OF CONTENT': '2',
                'CUSTOMER TESTIMONIALS': '',
                'GOODS CURRENCY': 'EUR',
                'ARTICLE 1 DESCRIPTION': 'Trading-Cards',
                'ARTICLE 1 QTY': str(i['articleCount']),
                'ARTICLE 1 VALUE': str(i['articleValue']),
                'ARTICLE 1 ORIGIN': sender['Country'],
                'ARTICLE 1 WEIGHT': str(round(0.002 * int(i['articleCount']), 3)),
                'ARTICLE 1 CUSTOMS TARIFF NO.': '',
                'ARTICLE 1 INVOICE DESCRIPTION': '',
                'ARTICLE 1 KEY': '',
                'ARTICLE 1 EXPORT LICENCE NO.': '',
                'ARTICLE 1 EXPORT LICENCE DATE': '',
                'ARTICLE 1 MOVEMENT CERTIFICATE': '',
                'ARTICLE 1 MOVEMENT CERTIFICATE NO.': '',
                'ARTICLE 2 DESCRIPTION': '',
                'ARTICLE 2 QTY': '',
                'ARTICLE 2 VALUE': '',
                'ARTICLE 2 ORIGIN': '',
                'ARTICLE 2 WEIGHT': '',
                'ARTICLE 2 CUSTOMS TARIFF NO.': '',
                'ARTICLE 2 INVOICE DESCRIPTION': '',
                'ARTICLE 2 KEY': '',
                'ARTICLE 2 EXPORT LICENCE NO.': '',
                'ARTICLE 2 EXPORT LICENCE DATE': '',
                'ARTICLE 2 MOVEMENT CERTIFICATE': '',
                'ARTICLE 2 MOVEMENT CERTIFICATE NO.': '',
                'ARTICLE 3 DESCRIPTION': '',
                'ARTICLE 3 QTY': '',
                'ARTICLE 3 VALUE': '',
                'ARTICLE 3 ORIGIN': '',
                'ARTICLE 3 WEIGHT': '',
                'ARTICLE 3 CUSTOMS TARIFF NO.': '',
                'ARTICLE 3 INVOICE DESCRIPTION': '',
                'ARTICLE 3 KEY': '',
                'ARTICLE 3 EXPORT LICENCE NO.': '',
                'ARTICLE 3 EXPORT LICENCE DATE': '',
                'ARTICLE 3 MOVEMENT CERTIFICATE': '',
                'ARTICLE 3 MOVEMENT CERTIFICATE NO': '',
            }
            cnfile.write('\n')
            for j in cndict:
                if cndict[j] == None:
                    print(cndict[j])
                    print(j)
                # Check for Characters not allowed by Post interface and replace them.
                for k in cndict[j]:
                    if k not in LegalCharPostList:
                        if k in CharTranslateDict:
                            cndict[j] = cndict[j].replace(k, CharTranslateDict[k])
                            print('Charachter replaced ' + k + ' to ' + CharTranslateDict[k] + ' in Order Nr. ' + str(i['idOrder']))
                        else:
                            cndict[j] = cndict[j].replace(k, '')
                            print('Unexpected charachter found: ' + k + ' in Order Nr. ' + str(i['idOrder']) + '. REPLACED WITH EMPTY SPACE!')
                cnfile.write(cndict[j] + ';')
    if filename != '':
        cnfile.close()
    return filename, order_list


#Creates a new selenium driver and logs into post
def driver_init_and_post_login(login):
    ff_op = Options()
    # ff_service = Service()
    # ff_op.add_argument("--headless")
    driver = Firefox(options=ff_op)  # , firefox_profile=fp)
    driver.get('https://account.post.ch/idp/?login')
    driver.maximize_window()
    if login:
        time.sleep(3)
        driver.find_element(By.ID,'externalIDP').click()
        time.sleep(3)
        try:
            login_email_box = driver.find_element(By.ID,'email')
            login_email_box.send_keys(sender['SwissIDUN'])
            login_pw_box = driver.find_element(By.ID,'password')
            login_pw_box.send_keys(sender['SwissIDPW'])
            login_button = driver.find_element(By.CLASS_NAME,'next_icon')
            login_button.click()
            time.sleep(2)
        except:
            pass

    return driver


#Order all the webstamps
def webstamp(stamp_type_list, stamp_type, driver, preview, stamp_test=False):
    if 'ch' in stamp_type: #0 = CH, 1 = Abroad, 2 = Abroad insured
        stamp_case = 0
    elif 'insured' in stamp_type:
        stamp_case = 2
    else:
        stamp_case = 1
    stamp_product_number = stamp_type_list[0]['stampNR']
    return_format = 'c6'


    def press_continue(driver=driver):
        driver.find_element(By.CSS_SELECTOR,'span[t="button.continue"]').click()  # Continue
        return


    def press_apply(driver=driver, minus=1):
        driver.find_elements(By.CSS_SELECTOR,'span[t="button.apply"]')[-minus].click()  # Apply
        return


    def wait_click(by_type, by_path, wait_time=20, driver=driver):
        try:
            WebDriverWait(driver, wait_time).until(EC.element_to_be_clickable((by_type, by_path))).click()
        except selenium.common.exceptions.ElementClickInterceptedException:
            time.sleep(2)
            WebDriverWait(driver, wait_time).until(EC.element_to_be_clickable((by_type, by_path))).click()
        return


    #::0::Start Page
    driver.get("https://webstamp.post.ch/")
    wait_click(By.CSS_SELECTOR, 'a[t="index.create-webstamp"]')
    time.sleep(1)
    cancel_button_list = driver.find_element(By.XPATH, '//button[text()="Discard"]')
    if len(cancel_button_list) != 0:
        cancel_button_list[0].click()
        time.sleep(1)

    #::1::Format Page
    if stamp_case > 0:  # EU Orders
        wait_click(By.ID, 'ws-product-zone-label-1')
        time.sleep(1)
        press_apply(driver, minus=2)
        if stamp_case > 1: # Select insured, default A-post.
            driver.find_element(By.ID,'ws-product-delivery-26313').click()
            press_apply(driver, minus=2)
            driver.find_element(By.ID,'ws-product-format-' + str(stamp_product_number)).click()
            press_apply(driver, minus=2)
            driver.find_element(By.ID,'ws-facet-product-addition-1929').click()
            press_continue(driver)
            #::1.5::Goods Page
            time.sleep(1)
            driver.find_element(By.ID,'section_19_input').send_keys('Trading cards')  # Artikel description
            driver.find_element(By.ID,'23_input').send_keys(stamp_type_list[0]['articleValue'])  # Artikel value
            driver.find_element(By.ID,'27_input').send_keys(str(round(0.002 * int(stamp_type_list[0]['articleCount']), 3)))  # Artikel weight
            Select(driver.find_element(By.ID,'30_select')).select_by_value(sender['Country'])
            driver.find_element(By.CSS_SELECTOR,'button[type="submit"]').click()
            time.sleep(1)
            Select(driver.find_element(By.ID,'59_select')).select_by_value('goods')
            driver.find_element(By.ID,'46_input').click()
            press_continue(driver)
        else:
            driver.find_element(By.ID,'ws-product-format-' + str(stamp_product_number)).click()
            press_apply(driver, minus=2)
            press_continue(driver)
    else:  #CH Orders
        if 'insured' in stamp_type:  # Select insured, default A-post.
            driver.find_element(By.ID,'ws-product-delivery-26189').click()
            press_apply(driver)
        press_apply(driver)
        if '200rp' in stamp_type:
            driver.find_element(By.ID,'ws-product-format-25729').click()  # Select 2.00 Letter
        press_apply(driver)
        press_continue(driver)

    #::2::Image Page
    if stamp_case < 2:
        wait_click(By.CSS_SELECTOR, 'span[t="image.options.without.description"]')
        press_continue(driver)

    #::3::Sender Page
    if stamp_case == 0:
        time.sleep(2)
        driver.find_element(By.ID,'ws-sender-address-' + str(sender['PostUsrAddressID'])).click()
        # wait_click(By.ID, 'ws-sender-address-usr' + str(sender['PostUsrAddressID']))
    press_continue()

    #::4::Recipient Page
    if stamp_case != 1:
        if stamp_case == 0:
            wait_click(By.CSS_SELECTOR, 'span[t="recipients.options.with.description"]')
        time.sleep(1)
        wait_click(By.CSS_SELECTOR, 'button[t="zopa.importbutton"]')
        time.sleep(5)
        address_csv_filename = temp_folder + 'case_' + str(stamp_case) + '_' + datetime.datetime.now().strftime("%y%m%d%H%M%S") + '_temp.csv'
        address_csv_file = open(address_csv_filename, "w+", encoding="Latin-1")
        if stamp_case == 2:
            address_csv_file.write('Last name;First name;Street;Country;Postcode;Location;E-Mail\n')
            for address in stamp_type_list:
                wsi = address['webstampInfo']
                address_string = wsi['last_name'] + ';' + wsi['first_name'] + ';' + wsi['street'] + ';' + wsi['country'] + ';' + wsi['plz'] + ';' + wsi['location'] + ';' + wsi['email'] + ';\n'
                address_csv_file.write(address_string)
        else:
            address_csv_file.write('First name;Last name;Street;Country;Postcode;Location;E-Mail\n')
            for address in stamp_type_list:
                wsi = address['webstampInfo']
                address_string = wsi['first_name'] + ';' + wsi['last_name'] + ';' + wsi['street'] + ';' + wsi['country'] + ';' + wsi['plz'] + ';' + wsi['location'] + ';' + wsi['email'] + ';\n'
                address_csv_file.write(address_string)
        address_csv_file.close()
        driver.switch_to.default_content()
        zopa_frame = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CLASS_NAME, 'zopa-widget')))
        driver.switch_to.frame(zopa_frame)
        time.sleep(1)
        upload_url = driver.find_element(By.ID,'fileUpload')
        upload_url.send_keys(os.getcwd() + '/' + address_csv_filename)
        time.sleep(1)
        wait_click(By.CSS_SELECTOR, 'button[id="forward"]')
        time.sleep(1)
        wait_click(By.CSS_SELECTOR, 'button[id="forward"]')
        time.sleep(2)
        wait_click(By.CSS_SELECTOR, 'button[id="complete"]')
        driver.switch_to.default_content()
    time.sleep(1)
    press_continue()

    #::5::Format Page
    time.sleep(2)
    if stamp_case < 2:
        wait_click(By.ID, 'ws-printoptions-media-type-label-2')
    else:
        wait_click(By.ID, 'ws-printoptions-media-type-label-13')
    press_apply(driver)
    Select(driver.find_element(By.ID,'ws-envelope-media-select-select')).select_by_value('105')
    'ws-envelope-media-select-select'
    if stamp_case == 2:
        if stamp_product_number in [25785, 25779]:
            Select(driver.find_element(By.ID,'ws-envelope-media-select-select')).select_by_value('141')  # C4
            return_format = 'c4'
        else:
            time.sleep(2)
            Select(driver.find_element(By.ID,'ws-envelope-media-select-select')).select_by_value('140')  # Benutzerdefiniert, because there's no C6 here lolol

    # exit()
            driver.execute_script("arguments[0].value = 162", driver.find_element(By.ID,'ws-label-custom-width'))
            driver.execute_script("arguments[0].value = 114", driver.find_element(By.ID,'ws-label-custom-width'))
            driver.execute_script("arguments[0].value = 14", driver.find_element(By.ID,'ws-label-custom-width'))
            driver.execute_script("arguments[0].value = 12", driver.find_element(By.ID,'ws-label-custom-width'))

    # driver.execute_script("arguments[0].value = 162", driver.find_element(By.ID,'ws-label-custom-width'))
    # driver.execute_script("arguments[0].value = 114", driver.find_element(By.ID,'ws-label-custom-width'))
    # driver.execute_script("arguments[0].value = 14", driver.find_element(By.ID,'ws-label-custom-width'))
    # driver.execute_script("arguments[0].value = 12", driver.find_element(By.ID,'ws-label-custom-width'))
    else:
        Select(driver.find_element(By.ID,'ws-envelope-media-select-select')).select_by_value('105')
        return_format = 'c6'
    time.sleep(2)

    press_continue(driver)

    #::6::Payment Page
    if stamp_case == 1:
        driver.find_element(By.ID,'ws-input-order-items').clear()
        driver.find_element(By.ID,'ws-input-order-items').send_keys(len(stamp_type_list))
    wait_click(By.CLASS_NAME, 'ws-prepaid-label')
    stamp_value_to_purchase = driver.find_element(By.ID,'ws-input-order-totalprice').get_attribute('value')

    local_storage = json.loads(driver.execute_script("return window.localStorage.getItem('appState');"))
    ord = local_storage['currentOrder']
    w_testf('local_store.json', json.dumps(ord, indent=2))

    preview_request_string = '{"lsoRequest": null, "contentType": "", "unregisteredReturn": null, "items": '
    preview_request_string += str(len(stamp_type_list)) + ', "firstLabel": '
    preview_request_string += str(ord['firstLabel']) + ', "printOrder": '
    preview_request_string += str(ord['printOrder']).lower() + ','
    try:
        if (ord['imageUUID'] == None):
            preview_request_string += '"imageUUID": null,'
    except KeyError:
        pass
    try:
        preview_request_string += '"isImportAddress": ' + str(ord['isImportAddress']).lower() + ','
    except KeyError:
        pass
    preview_request_string += '"paymentType": "prepaid", "fileFormat": "pdf", "product": {"id": '
    preview_request_string += str(ord['productId']) + '}, "systemMedia": {"id": '
    preview_request_string += str(ord['systemMedia']['id']) + '}'
    if ord['systemMedia']['id'] == 140:
        user_media = ord['userMedia']
        unwanted = set(user_media['mediaType']) - {'id'}
        for unwanted_key in unwanted: del user_media['mediaType'][unwanted_key]
        preview_request_string += ', "userMedia":'
        preview_request_string += json.dumps(user_media)
    if stamp_case != 1:
        preview_request_string += ',"recipients":'
        recipient_list = ord['recipients']
        for recipient_address in recipient_list:
            recipient_address.pop('hasChanged')
            recipient_address.pop('_meta')
        preview_request_string += json.dumps(recipient_list)
        preview_request_string += ', "senderAddress": {'
        if stamp_case == 0:
            preview_request_string += '"type": "'
            preview_request_string += str(ord['senderAddress']['type']) + '", "country": "'
            preview_request_string += str(ord['senderAddress']['country']) + '", "firstName": "'
            preview_request_string += str(ord['senderAddress']['firstName']) + '", "lastName": "'
            preview_request_string += str(ord['senderAddress']['lastName']) + '", "streetName": "'
            preview_request_string += str(ord['senderAddress']['streetName']) + '", "zip": "'
            preview_request_string += str(ord['senderAddress']['zip']) + '", "city": "'
            preview_request_string += str(ord['senderAddress']['city']) + '", "active": 1} '
        else:
            preview_request_string += '"id": ' + str(ord['senderAddress']['id']) + ' }'
    if stamp_case == 2:
        preview_request_string += ',"eadShipment": true, "eadOrderInfo": '
        ead_order_info = ord['eadOrderInfo']
        ead_order_info.pop('id')
        ead_order_info.pop('order')
        ead_order_info.pop('orderTemplate')
        ead_order_info.pop('senderAddress')
        ead_order_info['signature'].pop('_links')
        preview_request_string += json.dumps(ead_order_info)
    preview_request_string += '}'
    w_testf('request_string_unf.json', preview_request_string)
    # exit()
    preview_file_request = driver.request('POST', 'https://webstamp.post.ch/order-preview', data={'order': preview_request_string, 'options': 'undefined'})
    if preview_file_request.status_code != 200:
        stamp_file_input = ''
        while (stamp_file_input == ''):
            driver.find_element(By.CSS_SELECTOR,'a[t="checkout.printpreview"]').click()
            stamp_file_input = input('Please manually download file, move it to the _temp folder, then enter its FN.')
            if stamp_file_input == 'n':
                exit()
            stamp_file_link = temp_folder + stamp_file_input
    else:
        stamp_file_link = temp_folder + stamp_type + datetime.datetime.now().strftime("%y%m%d%H%M%S") + '.pdf'
        stamp_file = open(stamp_file_link, 'wb+')
        stamp_file.write(preview_file_request.content)
        stamp_file.close()
    if not preview:
        print(stamp_value_to_purchase)

        time.sleep(2)
        input_to_continue = ''
        while (input_to_continue != 'Y'):
            input_to_continue = input('Purchase Warning. Would like to purchase ' + stamp_value_to_purchase + '.- Stamps. Continue [Y/n]?')
            if input_to_continue == 'n':
                exit()
        driver.find_element(By.CSS_SELECTOR, 'span[t="button.buy"]').click()
        time.sleep(2)
        all_a = driver.find_elements(By.TAG_NAME, 'a')
        for a in all_a:
            href = a.get_attribute('href')
            if 'ws_stamps' in str(href):
                break

        stamp_file_request = driver.request('GET', href)

        if stamp_file_request.status_code != 200:
            stamp_file_input = ''
            while (stamp_file_input == ''):
                driver.find_element(By.CSS_SELECTOR,'a[t="checkout.printpreview"]').click()
                stamp_file_input = input('Please manually download file, move it to the _temp folder, then enter its FN.')
                if stamp_file_input == 'n':
                    exit()
                stamp_file_link = temp_folder + stamp_file_input
        else:
            stamp_file_link = temp_folder + stamp_type + datetime.datetime.now().strftime("%y%m%d%H%M%S") + '.pdf'
            stamp_file = open(stamp_file_link, 'wb+')
            stamp_file.write(stamp_file_request.content)
            stamp_file.close()


    ####download purchased file here

    return(return_format, stamp_file_link)


#Creates a CN22 (used for every uninsured abroad order)
def create_cn22(cn22_csv_filename, login):
    driver = driver_init_and_post_login(login)
    driver.get('https://account.post.ch/idp/?login&targetURL=https%3A%2F%2Fservice.post.ch%2Fvgkklp%2Fbegleitpapiere%2Fbegleitpapiere%2Fsecured%2F&lang=&service=&inIframe=&inMobileApp=')
    begleitpapier_button = driver.find_element(By.ID,'neuesBegleitpapierErstellen')
    begleitpapier_button.click()

    mehrfach_c22_button = driver.find_element(By.ID,'mehrfachCn22')
    mehrfach_c22_button.click()

    plz_box = driver.find_element(By.ID,'PlzOrtSearchTerm')
    plz_box.send_keys(sender['PLZ'] + ' ' + sender['City'])
    plz_box.send_keys(Keys.ENTER)
    time.sleep(1)
    continue_button = driver.find_element(By.ID,'weiterZuSchritt2Datenimport')
    continue_button.click()

    upload_url = driver.find_element(By.ID,'importFile')
    upload_url.send_keys(os.getcwd() + '/' + cn22_csv_filename)
    time.sleep(2)
    continue_button2 = driver.find_element(By.ID,'weiterZuSchritt3Abschluss')
    continue_button2.click()

    # 0 for A4, 1 for A6
    paperformat_select = Select(driver.find_element(By.ID,'Papierformat'))
    paperformat_select.select_by_index(1)
    time.sleep(1)
    # 0 for left upper, 1 for right upper, 2 for left lower, 3 for right lower
    # print_position_select = Select(driver.find_element(By.ID,'DruckenAbPosition'))
    # print_position_select.select_by_index(2)
    # time.sleep(1)
    # 1 for 21, 2 for 23
    bar_code_type_select = Select(driver.find_element(By.ID,'BarcodelistenTyp'))
    bar_code_type_select.select_by_index(1)
    time.sleep(1)
    checkbox = driver.find_element(By.ID,'GefaehrlicheGueterBestaetigung_label')
    checkbox.click()
    time.sleep(1)
    create_button = driver.find_element(By.ID,'dokumenteErstellen')
    create_button.click()
    time.sleep(4)

    zip_download_button = driver.find_element(By.ID,'zipHerunterladen')
    zip_url_online = zip_download_button.get_attribute('href')
    zip_request = driver.request('GET', zip_url_online)
    zip_url = 'zipfile_' + datetime.datetime.now().strftime("%y%m%d%H%M%S")
    zip_file = open(zip_url + '.zip', 'wb+')
    zip_file.write(zip_request.content)
    zip_file.close()

    zip_file = zipfile.ZipFile(zip_url + '.zip', 'r')
    zip_file.extractall(zip_url)
    zip_file.close()
    os.remove(zip_url + '.zip')

    input_file_name = zip_url + '/Address label.pdf'
    output_file_name = 'labels_to_print/' + sender['UN'] + '_labels_' + datetime.datetime.now().strftime("%y%m%d%H%M%S") + '.pdf'
    os.rename(input_file_name, output_file_name)
    try:
        shutil.rmtree(zip_url)
    except:
        print("Couldn't delete ZIP folder.")
    driver.quit()

    return output_file_name


#Logs into Post and creates the PDFs
def create_post_pdf_and_webstamps(filename_and_order_list, stamp_test=False, login=True):
    stamp_file_dict = {'c4': [],
                       'c6': []}

    if filename_and_order_list[0] != '':
        stamp_file_dict['c6'].append(create_cn22(filename_and_order_list[0], login=login))

    driver = driver_init_and_post_login(login=login)
    local_stamp_file = open(local_stamp_filename, 'r')
    local_stamps = json.loads(local_stamp_file.read())
    local_stamp_file.close()

    def stamp_em(stamp_file_dict, preview, stamp_type, stamp_type_list):
        w_testf(stamp_type + '.json', json.dumps(stamp_type_list))
        if not ('ch' in stamp_type) and ('insured' in stamp_type):
            for order in stamp_type_list:
                stamp_file_link_and_format = webstamp([order], stamp_type, driver, preview, stamp_test)
                stamp_file_dict[stamp_file_link_and_format[0]].append(stamp_file_link_and_format[1])
        elif not (not ('ch' in stamp_type) and preview):
            stamp_file_link_and_format = webstamp(stamp_type_list, stamp_type, driver, preview, stamp_test)
            stamp_file_dict[stamp_file_link_and_format[0]].append(stamp_file_link_and_format[1])
        return stamp_file_dict

    for stamp_type in filename_and_order_list[1]:
        stamp_type_list = filename_and_order_list[1][stamp_type]
        if stamp_type_list != []:
            if not(stamp_type == 'manual'):
                stamp_price_type = str(filename_and_order_list[1][stamp_type][0]['localPrice'])
                if stamp_price_type in local_stamps:
                    leftover_stamps = local_stamps[stamp_price_type] - len(stamp_type_list)
                    if leftover_stamps < 0:
                        print('Local ' + str(stamp_price_type) + ' used up. Used ' + str(local_stamps[stamp_price_type]) + ' of them.')
                        local_stamps.pop(stamp_price_type)
                        stamp_file_dict = stamp_em(stamp_file_dict, True, stamp_type, stamp_type_list[:leftover_stamps])
                        stamp_file_dict = stamp_em(stamp_file_dict, False, stamp_type, stamp_type_list[leftover_stamps:])
                    else:
                        print('Used ' + str(len(stamp_type_list)) + ' local ' + str(stamp_price_type) + ' stamps. ' + str(leftover_stamps) + ' remaining.')
                        stamp_file_dict = stamp_em(stamp_file_dict, True, stamp_type, stamp_type_list)
                        local_stamps[stamp_price_type] = leftover_stamps
                else:
                    stamp_file_dict = stamp_em(stamp_file_dict, False, stamp_type, stamp_type_list)

    local_stamp_file = open(local_stamp_filename, 'w')
    local_stamp_file.write(json.dumps(local_stamps, indent=2))
    local_stamp_file.close()

    # driver.quit()
    # zip_file = zipfile.ZipFile(zip_url + '.zip', 'r')
    # zip_file.extractall(zip_url)
    # zip_file.close()
    # os.remove(zip_url + '.zip')
    # Add all the PDF pieces together scaled to A4 so the printer prints them correctly.
    print_file_dict = {}
    for letter_format in stamp_file_dict:
        if stamp_file_dict[letter_format] != []:
            print_file_dict.update({letter_format: print_folder + datetime.datetime.now().strftime("%y%m%d") + '_' + letter_format + '_' + datetime.datetime.now().strftime("%H%M%S") + '.pdf'})
            writer = PyPDF2.PdfFileWriter()
            for letter_pdf in stamp_file_dict[letter_format]:
                letter_pdf_read = PyPDF2.PdfFileReader(letter_pdf)
                total_pages = letter_pdf_read.getNumPages()
                for pdf_page in range(0, total_pages):
                    page = letter_pdf_read.getPage(pdf_page)
                    if 'labels' in letter_pdf:
                        page.mediaBox.lowerLeft = (-30, -10)
                        page.cropBox.lowerLeft = (-30, -10)
                        page.mediaBox.upperRight = (812, 585)
                        page.cropBox.upperRight = (812, 585)
                    else:
                        page.mediaBox.upperRight = (842, 595)
                        page.cropBox.upperRight = (842, 595)
                    page.rotateClockwise(90)
                    writer.addPage(page)
            with open(print_file_dict[letter_format], 'wb+') as f:
                writer.write(f)

    if 'c6' in print_file_dict:
        os.system('lp ' + print_file_dict['c6']) # + ' -o orientation-requested=5')

    if 'c4' in print_file_dict:
        print('Please change to C4 envelopes!.')
        input_to_continue = ''
        while (input_to_continue != 'Y'):
            input_to_continue = input('Continue [Y/n]?')
            if input_to_continue == 'n':
                exit()
        os.system('lp ' + print_file_dict['c4'])

    return


####
error_log = []
sender['Presale'] = True
order_state = 2
single_order_number = 0
stamp_test = False            #if activated, instead uses file test_ol_2.json, which has an order list with 1 of each stamp type. Also stops purchasing.
manual_address_input = False
login = False
new_manual_addresses = {

}
manual_addresses = [{}]
#####

import logging
import http.client

httpclient_logger = logging.getLogger("http.client")

def httpclient_logging_patch(level=logging.DEBUG):
    """Enable HTTPConnection debug logging to the logging framework"""

    def httpclient_log(*args):
        httpclient_logger.log(level, " ".join(args))

    # mask the print() built-in in the http.client module to use
    # logging instead
    http.client.print = httpclient_log
    # enable debugging
    http.client.HTTPConnection.debuglevel = 1

httpclient_logging_patch()
logging.basicConfig(filename='_debug' + '/' + 'requests' + '_' + datetime.datetime.now().strftime("%y%m%d%H%M%S") + '.log', filemode='w+', level=logging.DEBUG)

if manual_address_input:
    try:
        address_log_file = open('address_log.json', 'r')
        address_log = json.loads(address_log_file.read())
        address_log_file.close()
    except FileNotFoundError:
        address_log = {}
    address_log_file = open('address_log.json', 'w+')
    for new_address in new_manual_addresses.keys():
        if not new_address in address_log:
            address_log.update({new_address: new_manual_addresses[new_address]})
    address_log_file.write(json.dumps(address_log, indent=2))
    address_log_file.close()

if not stamp_test:
    if manual_address_input:
        orders_paid = {"order": []}
        for manual_address in manual_addresses:
            orders_paid['order'].append({
                "idOrder": 0,
                "state": {
                    "datePaid": datetime.datetime.now().strftime('%d.%m.%Y_%H:%M')
                },
                "shippingMethod": {
                    "name": "Manual Choice",
                    "maxWeight": None,
                    "price": ".7",
                    "isInsured": manual_address['isInsured']
                },
                "isPresale": False,
                "temporaryEmail": "shipment-1063295591@cardmarket.com",
                "shippingAddress": address_log[manual_address['name']],
                "articleCount": manual_address['articleCount'],
                "articleValue": manual_address['articleValue']
            })

    else:
        # ::Open a session s to make sure we stay logged in.
        with requests.Session() as s:
            base = u'https://cardmarket.com/en/Magic'
            s.headers = {
                'Accept': '*/*',
                'Accept-Encoding': 'gzip, deflate, br',
                'Accept-Language': 'en-US,en;q=0.5',
                'Connection': 'keep-alive',
                'DNT': '1',
                'Host': 'www.cardmarket.com',
                # 'Origin': 'https://www.cardmarket.com',
                # 'Referer': 'https://www.cardmarket.com/en/Magic',
                # 'Sec-Fetch-Dest': 'document',
                'Sec-Fetch-Mode': 'navigate',
                'Sec-Fetch-Site': 'none',
                # 'Sec-Fetch-User': '?1',
                'Upgrade-Insecure-Requests': '1',
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/78.0.3904.108 Safari/537.36',
            }
            q = s.get(base)
            w_testf('login_pre.html', q.content, b='b')
            tkn = q.text.find('__cmtkn')
            cmtkn = q.text[tkn + 16:tkn + 80]

            data = {
                '__cmtkn': cmtkn,
                'referalPage': '/en/Magic',
                'username': sender['UN'],
                'userPassword': sender['PW'],
            }
            p = s.post(base + '/PostGetAction/User_Login', data=data)
            orders_paid = get_full_order_dict(create_order_id_list(order_state, s, base), order_state, s, base,)

    file_name_and_order_list = create_cn22_csv(orders_paid['order'], order_state, single_order_number)
    w_testf('fnol.json', json.dumps(file_name_and_order_list, indent=2))
    create_post_pdf_and_webstamps(file_name_and_order_list, stamp_test=stamp_test, login=login)

else:
    # test_ol_file = open('op2.json', 'r')
    # orders_paid = json.loads(test_ol_file.read())
    # test_ol_file.close()
    # file_name_and_order_list = create_cn22_csv(orders_paid['order'], order_state, single_order_number)
    # test_ol_file = open('test_ol_2.json', 'w+')
    # test_ol_file.write(json.dumps(file_name_and_order_list, indent=2))
    # test_ol_file.close()
    test_ol_file = open('man_fn_220913.json', 'r') #('test_ol_2.json', 'r')
    test_ol = json.loads(test_ol_file.read())
    test_ol_file.close()
    create_post_pdf_and_webstamps(test_ol, stamp_test=False, login=login)


####
if error_log != []:
    w_testf('error_log.json', json.dumps(error_log, indent=2))

print("DONE")


"""def get_order_nrs(pdf_url, nr_pp):
    pdf = open(pdf_url, "rb").read()

    startmark = b"\xff\xd8"
    startfix = 0
    endmark = b"\xff\xd9"
    endfix = 2
    i = 0
    njpg = 0
    nr_list = []

    while True:
        istream = pdf.find(b"stream", i)
        if istream < 0:
            break
        istart = pdf.find(startmark, istream, istream + 20)
        if istart < 0:
            i = istream + 20
            continue
        iend = pdf.find(b"endstream", istart)
        iend = pdf.find(endmark, iend - 20)
        istart += startfix
        iend += endfix
        jpg = pdf[istart:iend]
        str = pytesseract.image_to_string(Image.open(io.BytesIO(jpg)))
        nr_end = str.find('CH')
        if nr_end > 0:
            nr = str[nr_end - 15:nr_end + 2]
            nr_list.append(nr.replace(' ',''))
        njpg += 1
        i = iend

    if nr_pp == 4:
        for i in range(int(len(nr_list) / 4)):
            j = (i + 1) * 4 - 2
            a = nr_list[j]
            nr_list[j] = nr_list[j + 1]
            nr_list[j + 1] = a

    return (nr_list)"""
