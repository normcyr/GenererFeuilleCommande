# -*- coding: utf-8 -*-

'''
questionBabac.py - module to search the Cycle Babac catalog
author: Norm1 <normand.cyr@gmail.com>
originally built as a module to Sopel
'''

from http import cookiejar
import mechanicalsoup
import requests
import yaml
import re
import sys

def get_query():

    query = sys.argv[1]

    return query

def load_config():

    with open(pathToConfig + 'config.yml') as ymlfile:
        cfg = yaml.load(ymlfile)

    username = cfg['login']['username']
    password = cfg['login']['password']

    return username, password

def login(username, password):

    # Cookie Jar
    cj = cookiejar.CookieJar()
    s = requests.Session()
    s.cookies = cj

    br = mechanicalsoup.StatefulBrowser(soup_config={'features': 'lxml'}, session=s)
    login_url = "http://cyclebabac.com/wp-login.php"

    # Perform the actual login
    br.open(login_url)
    br.select_form('#loginform')
    br['log'] = str(username)
    br['pwd'] = str(password)
    br.submit_selected()

    return br

def search_item(br, query):

    sku_pattern = re.compile('^\d{2}\-*\d{3}$')  # accept 12-345 or 12345, but not 123456 or 1234
    url = "http://cyclebabac.com/"
    print(query)
    if query != None:
        search_url = url + '?s=' + query
    else:
        search_url = url
    search = br.open(search_url)
    soupsearchpage = br.get_current_page()

    if sku_pattern.match(query):
        query_type = 'sku_only'
        query = query[:2] + '-' + query[-3:]
        itemsfound = soupsearchpage.title
    else:
        query_type = 'text'
        itemsfound = soupsearchpage.findAll(attrs={'class': 'itemTitle'})
    print(len(itemsfound))

    return itemsfound, query_type, query, url

def print_results(br, itemsfound, query_type, query, url):

    if len(itemsfound)>0:
        if query_type == 'sku_only':
            sku = query
            result_url = url + '?s=' + sku

            itempage = br.open(result_url)
            soupitempagetext = br.get_current_page()

            shortitemname = itemsfound.text[13:63]
            skushort = str(sku)
            price = soupitempagetext.find('meta', itemprop='price')
            if price != None:
                pricenumber = float(str(price[u'content']))
                val = str('%.2f') % pricenumber
                checkifstock = soupitempagetext.find('input', attrs={'class': 'input-text qty text'})
                if checkifstock == None:
                    isinstock = 'Out of stock'
                else:
                    isinstock = 'In stock'
                print('You searched by product number, so I am returning a single item.')
                print('#Babac | ' + 'Item name'.ljust(40, ' ') + ' | Price    | Availability' )
                print(skushort + ' | ' + shortitemname.ljust(40, ' ') + ' | ' + val.rjust(6) + ' $' + ' | ' + isinstock)
            else:
                print('No product found :(')
        else:
            if 1 <= len(itemsfound) <= 10:
                print('Returning %i items.' % len(itemsfound))
            elif len(itemsfound) > 10:
                print('I found a lot of items. Returning the first 10 items.')
            print('#Babac | ' + 'Item name'.ljust(40, ' ') + ' | Price    | Availability' )
            for itemname in itemsfound:
                shortitemname = itemname.contents[1].string[:50]
                for itemlink in itemname.find_all('a'):

                    itempage = br.open(itemlink.get('href'))
                    soupitempagetext = br.get_current_page()

                    skushort = str(soupitempagetext.find_all('span', attrs={'class': 'sku'}))[34:40]
                    price = soupitempagetext.find('meta', itemprop='price')
                    pricenumber = float(str(price[u'content']))
                    val = str('%.2f') % pricenumber

                    # command to check if product is out of stock
                    checkifstock = soupitempagetext.find('input', attrs={'class': 'input-text qty text'})
                    if checkifstock == None:
                        isinstock = 'Out of stock'
                    else:
                        isinstock = 'In stock'

                print(skushort + ' | ' + shortitemname.ljust(40, ' ') + ' | ' + val.rjust(6) + ' $' + ' | ' + isinstock)
    else:
        print('No product found :(')

def babac():

    '''.babac <item name> - Search Cycle Babac catalog for the item.'''

    query = get_query()

    if query != None:
        terms_searched = query
        query = query.replace(' ', '+')
        print('Searching in the Babac catalog for: %s' % terms_searched)
    else:
        return print('.babac what? Please specify your query. For example ".babac training wheels"')
        exit(0)

    username, password = load_config()
    br = login(username, password)
    print('logged in')
    itemsfound, query_type, query, url = search_item(br, query)
    print_results(br, itemsfound, query_type, query, url)

if __name__ == '__main__':

    pathToConfig = '/home/biciklo/src/GenererFeuilleCommande/'

    babac()
