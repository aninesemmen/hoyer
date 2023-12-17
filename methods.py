import time
import requests
import urllib.request
import json
import certifi
import ssl
import pandas as pd

def findPriceFromGtins(gtins):
    try:
        url = "https://frontsystemsapis.frontsystems.no/restapi/V2/api/Product"

        hdr ={
        # Request headers
        'x-api-key': 'rSzTE8K.4PYkiCCRpFS89sfyhD6QFvSKJrvsc8Gi',
        'Cache-Control': 'no-cache',
        'Ocp-Apim-Subscription-Key': 'ba812cd412304d1daa83c558113acacc',
        'Content-Type': 'application/json',
        }

        gtins_json = json.dumps(gtins)
        response = requests.post(url=url, headers=hdr, data=gtins_json)
        productList = response.json()
        return productList
    except Exception as e:
        print(e)


def createExcelFileWithColumnNames(columnNames, fileName):
    dataframe = pd.DataFrame(columns=columnNames)
    dataframe.to_excel(fileName, engine='xlsxwriter')


def findPriceFromGtin(gtin):
    try:
        url = "https://frontsystemsapis.frontsystems.no/restapi/V2/api/Product/gtin/" + gtin

        hdr ={
        # Request headers
        'x-api-key': 'rSzTE8K.4PYkiCCRpFS89sfyhD6QFvSKJrvsc8Gi',
        'Cache-Control': 'no-cache',
        'Ocp-Apim-Subscription-Key': 'ba812cd412304d1daa83c558113acacc',
        }

        gcontext = ssl.create_default_context(cafile=certifi.where())

        req = urllib.request.Request(url, headers=hdr)
        time.sleep(2)

        req.get_method = lambda: 'GET'
        response = urllib.request.urlopen(req, context=gcontext)
        return response
    except Exception as e:
        print(e)


def deleteProductsFromProductId(productids):
    try:
        url = "https://frontsystemsapis.frontsystems.no/restapi/V2/api/products"

        hdr ={
        # Request headers
        'x-api-key': 'rSzTE8K.4PYkiCCRpFS89sfyhD6QFvSKJrvsc8Gi',
        'Cache-Control': 'no-cache',
        'Ocp-Apim-Subscription-Key': 'ba812cd412304d1daa83c558113acacc',
        }

        productids_json = json.dumps(productids)
        response = requests.delete(url=url, headers=hdr, data=productids_json)
        print(response)
       
    except Exception as e:
        print(e)