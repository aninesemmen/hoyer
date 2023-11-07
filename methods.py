import requests
import urllib.request
import json
import certifi
import ssl


def findPriceFromGtin(gtin):
    try:
        url = "https://frontsystemsapis.frontsystems.no/webshop/api/Product/gtin/" + gtin

        hdr ={
        # Request headers
        'x-api-key': 'rSzTE8K.4PYkiCCRpFS89sfyhD6QFvSKJrvsc8Gi',
        'Cache-Control': 'no-cache',
        'Ocp-Apim-Subscription-Key': 'ba812cd412304d1daa83c558113acacc',
        }

        gcontext = ssl.create_default_context(cafile=certifi.where())

        req = urllib.request.Request(url, headers=hdr)

        req.get_method = lambda: 'GET'
        response = urllib.request.urlopen(req, context=gcontext)
        return response
    except Exception as e:
        print(e)

