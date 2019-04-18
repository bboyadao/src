import xmltodict
import json
import os
import base64
import shopify
import requests

API_KEY = "a9452d93d8b18b78fec035be138daebc"
PASSWORD = "ee2fb8b0b04e52c48e7ee6c61586c176"
SHOP = "taolaadao"
shop_url = f"https://{API_KEY}:{PASSWORD}@kourse-com.myshopify.com/admin"
url = f"https://{API_KEY}:{PASSWORD}@{SHOP}.myshopify.com/admin/api/2019-04/products.json"
#
#
# _encode = base64.b64encode(bytes(url, 'utf-8')).decode('ascii')
#
headers = {"Authorization": f"Basic {_encode}"}
# r = requests.get(url, headers=headers)


# r = requests.post(url, headers=headers, json=product)
# res = r.json()
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

des_path = os.path.join(BASE_DIR, "src",  "Templates", "Shopify",
                        "shopify_copy_course.txt")
SEO_path = os.path.join(BASE_DIR, "src", "Templates", "Shopify",
                        "shopify_SEO_company.txt")


def add_product(cp_name, position, price):
    title = f"{cp_name} {position} Interview Preparation Online Course"
    with open(des_path) as r:
        f = r.read()
    if "[Company]" or "[company]" or "[Position]" or "[position]"in f:
        des1 = f.replace("[company]", cp_name)
        des2 = des1.replace("[position]", position)
        des3 = des2.replace("[Company]", cp_name)
        des4 = des2.replace("[Position]", position)

    des = des4
    with open(des_path) as r:
        f = r.read()
    if "[Company]" or "[company]" or "[Position]" or "[position]"in f:
        seo1 = f.replace("[company]", cp_name)
        seo2 = seo1.replace("[position]", position)
        seo3 = seo2.replace("[Company]", cp_name)
        seo4 = seo3.replace("[Position]", position)
    seo = seo4

    product = {
        "product": {
            "metafields_global_title_tag": seo2,
            "metafields_global_description_tag": seo2,
            "title": title,
            "body_html": des,
            "vendor": "Coursetake",
            "product_type": "Digital",
            "variants": [{"price": price}]
            # "tags": seo2
        }
    }
    r = requests.post(url, headers=headers, json=product)
    return r.json()


def upload_image(product_id, img):
    url = f"https://{API_KEY}:{PASSWORD}@{SHOP}.myshopify.com/admin/api/2019-04/products/{product_id}/images.json"

    data = {
        "image": {
            "src": img,
        }
    }
    r = requests.post(url, json=data, headers=headers)
    return r.json()


def send_owl(p_name, price, p_id, zip_path):
    url = 'https://3387976189ccd7b:c9c0214dc44a8efd567a@upload.sendowl.com/api/v1/products.xml'
    headers = {
        "Content-type": "multipart/form-data",
        "Accept": "application/json"
    }
    files = {
        'product[name]': (None, p_name),
        'product[product_type]': (None, 'digital'),
        'product[price]': (None, price),
        'product[shopify_variant_id]': p_id,
        'product[attachment]': (os.path.basename(zip_path), open(zip_path, 'rb')),
    }
    r = requests.post(url, files=files,)
    if r.status_code != 200:
        return None
    return json.dumps(xmltodict.parse(r.text))


def check_collection(name):
    url = f"https://{API_KEY}:{PASSWORD}@{SHOP}.myshopify.com/admin/api/2019-04/custom_collections.json"
    r = requests.get(url)
    data = r.json()
    for i in data['custom_collections']:
        if i['title'] == name:
            return (True, i['id'])
    return (False, None)


def create_collection(cp_name, logo):

    url = f"https://{API_KEY}:{PASSWORD}@{SHOP}.myshopify.com/admin/api/2019-04/custom_collections.json"
    collect = {
        "custom_collection": {
            "title": cp_name,
            "body": f"Courses and Study Guides to help you ace your upcoming interview at {cp_name}",
            "image": {
                "src": logo,
            },

        }}

    r = requests.post(url, json=collect)
    if r.status_code != 201:
        return False
    return True


def add_product_to_collection(p_id, c_id):
    url = f"https://{API_KEY}:{PASSWORD}@{SHOP}.myshopify.com/admin/api/2019-04/collects.json"
    data = {"collect": {
        "product_id": p_id,
        "collection_id": c_id
    }
    }
    r = requests.post(url, json=data)
    if r.status_code != 201:
        return False
    return True


zip_path = os.path.join(BASE_DIR, "src", "archive_name.zip")

if __name__ == '__main__':

    cp_name = "adaoalive"
    position = "adaoalive"
    price = 10
    logo = "https://www.sendowl.com/assets/merchant/sendowl_icon-d08b799319b20920594282f49150ccee72bf69e6c2b591386b5ea7018a3234e4.png"
    p_id = add_product(cp_name, position, price)
    p_id = p_id['product']['id']
    print(p_id)
    (_, c_id) = check_collection(cp_name)
    if c_id is not None:
        add_product_to_collection(p_id, c_id)
    else:
        create_collection(cp_name, logo)
        add_product_to_collection(p_id, c_id)

    b = upload_image(p_id, logo)
    send_owl(cp_name, price, p_id, zip_path)
