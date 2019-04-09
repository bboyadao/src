import shopify
import ssl

ssl._create_default_https_context = ssl._create_unverified_context
API_KEY, PASSWORD = "6413e468090c0e3acac67e564db39752", "7bd6e0bffac90790b6a6343d3042e923"
shop_url = "https://%s:%s@kourse-com.myshopify.com/admin" % (
    API_KEY, PASSWORD)

shopify.ShopifyResource.set_site(shop_url)


def get_all_resources(resource, **kwargs):
    resource_count = resource.count(**kwargs)
    resources = []
    if resource_count > 0:
        for page in range(1, ((resource_count - 1) // 250) + 2):
            kwargs.update({"limit": 250, "page": page})
            resources.extend(resource.find(**kwargs))
    return resources


def list_products():

    shop = shopify.Shop.current
    collec = shopify.CollectionListing
    products = get_all_resources(shopify.Product)

    return products


def create_collection(resources, cm_name, logo):
    collect = shopify.Collect(
        {'product_id': product_id, 'collection_id': collection_id})
    collect.save()
    pass


def create_product():
    new_product = shopify.Product()
    new_product.title = "Burton Custom Freestyle 151"
    new_product.product_type = "Snowboard"
    new_product.body_html = "<strong>Good snowboard!</strong>"
    new_product.vendor = "Burton"
    image1 = shopify.Image()
    image1.src = logo

    new_product.images = [image1]
    new_product.save()
    success = new_product.save()


if __name__ == '__main__':
    a = list_products()
    print(a)
