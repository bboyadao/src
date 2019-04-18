print(" ")
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
zip_path = os.path.join(
    parent_dir, f"Course – {cp_name} {position} Interview preparation.zip")
send_owl(cp_name, price, p_id, zip_path)
