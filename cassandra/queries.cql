USE bigdata_project;

-- 1) Ventes par pays et mois (example partition)
SELECT invoice_date, invoice_no, stock_code, quantity, unit_price, total_amount
FROM sales_by_country_month
WHERE country = 'United Kingdom' AND year_month = '2010-12'
LIMIT 20;

-- 2) Produits indisponibles d'une marque
SELECT pid, title, selling_price, average_rating, category, sub_category
FROM products_by_brand_availability
WHERE brand = 'York' AND out_of_stock = true
LIMIT 20;

-- 3) Métadonnées des images par catégorie + genre
SELECT product_id, product_title, image_name, image_object_key, image_url
FROM image_metadata_by_category_gender
WHERE category = 'Apparel' AND gender = 'Girls'
LIMIT 20;
