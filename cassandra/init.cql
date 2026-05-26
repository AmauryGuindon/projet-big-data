CREATE KEYSPACE IF NOT EXISTS bigdata_project
WITH replication = {'class': 'SimpleStrategy', 'replication_factor': 1};

USE bigdata_project;

CREATE TABLE IF NOT EXISTS sales_by_country_month (
  country text,
  year_month text,
  invoice_date timestamp,
  invoice_no text,
  stock_code text,
  description text,
  quantity int,
  unit_price decimal,
  customer_id text,
  total_amount decimal,
  PRIMARY KEY ((country, year_month), invoice_date, invoice_no, stock_code)
) WITH CLUSTERING ORDER BY (invoice_date DESC, invoice_no ASC, stock_code ASC);

CREATE TABLE IF NOT EXISTS products_by_brand_availability (
  brand text,
  out_of_stock boolean,
  pid text,
  title text,
  selling_price int,
  actual_price int,
  average_rating decimal,
  category text,
  sub_category text,
  crawled_at timestamp,
  PRIMARY KEY ((brand, out_of_stock), pid)
);

CREATE TABLE IF NOT EXISTS image_metadata_by_category_gender (
  category text,
  gender text,
  product_id int,
  product_type text,
  colour text,
  usage text,
  product_title text,
  image_name text,
  image_object_key text,
  image_url text,
  PRIMARY KEY ((category, gender), product_id)
);
