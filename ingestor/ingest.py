import csv
import json
import os
import time
from datetime import datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Dict, Optional

import boto3
from botocore.exceptions import BotoCoreError, ClientError
from cassandra.cluster import Cluster
from cassandra.query import BatchStatement


DATA_DIR = Path("/data")


def env_int(name: str, default: int) -> int:
    raw = os.getenv(name, str(default)).strip()
    return int(raw)


def env_bool(name: str, default: bool) -> bool:
    raw = os.getenv(name)
    if raw is None:
        return default
    return raw.strip().lower() in {"1", "true", "yes", "y", "on"}


def parse_money(value: Optional[str]) -> Optional[int]:
    if not value:
        return None
    clean = value.replace(",", "").strip()
    if not clean:
        return None
    try:
        return int(clean)
    except ValueError:
        return None


def parse_decimal(value: Optional[str]) -> Optional[Decimal]:
    if value is None:
        return None
    clean = str(value).strip()
    if not clean:
        return None
    try:
        return Decimal(clean)
    except InvalidOperation:
        return None


def parse_invoice_date(value: str) -> Optional[datetime]:
    try:
        return datetime.strptime(value, "%m/%d/%Y %H:%M")
    except ValueError:
        return None


def parse_json_date(value: str) -> Optional[datetime]:
    try:
        return datetime.strptime(value, "%d/%m/%Y, %H:%M:%S")
    except ValueError:
        return None


def cassandra_session():
    host = os.getenv("CASSANDRA_HOST", "cassandra")
    port = int(os.getenv("CASSANDRA_PORT", "9042"))
    keyspace = os.getenv("CASSANDRA_KEYSPACE", "bigdata_project")

    retries = 25
    for i in range(retries):
        try:
            cluster = Cluster([host], port=port)
            session = cluster.connect(keyspace)
            return cluster, session
        except Exception:
            if i == retries - 1:
                raise
            time.sleep(4)
    raise RuntimeError("Could not connect to Cassandra.")


def s3_client():
    region = os.getenv("AWS_REGION", "eu-west-3")
    endpoint_url = os.getenv("S3_ENDPOINT_URL", "").strip() or None
    access_key = os.getenv("AWS_ACCESS_KEY_ID", "").strip() or None
    secret_key = os.getenv("AWS_SECRET_ACCESS_KEY", "").strip() or None
    session_token = os.getenv("AWS_SESSION_TOKEN", "").strip() or None

    kwargs = {"service_name": "s3", "region_name": region}
    if endpoint_url:
        kwargs["endpoint_url"] = endpoint_url
    if access_key and secret_key:
        kwargs["aws_access_key_id"] = access_key
        kwargs["aws_secret_access_key"] = secret_key
    if session_token:
        kwargs["aws_session_token"] = session_token

    return boto3.client(**kwargs)


def ensure_bucket(client, bucket: str, region: str, create_if_missing: bool):
    try:
        client.head_bucket(Bucket=bucket)
        return
    except ClientError as err:
        code = err.response.get("Error", {}).get("Code", "")
        if code not in {"404", "NoSuchBucket", "NotFound"}:
            raise RuntimeError(f"S3 bucket access failed for '{bucket}': {err}") from err

    if not create_if_missing:
        raise RuntimeError(
            f"S3 bucket '{bucket}' not found. Create it first or set S3_CREATE_BUCKET=true."
        )

    try:
        if region == "us-east-1":
            client.create_bucket(Bucket=bucket)
        else:
            client.create_bucket(
                Bucket=bucket,
                CreateBucketConfiguration={"LocationConstraint": region},
            )
    except ClientError as err:
        code = err.response.get("Error", {}).get("Code", "")
        if code in {"BucketAlreadyOwnedByYou", "BucketAlreadyExists"}:
            return
        raise RuntimeError(f"S3 bucket creation failed for '{bucket}': {err}") from err


def upload_file(client, bucket: str, source: Path, object_key: str):
    if not source.exists():
        return

    retries = 5
    for i in range(retries):
        try:
            client.upload_file(str(source), bucket, object_key)
            return
        except (ClientError, BotoCoreError) as err:
            if i == retries - 1:
                raise RuntimeError(f"S3 upload failed for {source}: {err}") from err
            time.sleep(2)


def upload_raw_assets(client, bucket: str):
    upload_file(client, bucket, DATA_DIR / "data.csv", "raw/structured/data.csv")
    upload_file(
        client,
        bucket,
        DATA_DIR / "flipkart_fashion_products_dataset.json",
        "raw/semi_structured/flipkart_fashion_products_dataset.json",
    )
    upload_file(client, bucket, DATA_DIR / "data image" / "fashion.csv", "raw/non_structured/fashion.csv")


def ingest_structured_csv(session, limit: int):
    source = DATA_DIR / "data.csv"
    inserted = 0

    stmt = session.prepare(
        """
        INSERT INTO sales_by_country_month (
          country, year_month, invoice_date, invoice_no, stock_code,
          description, quantity, unit_price, customer_id, total_amount
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """
    )

    with source.open("r", encoding="utf-8", newline="") as f:
        reader = csv.DictReader(f)
        batch = BatchStatement()
        batch_count = 0
        for row in reader:
            if limit >= 0 and inserted >= limit:
                break

            invoice_date = parse_invoice_date(row.get("InvoiceDate", ""))
            if not invoice_date:
                continue

            country = (row.get("Country") or "UNKNOWN").strip() or "UNKNOWN"
            ym = invoice_date.strftime("%Y-%m")

            try:
                quantity = int(row.get("Quantity", "0"))
            except ValueError:
                quantity = 0

            unit_price = parse_decimal(row.get("UnitPrice"))
            if unit_price is None:
                unit_price = Decimal("0")

            total_amount = unit_price * Decimal(quantity)
            customer_id = (row.get("CustomerID") or "UNKNOWN").strip() or "UNKNOWN"

            batch.add(
                stmt,
                (
                    country,
                    ym,
                    invoice_date,
                    row.get("InvoiceNo", ""),
                    row.get("StockCode", ""),
                    row.get("Description", ""),
                    quantity,
                    unit_price,
                    customer_id,
                    total_amount,
                ),
            )
            inserted += 1
            batch_count += 1

            if batch_count >= 100:
                session.execute(batch)
                batch = BatchStatement()
                batch_count = 0

        if batch_count > 0:
            session.execute(batch)

    print(f"[OK] Structured CSV inserted rows: {inserted}")


def ingest_semi_structured_json(session, limit: int):
    source = DATA_DIR / "flipkart_fashion_products_dataset.json"
    inserted = 0

    stmt = session.prepare(
        """
        INSERT INTO products_by_brand_availability (
          brand, out_of_stock, pid, title, selling_price, actual_price,
          average_rating, category, sub_category, crawled_at
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """
    )

    with source.open("r", encoding="utf-8") as f:
        docs = json.load(f)

    batch = BatchStatement()
    batch_count = 0
    for item in docs:
        if limit >= 0 and inserted >= limit:
            break

        brand = (item.get("brand") or "UNKNOWN").strip() or "UNKNOWN"
        out_of_stock = bool(item.get("out_of_stock", False))
        pid = (item.get("pid") or "").strip()
        if not pid:
            continue

        average_rating = parse_decimal(item.get("average_rating"))
        crawled_at = parse_json_date(item.get("crawled_at", ""))

        batch.add(
            stmt,
            (
                brand,
                out_of_stock,
                pid,
                (item.get("title") or "").strip(),
                parse_money(item.get("selling_price")),
                parse_money(item.get("actual_price")),
                average_rating,
                (item.get("category") or "").strip(),
                (item.get("sub_category") or "").strip(),
                crawled_at,
            ),
        )
        inserted += 1
        batch_count += 1

        if batch_count >= 100:
            session.execute(batch)
            batch = BatchStatement()
            batch_count = 0

    if batch_count > 0:
        session.execute(batch)

    print(f"[OK] Semi-structured JSON inserted rows: {inserted}")


def build_image_index(image_root: Path) -> Dict[str, Path]:
    index: Dict[str, Path] = {}
    for base, _, files in os.walk(image_root):
        for name in files:
            if name.lower().endswith(".jpg"):
                full = Path(base) / name
                index[name] = full
    return index


def ingest_non_structured_images(session, client, bucket: str, limit: int):
    meta_csv = DATA_DIR / "data image" / "fashion.csv"
    image_root = DATA_DIR / "data image"
    image_index = build_image_index(image_root)

    stmt = session.prepare(
        """
        INSERT INTO image_metadata_by_category_gender (
          category, gender, product_id, product_type, colour, usage,
          product_title, image_name, image_object_key, image_url
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """
    )

    inserted = 0
    uploaded = 0
    with meta_csv.open("r", encoding="utf-8", newline="") as f:
        reader = csv.DictReader(f)
        batch = BatchStatement()
        batch_count = 0
        for row in reader:
            if limit >= 0 and inserted >= limit:
                break

            image_name = (row.get("Image") or "").strip()
            if not image_name:
                continue

            source_path = image_index.get(image_name)
            object_key = f"raw/non_structured/images/{image_name}"
            if source_path and source_path.exists():
                upload_file(client, bucket, source_path, object_key)
                uploaded += 1

            product_id_raw = (row.get("ProductId") or "").strip()
            try:
                product_id = int(product_id_raw)
            except ValueError:
                continue

            batch.add(
                stmt,
                (
                    (row.get("Category") or "UNKNOWN").strip() or "UNKNOWN",
                    (row.get("Gender") or "UNKNOWN").strip() or "UNKNOWN",
                    product_id,
                    (row.get("ProductType") or "").strip(),
                    (row.get("Colour") or "").strip(),
                    (row.get("Usage") or "").strip(),
                    (row.get("ProductTitle") or "").strip(),
                    image_name,
                    object_key,
                    (row.get("ImageURL") or "").strip(),
                ),
            )

            inserted += 1
            batch_count += 1
            if batch_count >= 100:
                session.execute(batch)
                batch = BatchStatement()
                batch_count = 0

        if batch_count > 0:
            session.execute(batch)

    print(f"[OK] Non-structured image metadata inserted rows: {inserted}")
    print(f"[OK] Non-structured images uploaded to S3: {uploaded}")


def main():
    csv_limit = env_int("CSV_LIMIT", 10000)
    json_limit = env_int("JSON_LIMIT", 5000)
    image_limit = env_int("IMAGE_LIMIT", 500)

    bucket = os.getenv("S3_BUCKET", "").strip()
    if not bucket:
        raise RuntimeError("S3_BUCKET is required.")

    region = os.getenv("AWS_REGION", "eu-west-3").strip()
    create_bucket = env_bool("S3_CREATE_BUCKET", False)

    print(
        f"Starting ingestion with limits: CSV_LIMIT={csv_limit}, JSON_LIMIT={json_limit}, IMAGE_LIMIT={image_limit}"
    )

    cluster, session = cassandra_session()
    client = s3_client()
    ensure_bucket(client, bucket, region, create_bucket)

    upload_raw_assets(client, bucket)
    print("[OK] Raw source files uploaded to S3.")

    ingest_structured_csv(session, csv_limit)
    ingest_semi_structured_json(session, json_limit)
    ingest_non_structured_images(session, client, bucket, image_limit)

    cluster.shutdown()
    print("[DONE] Ingestion pipeline completed.")


if __name__ == "__main__":
    main()
