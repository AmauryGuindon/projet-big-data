"""Génère les diagrammes PNG pour le rapport Big Data."""
from pathlib import Path

import matplotlib.pyplot as plt
import matplotlib.patches as patches
from matplotlib.patches import FancyArrowPatch, FancyBboxPatch

OUT = Path(__file__).parent / "diagrams"
OUT.mkdir(parents=True, exist_ok=True)

# Palette couleurs cohérente (style sobre, contrastes lisibles à l'impression)
COL_SRC = "#FFE4B5"
COL_INGEST = "#FFB347"
COL_CASS = "#4A90E2"
COL_MINIO = "#7AC74F"
COL_DOCKER = "#E8E8E8"
COL_BORDER = "#2C3E50"
COL_TEXT = "#1A1A1A"


def box(ax, x, y, w, h, text, fill, border=COL_BORDER, fontsize=10, weight="bold", text_color=COL_TEXT):
    fbox = FancyBboxPatch(
        (x, y), w, h,
        boxstyle="round,pad=0.02,rounding_size=0.08",
        linewidth=1.5, edgecolor=border, facecolor=fill,
    )
    ax.add_patch(fbox)
    ax.text(x + w / 2, y + h / 2, text, ha="center", va="center",
            fontsize=fontsize, fontweight=weight, color=text_color, wrap=True)


def arrow(ax, x1, y1, x2, y2, label=None, color=COL_BORDER, ls="-"):
    arr = FancyArrowPatch((x1, y1), (x2, y2),
                          arrowstyle="-|>", mutation_scale=18,
                          color=color, linewidth=1.6, linestyle=ls)
    ax.add_patch(arr)
    if label:
        ax.text((x1 + x2) / 2, (y1 + y2) / 2 + 0.12, label,
                ha="center", va="bottom", fontsize=8.5, style="italic", color=color)


# ---------- 1. Architecture globale ----------
fig, ax = plt.subplots(figsize=(11, 6.5))
ax.set_xlim(0, 14)
ax.set_ylim(0, 9)
ax.axis("off")

# Cadre Docker englobant
docker_frame = FancyBboxPatch((0.4, 0.3), 13.2, 7.6,
                              boxstyle="round,pad=0.02,rounding_size=0.1",
                              linewidth=2, edgecolor="#0DB7ED", facecolor="#F5FBFF",
                              linestyle="--")
ax.add_patch(docker_frame)
ax.text(0.6, 0.55, "Docker / docker-compose", ha="left", va="bottom",
        fontsize=11, fontweight="bold", color="#0DB7ED", style="italic")

# Sources
ax.text(2.0, 8.45, "Sources de données brutes", ha="center", fontsize=11, fontweight="bold")
box(ax, 0.7, 6.6, 2.6, 0.8, "CSV — Transactions\ne-commerce", COL_SRC, fontsize=9)
box(ax, 0.7, 5.6, 2.6, 0.8, "JSON — Catalogue\nFlipkart Fashion", COL_SRC, fontsize=9)
box(ax, 0.7, 4.6, 2.6, 0.8, "Images JPG +\nmétadonnées CSV", COL_SRC, fontsize=9)

# Ingestor
box(ax, 4.5, 5.3, 2.6, 1.5, "Service\nd'ingestion\n(Python)", COL_INGEST, fontsize=12)
ax.text(5.8, 4.95, "Parsing • Validation • Batch", ha="center", fontsize=8.5, style="italic")

# Cassandra
box(ax, 8.5, 6.6, 3.2, 1.5, "Apache Cassandra 4.1\nBase distribuée NoSQL",
    COL_CASS, fontsize=11, text_color="white")
ax.text(10.1, 6.4, "Requêtes analytiques (CQL)", ha="center", fontsize=8.5, style="italic")

# MinIO
box(ax, 8.5, 3.6, 3.2, 1.5, "MinIO / AWS S3\nStockage objet — Data Lake",
    COL_MINIO, fontsize=11, text_color="white")
ax.text(10.1, 3.4, "Fichiers bruts (raw zone)", ha="center", fontsize=8.5, style="italic")

# Consommateurs (côté droit)
box(ax, 12.3, 6.9, 1.1, 0.9, "Analyste\nCQL", "#FFFFFF", fontsize=8.5)
box(ax, 12.3, 3.9, 1.1, 0.9, "Archives\nML / BI", "#FFFFFF", fontsize=8.5)

# Volumes persistants
box(ax, 8.5, 1.0, 1.5, 0.9, "Volume\ncassandra_data", "#FFFFFF", fontsize=8, weight="normal")
box(ax, 10.2, 1.0, 1.5, 0.9, "Volume\nminio_data", "#FFFFFF", fontsize=8, weight="normal")

# Flèches
arrow(ax, 3.3, 7.0, 4.5, 6.3)
arrow(ax, 3.3, 6.0, 4.5, 6.05)
arrow(ax, 3.3, 5.0, 4.5, 5.8)
arrow(ax, 7.1, 6.4, 8.5, 7.0, label="INSERT CQL")
arrow(ax, 7.1, 5.7, 8.5, 4.6, label="PUT objet (raw)")
arrow(ax, 11.7, 7.35, 12.3, 7.35)
arrow(ax, 11.7, 4.35, 12.3, 4.35)
arrow(ax, 9.5, 6.6, 9.2, 1.9, color="#888", ls=":")
arrow(ax, 10.5, 3.6, 10.9, 1.9, color="#888", ls=":")

ax.text(7, 0.05, "Figure 1 — Architecture globale du pipeline Big Data",
        ha="center", fontsize=10, style="italic")
plt.tight_layout()
plt.savefig(OUT / "01_architecture.png", dpi=180, bbox_inches="tight", facecolor="white")
plt.close()


# ---------- 2. Flux de données détaillé ----------
fig, ax = plt.subplots(figsize=(11, 5.5))
ax.set_xlim(0, 14)
ax.set_ylim(0, 7)
ax.axis("off")

stages = [
    (0.4, "Lecture\nfichier source", COL_SRC),
    (3.0, "Parsing &\nnettoyage", "#FFD580"),
    (5.6, "Préparation\nbatch", "#FFD580"),
    (8.2, "Insertion\nCassandra", COL_CASS),
    (10.8, "Upload S3\n(MinIO)", COL_MINIO),
]
for x, label, fill in stages:
    tc = "white" if fill in (COL_CASS,) or fill == COL_MINIO else COL_TEXT
    box(ax, x, 3.0, 2.4, 1.6, label, fill, fontsize=10.5, text_color=tc)

for i in range(len(stages) - 1):
    x1 = stages[i][0] + 2.4
    x2 = stages[i + 1][0]
    arrow(ax, x1, 3.8, x2, 3.8)

# Légendes sous chaque étape
notes = [
    (0.4 + 1.2, "CSV / JSON / JPG"),
    (3.0 + 1.2, "Conversion types\nGestion valeurs nulles"),
    (5.6 + 1.2, "BatchStatement\n(100 lignes)"),
    (8.2 + 1.2, "3 tables\norientées requêtes"),
    (10.8 + 1.2, "raw/structured\nraw/semi_structured\nraw/non_structured"),
]
for x, txt in notes:
    ax.text(x, 2.4, txt, ha="center", va="top", fontsize=8.5, color="#444")

ax.text(7, 5.6, "Pipeline d'ingestion (one-shot job)", ha="center",
        fontsize=12, fontweight="bold")
ax.text(7, 0.1, "Figure 2 — Flux de données du service d'ingestion",
        ha="center", fontsize=10, style="italic")
plt.tight_layout()
plt.savefig(OUT / "02_flux_donnees.png", dpi=180, bbox_inches="tight", facecolor="white")
plt.close()


# ---------- 3. Modèle Cassandra ----------
fig, ax = plt.subplots(figsize=(11, 7))
ax.set_xlim(0, 14)
ax.set_ylim(0, 10)
ax.axis("off")

ax.text(7, 9.6, "Keyspace : bigdata_project   •   replication = SimpleStrategy{1}",
        ha="center", fontsize=11, fontweight="bold", color=COL_BORDER)


def table_box(ax, x, y, title, fields, fill):
    w, h = 4.0, 5.5
    fbox = FancyBboxPatch((x, y), w, h,
                          boxstyle="round,pad=0.02,rounding_size=0.08",
                          linewidth=1.5, edgecolor=COL_BORDER, facecolor=fill)
    ax.add_patch(fbox)
    ax.text(x + w / 2, y + h - 0.35, title, ha="center", va="center",
            fontsize=10.5, fontweight="bold", color="white")
    # Lignes
    line_y = y + h - 0.85
    for label, val, style_color in fields:
        ax.text(x + 0.25, line_y, label, ha="left", va="center",
                fontsize=8.5, color=style_color, fontweight="bold")
        ax.text(x + w - 0.25, line_y, val, ha="right", va="center",
                fontsize=8.5, color="#1A1A1A")
        line_y -= 0.38


# Table 1 — sales_by_country_month
fields_1 = [
    ("PK partition", "(country, year_month)", "#FFD93D"),
    ("CK clustering", "invoice_date DESC", "#FFD93D"),
    ("CK clustering", "invoice_no, stock_code", "#FFD93D"),
    ("col", "description : text", "#FFFFFF"),
    ("col", "quantity : int", "#FFFFFF"),
    ("col", "unit_price : decimal", "#FFFFFF"),
    ("col", "customer_id : text", "#FFFFFF"),
    ("col", "total_amount : decimal", "#FFFFFF"),
    ("Requête cible", "ventes d'un pays\nsur un mois donné", "#FFFFFF"),
]
table_box(ax, 0.4, 2.5, "sales_by_country_month", fields_1, COL_CASS)

# Table 2 — products_by_brand_availability
fields_2 = [
    ("PK partition", "(brand, out_of_stock)", "#FFD93D"),
    ("CK clustering", "pid", "#FFD93D"),
    ("col", "title : text", "#FFFFFF"),
    ("col", "selling_price : int", "#FFFFFF"),
    ("col", "actual_price : int", "#FFFFFF"),
    ("col", "average_rating : decimal", "#FFFFFF"),
    ("col", "category : text", "#FFFFFF"),
    ("col", "sub_category : text", "#FFFFFF"),
    ("Requête cible", "produits (in)disponibles\nd'une marque", "#FFFFFF"),
]
table_box(ax, 5.0, 2.5, "products_by_brand_availability", fields_2, COL_CASS)

# Table 3 — image_metadata_by_category_gender
fields_3 = [
    ("PK partition", "(category, gender)", "#FFD93D"),
    ("CK clustering", "product_id", "#FFD93D"),
    ("col", "product_type : text", "#FFFFFF"),
    ("col", "colour : text", "#FFFFFF"),
    ("col", "usage : text", "#FFFFFF"),
    ("col", "product_title : text", "#FFFFFF"),
    ("col", "image_name : text", "#FFFFFF"),
    ("col", "image_object_key → S3", "#FFFFFF"),
    ("Requête cible", "images d'une catégorie\npour un genre", "#FFFFFF"),
]
table_box(ax, 9.6, 2.5, "image_metadata_by_category_gender", fields_3, COL_CASS)

# Légende
ax.text(0.4, 1.6, "Légende :", fontsize=9.5, fontweight="bold")
ax.add_patch(patches.Rectangle((1.6, 1.5), 0.3, 0.2, facecolor="#FFD93D", edgecolor="#000"))
ax.text(2.0, 1.6, "Clé de partition / clustering", fontsize=9, va="center")
ax.add_patch(patches.Rectangle((6.4, 1.5), 0.3, 0.2, facecolor="#FFFFFF", edgecolor="#000"))
ax.text(6.8, 1.6, "Colonne classique", fontsize=9, va="center")
ax.text(7, 0.2, "Figure 3 — Modèle physique Cassandra orienté requêtes",
        ha="center", fontsize=10, style="italic")
plt.tight_layout()
plt.savefig(OUT / "03_modele_cassandra.png", dpi=180, bbox_inches="tight", facecolor="white")
plt.close()


# ---------- 4. CAP Theorem ----------
fig, ax = plt.subplots(figsize=(8, 7.5))
ax.set_xlim(-1.2, 1.2)
ax.set_ylim(-1.0, 1.3)
ax.axis("off")

# Triangle CAP
import numpy as np
top = (0, 1.0)
bl = (-0.95, -0.55)
br = (0.95, -0.55)
tri = patches.Polygon([top, bl, br], closed=True,
                     facecolor="#F0F4F8", edgecolor=COL_BORDER, linewidth=2)
ax.add_patch(tri)

# Sommets
ax.text(top[0], top[1] + 0.10, "C — Consistency", ha="center",
        fontsize=13, fontweight="bold", color="#C0392B")
ax.text(bl[0] - 0.05, bl[1] - 0.1, "A — Availability", ha="right",
        fontsize=13, fontweight="bold", color="#16A085")
ax.text(br[0] + 0.05, br[1] - 0.1, "P — Partition tolerance", ha="left",
        fontsize=13, fontweight="bold", color="#2980B9")

# Position Cassandra (entre A et P, plus proche de A)
cass = (0.1, -0.40)
ax.scatter(*cass, s=320, color="#4A90E2", edgecolor="white", linewidth=2, zorder=5)
ax.text(cass[0], cass[1] - 0.13, "Cassandra\n(AP)", ha="center", fontsize=10.5,
        fontweight="bold", color="#1A1A1A")

# Annotations latérales
ax.text(-1.05, 0.6, "RDBMS\n(CA)", ha="left", fontsize=9, color="#888", style="italic")
ax.text(0.6, 0.65, "MongoDB,\nHBase (CP)", ha="left", fontsize=9, color="#888", style="italic")

ax.text(0, -1.0, "Figure 4 — Positionnement de Cassandra dans le théorème CAP",
        ha="center", fontsize=10, style="italic")
plt.tight_layout()
plt.savefig(OUT / "04_cap_theorem.png", dpi=180, bbox_inches="tight", facecolor="white")
plt.close()

print("Diagrammes générés dans", OUT)
