---
name: odoo-excel-import
description: Detailed guide on importing product lists from varying Excel sheet formats into Odoo ERP using python xmlrpc external API.
---

# Odoo Excel Import Skill Guide

This skill guides agents in importing, matching, and updating product catalogs in Odoo ERP from different spreadsheet formats.

---

## 1. Odoo XML-RPC Integration Reference

Odoo's external API operates over XML-RPC. It requires the `xmlrpc.client` library in Python.

### Authentication Endpoint
- **URL**: `<odoo_url>/xmlrpc/2/common`
- **Method**: `authenticate(db, username, password, {})` -> returns user `uid` (integer)

```python
import xmlrpc.client
common = xmlrpc.client.ServerProxy(f"{url}/xmlrpc/2/common")
uid = common.authenticate(db, username, password, {})
```

### Database Object Endpoint
- **URL**: `<odoo_url>/xmlrpc/2/object`
- **Method**: `execute_kw(db, uid, password, model, method, args, kwargs)`

```python
models = xmlrpc.client.ServerProxy(f"{url}/xmlrpc/2/object")
```

---

## 2. Core API Workflows

### Search Category ("Goods")
Always search for or create the target category to assign it to products.
```python
# Search for category "Goods"
categ_ids = models.execute_kw(db, uid, password, 'product.category', 'search', [[['name', '=', 'Goods']]])
if not categ_ids:
    # Create category if missing
    categ_id = models.execute_kw(db, uid, password, 'product.category', 'create', [{'name': 'Goods'}])
else:
    categ_id = categ_ids[0]
```

### Search Existing Product
Always check Odoo before creating to prevent duplicate SKUs:
```python
domain = [('default_code', '=', sku)]
fields = ['id', 'name', 'list_price', 'standard_price', 'barcode', 'hs_code', 'description_sale', 'type', 'is_storable', 'weight', 'categ_id', 'default_code']
existing = models.execute_kw(db, uid, password, 'product.template', 'search_read', [domain], {'fields': fields, 'limit': 1})
```

### Create Product
Odoo 18 / SaaS-19.3 product structure requires:
- `type`: `'consu'` (physical goods)
- `is_storable`: `True` (inventory tracking)
- `categ_id`: target category ID (Goods)

```python
new_tmpl_id = models.execute_kw(db, uid, password, 'product.template', 'create', [vals])
```

### Update Product
Only write fields that have actually changed. When comparing:
- **Floats (Prices, Weights)**: Compare using absolute difference > `0.001` (to account for float precision).
- **Many2one relation (categ_id)**: Odoo returns list/tuple `[id, name]`. Extract the ID (`existing_val[0]`) before comparing it with the target ID.
- **Strings/Other**: Normalize `None` and empty strings/`False` to `False` before comparing.

```python
models.execute_kw(db, uid, password, 'product.template', 'write', [[product_tmpl_id], changed_fields])
```

---

## 3. Dynamic Excel Synonym Mappings

Spreadsheets come in different formats (OTC Range lists, Wallbox lists, Amazon Reports). Use dynamic matching to resolve column indices:

```python
synonyms = {
    "sku": ["sku", "item no", "item_no", "article", "product code", "reference"],
    "name": ["title", "name", "product name", "product_name", "artikelname", "artikel-name", "item name", "item_name"],
    "description": ["beschreibung des produkts", "produktbeschreibung", "description", "sales description", "sales_description"],
    "reseller_price": ["reseller price", "reseller_price", "price eur", "price_eur", "price", "wholesale price", "wholesale_price", "cost"],
    "rrp": ["listenpreis (uvp)", "listenpreis mit steuern", "rrp", "rrp eur", "rrp_eur", "rrp price", "sales price", "sales_price", "retail price", "retail_price"],
    "barcode": ["ean", "barcode", "upc", "produkt-id", "produkt_id", "product id", "product_id"],
    "barcode_type": ["art der produkt-id", "art_der_produkt_id", "product id type", "product_id_type", "product-id-type"],
    "hs_code": ["hs code", "hs_code", "zolltarifnummer", "tariff code", "tariff_code"],
    "weight": ["nw", "net weight", "net_weight", "weight", "gewicht des pakets", "artikelanzeige-gewicht"]
}
```

---

## 4. Key Data Protection Rules

1.  **Omit Missing Columns**: If a column was not present in the spreadsheet headers (e.g. no Cost Price column in an Amazon report), **do not** include that key in the Odoo `vals` payload. This prevents overwriting existing Odoo records with default values like `0.0` or `False`.
2.  **Filter Out Non-EAN Barcodes (ASINs)**: If `barcode_type` column is present and lists `ASIN` or anything other than `EAN`/`UPC`/`GTIN`/`ISBN`, **do not** write the value to the `barcode` field in Odoo. This protects Odoo's barcode database from invalid values.
3.  **Bypass Technical Rows**: Skip rows where the resolved SKU is `ABC123` (sample row) or contains technical identifiers like `#` or `::` (e.g. column descriptions in Amazon Listing reports).
4.  **Row Detection**: Auto-detect header row index by scanning the first 10 rows for a cell matching one of the SKU synonyms (e.g. Amazon templates start headers on Row 4, not Row 1).
