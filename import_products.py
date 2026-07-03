#!/usr/bin/env python3
import os
import sys
import argparse
import xmlrpc.client
import openpyxl
from dotenv import load_dotenv

# Load Odoo API environment credentials
load_dotenv()

def main():
    # Parse CLI arguments
    parser = argparse.ArgumentParser(description="Import product data into Odoo from Excel.")
    parser.add_argument(
        "--file", 
        default="/Users/harald2018/Library/CloudStorage/GoogleDrive-harald.pilch@e-mobility-brands.com/Meine Ablage/Blaupunkt EV/Preislisten/2024_reseller_OTC_Range_RRP.xlsx", 
        help="Path to the Excel file."
    )
    parser.add_argument(
        "--dry-run", 
        action="store_true", 
        help="Perform a dry run without making actual database changes."
    )
    args = parser.parse_args()

    # Load and validate environment variables
    url = os.getenv("ODOO_URL")
    db = os.getenv("ODOO_DB")
    username = os.getenv("ODOO_USERNAME")
    password = os.getenv("ODOO_PASSWORD")

    if not all([url, db, username, password]):
        print("Error: Missing Odoo configuration in .env file.")
        print("Please ensure ODOO_URL, ODOO_DB, ODOO_USERNAME, and ODOO_PASSWORD are set.")
        sys.exit(1)

    if not url.startswith("http"):
        print("Error: ODOO_URL must start with http:// or https://")
        sys.exit(1)

    # Establish connection with Odoo XML-RPC
    try:
        print(f"Connecting to Odoo at {url}...")
        common = xmlrpc.client.ServerProxy(f"{url}/xmlrpc/2/common")
        version_info = common.version()
        print(f"Successfully connected! Odoo Server Version: {version_info.get('server_version')}")
    except Exception as e:
        print(f"Connection failed: {e}")
        sys.exit(1)

    # Authenticate user
    try:
        print(f"Authenticating user '{username}' for database '{db}'...")
        uid = common.authenticate(db, username, password, {})
        if not uid:
            print("Authentication failed. Please check database name, username, or API Key/Password.")
            sys.exit(1)
        print(f"Authenticated successfully! User UID: {uid}")
    except Exception as e:
        print(f"Authentication error: {e}")
        sys.exit(1)

    # Set up Odoo models proxy
    models = xmlrpc.client.ServerProxy(f"{url}/xmlrpc/2/object")

    # Find or create product category "Goods" dynamically
    print("Checking product category 'Goods'...")
    try:
        categ_domain = [("name", "=", "Goods")]
        existing_categ = models.execute_kw(
            db, uid, password, "product.category", "search", [categ_domain]
        )
        if existing_categ:
            categ_id = existing_categ[0]
            print(f"Found product category 'Goods' (ID: {categ_id}).")
        else:
            categ_id = models.execute_kw(
                db, uid, password, "product.category", "create", [{"name": "Goods"}]
            )
            print(f"Created product category 'Goods' (ID: {categ_id}).")
    except Exception as e:
        print(f"Warning: Could not resolve product category 'Goods': {e}. Defaulting to Odoo's standard category.")
        categ_id = False

    # Read the Excel sheet
    excel_path = args.file
    print(f"Reading Excel file from {excel_path}...")
    try:
        wb = openpyxl.load_workbook(excel_path, data_only=True)
        sheet = wb.active
        print(f"Opened sheet: {sheet.title}")
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        sys.exit(1)

    # Get rows
    rows = list(sheet.iter_rows(values_only=True))
    if not rows:
        print("Error: Excel sheet is empty.")
        sys.exit(1)

    # Auto-detect headers row by scanning the first 10 rows for a SKU synonym
    header_row_idx = 0
    sku_syns = ["sku", "item no", "item_no", "article", "product code", "reference"]
    for r_idx, row in enumerate(rows[:10]):
        row_vals_lower = [str(cell).lower().strip() if cell is not None else "" for cell in row]
        if any(syn in row_vals_lower for syn in sku_syns):
            header_row_idx = r_idx
            print(f"Auto-detected headers row at index {header_row_idx} (Row {header_row_idx + 1})")
            break

    header_row = rows[header_row_idx]
    headers = [str(h).strip() if h is not None else "" for h in header_row]
    print("Found headers (showing first 15):", headers[:15])

    # Map headers to indices using synonyms
    synonyms = {
        "sku": ["sku", "item no", "item_no", "article", "product code", "reference"],
        "name": ["title", "name", "product name", "product_name", "artikelname", "artikel-name", "item name", "item_name"],
        "description": ["beschreibung des produkts", "produktbeschreibung", "description", "sales description", "sales_description"],
        "reseller_price": ["reseller price", "reseller_price", "price eur", "price_eur", "price", "wholesale price", "wholesale_price", "cost"],
        "rrp": ["listenpreis (uvp)", "listenpreis mit steuern", "rrp", "rrp eur", "rrp_eur", "rrp price", "sales price", "sales_price", "retail price", "retail_price"],
        "barcode": ["ean", "barcode", "upc", "produkt-id", "produkt_id", "product id", "product_id"],
        "barcode_type": ["art der produkt-id", "art_der_produkt_id", "product id type", "product_id_type", "product-id-type"],
        "hs_code": ["hs code", "hs_code", "zolltarifnummer", "tariff code", "tariff_code"],
        "masterbox": ["masterbox", "master box", "package qty"],
        "weight": ["nw", "net weight", "net_weight", "weight", "gewicht des pakets", "artikelanzeige-gewicht"]
    }

    indices = {key: -1 for key in synonyms.keys()}
    headers_lower = [str(h).lower().strip() for h in headers]

    for key, syn_list in synonyms.items():
        for syn in syn_list:
            if syn in headers_lower:
                indices[key] = headers_lower.index(syn)
                break

    # Validate required SKU column
    if indices["sku"] == -1:
        print("Error: Could not find SKU / Article column in the Excel headers.")
        print(f"Excel headers (showing first 15): {headers[:15]}")
        sys.exit(1)

    # Keep track of statistics
    created_count = 0
    updated_count = 0
    unchanged_count = 0
    failed_count = 0

    print("\nStarting product synchronization...")
    for row_idx, row in enumerate(rows[header_row_idx + 1:], start=header_row_idx + 2):
        # Skip completely empty rows
        if all(val is None for val in row):
            continue

        item_no = row[indices["sku"]] if indices["sku"] != -1 else None

        # Validate SKU
        if item_no is None:
            print(f"Row {row_idx}: Skipped due to missing SKU / Article.")
            failed_count += 1
            continue

        sku = str(item_no).strip()
        if sku.endswith(".0"):
            sku = sku[:-2]  # Remove trailing float representation

        # Bypass sample row from Amazon templates
        if sku.upper() == "ABC123":
            continue

        # Dynamic name resolution per-row (chooses first non-empty synonym value)
        name_val = None
        name_col_idx = -1
        for syn_key in synonyms["name"]:
            if syn_key in headers_lower:
                col_idx = headers_lower.index(syn_key)
                val = row[col_idx]
                if val is not None and str(val).strip() != "":
                    name_val = str(val).strip()
                    name_col_idx = col_idx
                    break

        if not name_val:
            print(f"Row {row_idx} (SKU {sku}): Skipped due to missing Product Name.")
            failed_count += 1
            continue

        # Dynamic description resolution per-row (excluding the column chosen for name)
        description_val = None
        for syn_key in synonyms["description"]:
            if syn_key in headers_lower:
                col_idx = headers_lower.index(syn_key)
                if col_idx != name_col_idx:
                    val = row[col_idx]
                    if val is not None and str(val).strip() != "":
                        description_val = str(val).strip()
                        break

        reseller_price = row[indices["reseller_price"]] if indices["reseller_price"] != -1 else None
        rrp = row[indices["rrp"]] if indices["rrp"] != -1 else None
        ean = row[indices["barcode"]] if indices["barcode"] != -1 else None
        hs_code = row[indices["hs_code"]] if indices["hs_code"] != -1 else None
        masterbox = row[indices["masterbox"]] if indices["masterbox"] != -1 else None
        weight_val = row[indices["weight"]] if indices["weight"] != -1 else None

        # Parse prices safely (only if not empty/None to prevent overwriting with 0.0)
        cost_price = None
        if indices["reseller_price"] != -1:
            if reseller_price is not None and str(reseller_price).strip() not in ("", "None", "False"):
                try:
                    cost_price = float(reseller_price)
                except ValueError:
                    print(f"Row {row_idx} (SKU {sku}): Invalid Cost Price '{reseller_price}', ignoring cost update.")

        sales_price = None
        if indices["rrp"] != -1:
            if rrp is not None and str(rrp).strip() not in ("", "None", "False"):
                try:
                    sales_price = float(rrp)
                except ValueError:
                    print(f"Row {row_idx} (SKU {sku}): Invalid Sales Price '{rrp}', ignoring price update.")

        # Parse weight safely (only if not empty/None to prevent overwriting with 0.0)
        weight = None
        if indices["weight"] != -1:
            if weight_val is not None and str(weight_val).strip() not in ("", "None", "False"):
                try:
                    weight = float(weight_val)
                except ValueError:
                    print(f"Row {row_idx} (SKU {sku}): Invalid Weight '{weight_val}', ignoring weight update.")

        # Parse EAN / Barcode safely with type filtering (ignores ASINs)
        barcode = None
        if indices["barcode"] != -1:
            raw_barcode = row[indices["barcode"]]
            if raw_barcode is not None:
                barcode_str = str(raw_barcode).strip()
                if barcode_str.endswith(".0"):
                    barcode_str = barcode_str[:-2]
                
                is_valid_type = True
                if indices["barcode_type"] != -1:
                    raw_type = row[indices["barcode_type"]]
                    if raw_type is not None:
                        b_type = str(raw_type).strip().upper()
                        # Only accept standard barcode ID types (EAN, UPC, GTIN, ISBN)
                        if b_type not in ["EAN", "UPC", "GTIN", "ISBN"]:
                            is_valid_type = False
                
                if is_valid_type and barcode_str not in ("", "None", "False"):
                    barcode = barcode_str

        # Parse HS Code safely
        hs_code_str = None
        if indices["hs_code"] != -1:
            if hs_code is not None:
                h_str = str(hs_code).strip()
                if h_str.endswith(".0"):
                    h_str = h_str[:-2]
                if h_str not in ("", "None", "False"):
                    hs_code_str = h_str

        try:
            # Check Odoo for existing product by default_code (SKU)
            domain = [("default_code", "=", sku)]
            fields = ["id", "name", "list_price", "standard_price", "barcode", "hs_code", "description_sale", "type", "is_storable", "weight", "categ_id", "default_code"]
            
            existing = models.execute_kw(
                db, uid, password, "product.template", "search_read", 
                [domain], {"fields": fields, "limit": 1}
            )

            # Target fields mapping (built dynamically based on present columns with values)
            vals = {
                "name": name_val,
                "default_code": sku,
                "type": "consu",
                "is_storable": True,
            }
            if description_val is not None:
                vals["description_sale"] = description_val
            if sales_price is not None:
                vals["list_price"] = sales_price
            if cost_price is not None:
                vals["standard_price"] = cost_price
            if barcode is not None:
                vals["barcode"] = barcode
            if hs_code_str is not None:
                vals["hs_code"] = hs_code_str
            if weight is not None:
                vals["weight"] = weight
            if categ_id:
                vals["categ_id"] = categ_id

            if existing:
                # Update product if it changed
                product_tmpl_id = existing[0]["id"]
                p_data = existing[0]
                
                changed = {}
                for k, val in vals.items():
                    existing_val = p_data.get(k)
                    
                    is_changed = False
                    if k in ["list_price", "standard_price", "weight"]:
                        if abs((existing_val or 0.0) - val) > 0.001:
                            is_changed = True
                    elif k in ["barcode", "hs_code"]:
                        norm_existing = existing_val if existing_val else False
                        norm_new = val if val else False
                        if norm_existing != norm_new:
                            is_changed = True
                    elif k == "categ_id":
                        # Odoo returns many2one relation as [id, name] list/tuple
                        existing_id = existing_val[0] if isinstance(existing_val, (list, tuple)) else existing_val
                        norm_existing = existing_id if existing_id else False
                        norm_new = val if val else False
                        if norm_existing != norm_new:
                            is_changed = True
                    else:
                        if existing_val != val:
                            is_changed = True
                            
                    if is_changed:
                        changed[k] = val

                if changed:
                    if args.dry_run:
                        print(f"[DRY-RUN] Would update product SKU {sku} (ID: {product_tmpl_id}): changes = {changed}")
                    else:
                        models.execute_kw(db, uid, password, "product.template", "write", [[product_tmpl_id], changed])
                        print(f"Updated product SKU {sku} (ID: {product_tmpl_id}): changes = {changed}")
                    updated_count += 1
                else:
                    print(f"Product SKU {sku} is up-to-date. No changes needed.")
                    unchanged_count += 1
                
                # Handle packaging if masterbox is specified
                if masterbox is not None and not args.dry_run:
                    manage_packaging(models, db, uid, password, product_tmpl_id, sku, masterbox)

            else:
                # Create product
                if args.dry_run:
                    print(f"[DRY-RUN] Would create product SKU {sku}: {vals}")
                else:
                    new_tmpl_id = models.execute_kw(db, uid, password, "product.template", "create", [vals])
                    print(f"Created product SKU {sku} (ID: {new_tmpl_id}): {name_val} - Sale Price: {sales_price}, Cost Price: {cost_price}, Weight: {weight}")
                    
                    # Create packaging if masterbox is specified
                    if masterbox is not None:
                        manage_packaging(models, db, uid, password, new_tmpl_id, sku, masterbox)
                
                created_count += 1

        except Exception as e:
            print(f"Row {row_idx} (SKU {sku}): Error processing product: {e}")
            failed_count += 1

    # Print summary
    print("\n--- Import Summary ---")
    status = "Dry-run completed successfully" if args.dry_run else "Import completed"
    print(f"Status: {status}")
    print(f"Products Created: {created_count}")
    print(f"Products Updated: {updated_count}")
    print(f"Products Unchanged: {unchanged_count}")
    print(f"Failed Rows: {failed_count}")

def manage_packaging(models, db, uid, password, product_tmpl_id, sku, masterbox_val):
    """
    Safely manage packaging (Masterbox) in Odoo for the first variant of product_tmpl_id.
    """
    try:
        m_qty = int(float(masterbox_val))
        if m_qty <= 0:
            return
            
        # Get product variant IDs for product.template
        variants = models.execute_kw(
            db, uid, password, "product.product", "search", 
            [[("product_tmpl_id", "=", product_tmpl_id)]]
        )
        if not variants:
            return

        variant_id = variants[0]

        # Check if packaging already exists
        pack_domain = [("product_id", "=", variant_id), ("qty", "=", m_qty)]
        existing_packs = models.execute_kw(db, uid, password, "product.packaging", "search", [pack_domain])
        
        if not existing_packs:
            pack_vals = {
                "name": f"Masterbox ({m_qty} pcs)",
                "qty": m_qty,
                "product_id": variant_id,
            }
            models.execute_kw(db, uid, password, "product.packaging", "create", [pack_vals])
            print(f"  Created packaging for SKU {sku}: Masterbox ({m_qty} pcs)")
    except Exception as e:
        # Wrap packaging in try/except so it doesn't break the main import if the module/permissions are absent
        print(f"  [Warning] Could not manage packaging for SKU {sku}: {e}")

if __name__ == "__main__":
    main()
