from classes import OrderHeader, OrderLine
import pdfplumber
import pandas as pd
from pathlib import Path
import re


def split_cell(cell):
    if not cell:
        return []
    return [x.strip() for x in cell.split("\n") if x.strip()]


def parse_pdf_orders(pdf_path):

    orders = []

    with pdfplumber.open(pdf_path) as pdf:

        for page in pdf.pages:

            text = page.extract_text()

            if not text:
                continue

            # -----------------------------
            # INVOICE NUMBER
            # -----------------------------
            invoice_match = re.search(r'INVOICE\s+(\d+)', text)
            order_id = invoice_match.group(1) if invoice_match else None

            # -----------------------------
            # DATE
            # -----------------------------
            date_match = re.search(r'DATE\s+(\d{2}/\d{2}/\d{4})', text)
            order_date = date_match.group(1) if date_match else None

            # -----------------------------
            # CUSTOMER (BILL TO)
            # -----------------------------
            bill_block = re.search(
                r'BILLTO.*?\n(.*?)DATE',
                text,
                re.S
            )

            if bill_block:
                block = bill_block.group(1)

                customer_match = re.search(r'([A-Z ]+\/\d+)', block)

                if customer_match:
                    customer = customer_match.group(1).strip().split("/")[0].strip()

            order = OrderHeader(order_id, customer, order_date)

            # -----------------------------
            # ITEMS SECTION
            # -----------------------------
            items_section = re.search(
                r'ACTIVITY DESCRIPTION QTY RATE AMOUNT(.*?)SUBTOTAL',
                text,
                re.S
            )

            if items_section:

                items_text = items_section.group(1)

                # pattern de linea de producto
                item_pattern = re.findall(
                    r'([A-Z0-9\s\-]+?)\s+([A-Z0-9\s\-,]+?)\s+(\d+)\s+([\d\.]+)',
                    items_text
                )

                for i, match in enumerate(item_pattern):

                    activity = match[0].strip()
                    description = match[1].strip()
                    qty = float(match[2])
                    rate = float(match[3])

                    line = OrderLine(
                        line_id=i + 1,
                        order_id=order_id,
                        product_id=activity,
                        quantity=qty,
                        price=rate,
                        activity=activity,
                        description=description
                    )

                    order.add_line(line)

            orders.append(order)

    return orders

MAX_ITEMS = 30

def export_orders_to_excel(orders, doc_type):

    rows = []

    for order in orders:

        multiplier = -1 if doc_type.lower() == "refund" else 1

        row = {
            "DATE": order.order_date,
            "INVOICE": order.order_id,
            "BILL TO": order.customer_id,
            "TYPE": doc_type
        }

        for i in range(MAX_ITEMS):

            if i < len(order.lines):

                item = order.lines[i]

                row[f"ACTIVITY {i+1}"] = item.activity
                row[f"DESCRIPTION {i+1}"] = item.description
                row[f"WHOLE PRICE {i+1}"] = ""
                row[f"RATE {i+1}"] = item.price
                row[f"QTY {i+1}"] = multiplier * item.quantity
                row[f"COMPRA {i+1}"] = ""
                row[f"VENTA {i+1}"] = ""
                row[f"GANANCIA {i+1}"] = ""

            else:

                row[f"ACTIVITY {i+1}"] = ""
                row[f"DESCRIPTION {i+1}"] = ""
                row[f"WHOLE PRICE {i+1}"] = ""
                row[f"RATE {i+1}"] = ""
                row[f"COMPRA {i+1}"] = ""
                row[f"VENTA {i+1}"] = ""
                row[f"GANANCIA {i+1}"] = ""

        rows.append(row)

    df = pd.DataFrame(rows)

    downloads = Path.home() / "Downloads"
    file_path = downloads / f"orders_export_{pd.Timestamp.now().strftime('%Y-%m-%d_%H-%M-%S')}.xlsx"

    df.to_excel(file_path, index=False)

    return file_path