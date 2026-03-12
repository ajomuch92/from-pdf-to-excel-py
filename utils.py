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

            raw_lines = page.extract_text_lines()

            if not raw_lines:
                continue

            lines = [l["text"].strip() for l in raw_lines if l.get("text")]

            order_id = None
            order_date = None
            customer = None
            order = None
            doc_type = "invoice"

            items_section = False
            line_id = 1

            for i, line in enumerate(lines):

                normalized = line.upper().replace(" ", "")

                # -----------------------------
                # DOCUMENT TYPE
                # -----------------------------
                if "REFUNDRECEIPT" in normalized:
                    doc_type = "refund"

                # -----------------------------
                # INVOICE NUMBER
                # -----------------------------
                invoice_match = re.search(r'INVOICE\s+(\d+)', line)
                if invoice_match:
                    order_id = invoice_match.group(1)

                # -----------------------------
                # REFUND NUMBER
                # -----------------------------
                refund_match = re.search(r'REFUND\s+(\d+)', line)
                if refund_match and "REFUND DATE" not in line:
                    order_id = refund_match.group(1)

                # -----------------------------
                # DATE
                # -----------------------------
                date_match = re.search(r'DATE\s+(\d{2}/\d{2}/\d{4})', line)
                if date_match:
                    order_date = date_match.group(1)

                refund_date_match = re.search(r'REFUND DATE\s+(\d{2}/\d{2}/\d{4})', line)
                if refund_date_match:
                    order_date = refund_date_match.group(1)

                # -----------------------------
                # CUSTOMER (INVOICE)
                # -----------------------------
                if "BILLTO" in normalized:
                    if i + 1 < len(lines):
                        next_line = lines[i + 1]

                        customer_match = re.search(r'([A-Z]+\s?[A-Z]*\s?/[0-9\-]+)', next_line)

                        if customer_match:
                            customer = customer_match.group(1)
                            if "/" in customer:
                                customer = customer.split("/")[0].strip()
                        else:
                            customer = next_line.strip()

                # -----------------------------
                # CUSTOMER (REFUND)
                # -----------------------------
                if "REFUNDTO" in normalized:
                    if i + 1 < len(lines):
                        next_line = lines[i + 1]
                        customer_match = re.search(r'([A-Z]+\s?[A-Z]*\s?/[0-9\-]+)', next_line)

                        if customer_match:
                            customer = customer_match.group(1)
                            if "/" in customer:
                                customer = customer.split("/")[0].strip()
                        else:
                            customer = next_line.strip()

                # -----------------------------
                # START ITEMS
                # -----------------------------
                if "ACTIVITY" in line and "QTY" in line:

                    items_section = True

                    if order is None and order_id:
                        order = OrderHeader(order_id, customer, order_date)

                    continue

                # -----------------------------
                # ITEMS
                # -----------------------------
                if items_section:

                    if any(x in line for x in ["SUBTOTAL", "TAX", "TOTAL", "BALANCE"]):
                        items_section = False
                        continue

                    item_match = re.search(
                        r'(.+?)\s+(\d+)\s+([\d.]+)\s+([\d.]+)',
                        line
                    )

                    if item_match and order:

                        activity = item_match.group(1).strip()
                        qty = float(item_match.group(2))
                        rate = float(item_match.group(3))

                        line_obj = OrderLine(
                            line_id=line_id,
                            order_id=order_id,
                            product_id=activity,
                            quantity=qty,
                            price=rate,
                            activity=activity,
                            description=activity
                        )

                        order.add_line(line_obj)

                        line_id += 1

            if order:
                orders.append(order)

    return orders

MAX_ITEMS = 30

def export_orders_to_excel(orders, doc_type):

    rows = []

    for order in orders:

        multiplier = 1 if doc_type.upper() == "INVOICES" else -1

        row = {
            "DATE": order.order_date,
            "INVOICE" if doc_type.upper() == "INVOICES" else "REFUND": order.order_id,
            "BILL TO": order.customer_id,
            # "TYPE": doc_type
        }

        for i in range(MAX_ITEMS):

            if i < len(order.lines):

                item = order.lines[i]

                row[f"QTY {i+1}"] = multiplier * item.quantity
                row[f"ACTIVITY {i+1}"] = item.activity
                row[f"DESCRIPTION {i+1}"] = item.description
                row[f"WHOLE PRICE {i+1}"] = ""
                row[f"RATE {i+1}"] = item.price
                row[f"COMPRA {i+1}"] = ""
                row[f"VENTA {i+1}"] = ""
                row[f"GANANCIA {i+1}"] = ""

            else:
                row[f"QTY {i+1}"] = ""
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