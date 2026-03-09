from classes import OrderHeader, OrderLine
import pdfplumber
import pandas as pd
from pathlib import Path


def split_cell(cell):
    if not cell:
        return []
    return [x.strip() for x in cell.split("\n") if x.strip()]


def parse_pdf_orders(pdf_path):

    orders = []

    with pdfplumber.open(pdf_path) as pdf:

        for page in pdf.pages:

            order_id = None
            order_date = None
            customer = None

            tables = page.extract_tables()

            order = None

            for table in tables:

                header = table[0]

                # -----------------------------
                # BILL TO TABLE
                # -----------------------------
                if header and "BILL TO" in header[0]:

                    customer_cell = table[1][0]
                    customer_lines = split_cell(customer_cell)

                    if customer_lines:
                        customer = customer_lines[0]

                # -----------------------------
                # INVOICE HEADER TABLE
                # -----------------------------
                if header and "INVOICE #" in header[0]:

                    row = table[1]

                    order_id = row[0]
                    order_date = row[1]

                    order = OrderHeader( order_id, customer, order_date)

                # -----------------------------
                # ITEMS TABLE
                # -----------------------------
                if header and "ACTIVITY" in header[0]:

                    row = table[1]

                    activities = split_cell(row[0])
                    descriptions = split_cell(row[1])
                    qtys = split_cell(row[2])
                    rates = split_cell(row[3])

                    max_len = max(len(activities), len(qtys))

                    for i in range(max_len):

                        product = activities[i] if i < len(activities) else None
                        qty = float(qtys[i]) if i < len(qtys) else 0
                        rate = float(rates[i]) if i < len(rates) else 0

                        line = OrderLine(
                            line_id=i + 1,
                            order_id=order_id,
                            product_id=product,
                            quantity=qty,
                            price=rate,
                            activity=activities[i] if i < len(activities) else "",
                            description=descriptions[i] if i < len(descriptions) else ""
                        )

                        if order:
                            order.add_line(line)

            if order:
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