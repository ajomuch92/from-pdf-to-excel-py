from classes import OrderHeader, OrderLine
import pdfplumber


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

                    order = OrderHeader(order_id, customer, order_date)

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
                            price=rate
                        )

                        if order:
                            order.add_line(line)

            if order:
                orders.append(order)

    return orders