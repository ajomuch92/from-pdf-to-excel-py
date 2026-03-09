class OrderHeader:
    def __init__(self, order_id, customer_id, order_date):
        self.order_id = order_id
        self.customer_id = customer_id
        self.order_date = order_date
        self.lines = []  # List to hold associated OrderLine objects

    def __str__(self):
        return f"OrderHeader(order_id={self.order_id}, customer_id={self.customer_id}, order_date={self.order_date})"
    
    def add_line(self, line):
        if isinstance(line, OrderLine):
            self.lines.append(line)
        else:
            raise ValueError("line must be an instance of OrderLine")


class OrderLine:
    def __init__(self, line_id, order_id, product_id, quantity, price, activity = "", description=""):
        self.line_id = line_id
        self.order_id = order_id
        self.product_id = product_id
        self.quantity = quantity
        self.price = price
        self.activity = activity
        self.description = description

    def __str__(self):
        return f"OrderLine(line_id={self.line_id}, order_id={self.order_id}, product_id={self.product_id}, quantity={self.quantity}, price={self.price})"