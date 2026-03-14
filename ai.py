from classes import OrderHeader, OrderLine
import pdfplumber
import pandas as pd
from pathlib import Path
import requests
import json

def parse_pdf_orders_ai(pdf_path, api_key):

    orders = []

    with pdfplumber.open(pdf_path) as pdf:

        for page in pdf.pages:

            text = page.extract_text()

            if not text:
                continue

            data = parse_with_ai(text, api_key)

            order = create_order_from_ai(data)

            orders.append(order)

    return orders

def parse_with_ai(text, api_key):

    url = "https://openrouter.ai/api/v1/chat/completions"

    prompt = f"""
        You are an invoice parser.

        Extract structured data from this document:

        {text}

        Return ONLY JSON with this schema:

        {{
            "document_type": "invoice | refund",
            "id": "string",
            "date": "MM/DD/YYYY",
            "customer": "string",
            "lines": [
                {{
                "activity": "string",
                "description": "string",
                "qty": number,
                "rate": number
                }}
            ]
        }}

        Rules:
            - If document says REFUND, set document_type = refund
            - If document says INVOICE, set document_type = invoice
            - qty and rate must be numbers
            - Do not include explanations
            - Output JSON only
    """

    response = requests.post(
        url,
        headers={
            "Authorization": f"Bearer {api_key}",
            "Content-Type": "application/json"
        },
        json={
            "model": "meta-llama/llama-3.1-8b-instruct",
            "messages": [
                {"role": "system", "content": "You are a helpful assistant that extracts structured data from invoices and refunds."},
                {"role": "system", "content": "Always respond with JSON in the specified format. Do not include any explanations or text outside the JSON."},
                {"role": "user", "content": prompt},
            ],
            "temperature": 0
        }
    )

    result = response.json()

    content = result["choices"][0]["message"]["content"]

    return json.loads(content)

def create_order_from_ai(data):

    order = OrderHeader(
        data["id"],
        data["customer"],
        data["date"]
    )

    for i, line in enumerate(data["lines"], start=1):

        order_line = OrderLine(
            line_id=i,
            order_id=data["id"],
            product_id=line["activity"],
            quantity=line["qty"],
            price=line["rate"],
            activity=line["activity"],
            description=line["description"]
        )

        order.add_line(order_line)

    return order