from classes import OrderHeader, OrderLine
import requests
import json
import base64
import io
import pypdfium2 as pdfium

def image_to_base64(img):

    buffer = io.BytesIO()
    img.save(buffer, format="PNG")

    return base64.b64encode(buffer.getvalue()).decode()

def pdf_to_images(pdf_path):

    pdf = pdfium.PdfDocument(pdf_path)

    images = []

    for page in pdf:

        bitmap = page.render(scale=3)
        pil_image = bitmap.to_pil()

        images.append(pil_image)

    return images

def parse_pdf_orders_ai(pdf_path, api_key, model_name):

    images = pdf_to_images(pdf_path)

    orders = []

    for img in images:

        data = parse_invoice_image(img, api_key, model_name)

        order = create_order_from_ai(data)

        orders.append(order)

    return orders

def parse_invoice_image(image, api_key, model_name):

    if not api_key:
        raise Exception("API key is required")
    
    if not model_name:
        raise Exception("Model selection is required")

    image_b64 = image_to_base64(image)

    prompt = """
        Extract structured data from this invoice or refund.

        Return ONLY JSON:

        {
            "document_type": "invoice | refund",
            "id": "string",
            "date": "MM/DD/YYYY",
            "customer": "string",
            "lines": [
                {
                "activity": "string",
                "description": "string",
                "qty": number,
                "rate": number
                }
            ]
        }

        
        Rules:
            - If document says REFUND, set document_type = refund
            - If document says INVOICE, set document_type = invoice
            - qty and rate must be numbers
            - Do not include explanations
            - Output JSON only
    """

    response = requests.post(
        "https://openrouter.ai/api/v1/chat/completions",
        headers={
            "Authorization": f"Bearer {api_key}",
            "Content-Type": "application/json"
        },
        json={
            "model": model_name,
            "messages": [
                {
                    "role": "user",
                    "content": [
                        {"type": "text", "text": prompt},
                        {
                            "type": "image_url",
                            "image_url": {
                                "url": f"data:image/png;base64,{image_b64}"
                            }
                        }
                    ]
                }
            ],
            "temperature": 0
        }
    )

    result = response.json()
    if result.get("error") and result["error"].get("message"):
        raise Exception(result["error"]["message"])
    content = result["choices"][0]["message"]["content"]

    content = content.replace('```json', '').replace('```', '').strip()

    return json.loads(content)

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
            "model": "google/gemma-2-9b-it",
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

    content = content.replace('```json', '').replace('```', '').strip()

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