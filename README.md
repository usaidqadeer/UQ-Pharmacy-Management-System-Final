# UQ Pharmacy Management System

A responsive pharmacy POS and inventory application built with FastAPI, SQLAlchemy, SQLite, JWT, HTML/CSS and JavaScript.

## Features

- Owner registration and secure JWT login
- Medicine batches, barcodes, stock and expiry tracking
- POS cart, stock deduction and printable PDF invoices
- Customers with credit balance, suppliers and expenses
- Dashboard KPIs, estimated profit, low-stock and expiry alerts
- Excel inventory/sales report and audit trail
- Responsive dark UI for desktop and mobile

## Run on Windows

```bash
python -m venv venv
venv\Scripts\activate
pip install -r requirements.txt
uvicorn main:app --reload
```

Open http://127.0.0.1:8000 and click **Create owner account** the first time.

## Important

Copy `.env.example` to `.env`, set a strong `SECRET_KEY`, and use PostgreSQL for production. This is an inventory and billing tool; it does not provide medical advice or replace a licensed pharmacist.
