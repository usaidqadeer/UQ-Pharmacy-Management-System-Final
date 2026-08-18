from datetime import datetime, timedelta, date
from io import BytesIO
from pathlib import Path
import os

from fastapi import FastAPI, Depends, HTTPException, status
from fastapi.responses import FileResponse, StreamingResponse
from fastapi.security import OAuth2PasswordBearer, OAuth2PasswordRequestForm
from fastapi.staticfiles import StaticFiles
from jose import jwt, JWTError
from passlib.context import CryptContext
from pydantic import BaseModel
from sqlalchemy import create_engine, String, Integer, Float, Date, DateTime, ForeignKey, func
from sqlalchemy.orm import DeclarativeBase, Mapped, mapped_column, sessionmaker, Session
from openpyxl import Workbook
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas

APP_NAME = os.getenv("APP_NAME", "UQ Pharmacy Management")
SECRET = os.getenv("SECRET_KEY", "dev-secret-change-me")
DB_URL = os.getenv("DATABASE_URL", "sqlite:///./pharmacy.db")
engine = create_engine(DB_URL, connect_args={"check_same_thread": False} if DB_URL.startswith("sqlite") else {})
SessionLocal = sessionmaker(bind=engine, autoflush=False)

class Base(DeclarativeBase): pass
class User(Base):
    __tablename__="users"; id:Mapped[int]=mapped_column(primary_key=True); username:Mapped[str]=mapped_column(String(60),unique=True); password:Mapped[str]=mapped_column(String(255)); role:Mapped[str]=mapped_column(String(20),default="owner")
class Supplier(Base):
    __tablename__="suppliers"; id:Mapped[int]=mapped_column(primary_key=True); name:Mapped[str]=mapped_column(String(100)); phone:Mapped[str]=mapped_column(String(30),default=""); company:Mapped[str]=mapped_column(String(100),default="")
class Customer(Base):
    __tablename__="customers"; id:Mapped[int]=mapped_column(primary_key=True); name:Mapped[str]=mapped_column(String(100)); phone:Mapped[str]=mapped_column(String(30),default=""); credit:Mapped[float]=mapped_column(Float,default=0)
class Medicine(Base):
    __tablename__="medicines"; id:Mapped[int]=mapped_column(primary_key=True); name:Mapped[str]=mapped_column(String(120)); generic:Mapped[str]=mapped_column(String(120),default=""); category:Mapped[str]=mapped_column(String(60),default="General"); manufacturer:Mapped[str]=mapped_column(String(100),default=""); barcode:Mapped[str]=mapped_column(String(60),unique=True); batch:Mapped[str]=mapped_column(String(60)); expiry:Mapped[date]=mapped_column(Date); quantity:Mapped[int]=mapped_column(Integer,default=0); purchase_price:Mapped[float]=mapped_column(Float); sale_price:Mapped[float]=mapped_column(Float); supplier_id:Mapped[int|None]=mapped_column(ForeignKey("suppliers.id"),nullable=True)
class Sale(Base):
    __tablename__="sales"; id:Mapped[int]=mapped_column(primary_key=True); customer_id:Mapped[int|None]=mapped_column(ForeignKey("customers.id"),nullable=True); total:Mapped[float]=mapped_column(Float); paid:Mapped[float]=mapped_column(Float); created_at:Mapped[datetime]=mapped_column(DateTime,default=datetime.utcnow)
class SaleItem(Base):
    __tablename__="sale_items"; id:Mapped[int]=mapped_column(primary_key=True); sale_id:Mapped[int]=mapped_column(ForeignKey("sales.id")); medicine_id:Mapped[int]=mapped_column(ForeignKey("medicines.id")); name:Mapped[str]=mapped_column(String(120)); qty:Mapped[int]=mapped_column(Integer); price:Mapped[float]=mapped_column(Float); subtotal:Mapped[float]=mapped_column(Float)
class Expense(Base):
    __tablename__="expenses"; id:Mapped[int]=mapped_column(primary_key=True); title:Mapped[str]=mapped_column(String(120)); category:Mapped[str]=mapped_column(String(60)); amount:Mapped[float]=mapped_column(Float); created_at:Mapped[datetime]=mapped_column(DateTime,default=datetime.utcnow)
class Audit(Base):
    __tablename__="audit_logs"; id:Mapped[int]=mapped_column(primary_key=True); action:Mapped[str]=mapped_column(String(200)); username:Mapped[str]=mapped_column(String(60)); created_at:Mapped[datetime]=mapped_column(DateTime,default=datetime.utcnow)
Base.metadata.create_all(engine)

pwd=CryptContext(schemes=["pbkdf2_sha256"],deprecated="auto"); oauth=OAuth2PasswordBearer(tokenUrl="api/login")
def db():
    s=SessionLocal()
    try: yield s
    finally: s.close()
def current(token:str=Depends(oauth),s:Session=Depends(db)):
    try: name=jwt.decode(token,SECRET,algorithms=["HS256"])["sub"]
    except JWTError: raise HTTPException(401,"Invalid or expired login")
    u=s.query(User).filter(User.username==name).first()
    if not u: raise HTTPException(401,"User not found")
    return u
def audit(s,u,msg): s.add(Audit(action=msg,username=u.username))
def out(x): return {c.name:getattr(x,c.name) for c in x.__table__.columns}

class Register(BaseModel): username:str; password:str
class SupplierIn(BaseModel): name:str; phone:str=""; company:str=""
class CustomerIn(BaseModel): name:str; phone:str=""; credit:float=0
class MedicineIn(BaseModel): name:str; generic:str=""; category:str="General"; manufacturer:str=""; barcode:str; batch:str; expiry:date; quantity:int; purchase_price:float; sale_price:float; supplier_id:int|None=None
class CartItem(BaseModel): medicine_id:int; qty:int
class SaleIn(BaseModel): customer_id:int|None=None; paid:float; items:list[CartItem]
class ExpenseIn(BaseModel): title:str; category:str="General"; amount:float

app=FastAPI(title=APP_NAME,version="1.0")
@app.post("/api/register")
def register(x:Register,s:Session=Depends(db)):
    if s.query(User).count()>0: raise HTTPException(403,"Owner already exists")
    u=User(username=x.username,password=pwd.hash(x.password),role="owner"); s.add(u); s.commit(); return {"message":"Owner created"}
@app.post("/api/login")
def login(f:OAuth2PasswordRequestForm=Depends(),s:Session=Depends(db)):
    u=s.query(User).filter(User.username==f.username).first()
    if not u or not pwd.verify(f.password,u.password): raise HTTPException(401,"Wrong username or password")
    token=jwt.encode({"sub":u.username,"exp":datetime.utcnow()+timedelta(hours=8)},SECRET,algorithm="HS256")
    return {"access_token":token,"token_type":"bearer","role":u.role}
@app.get("/api/me")
def me(u=Depends(current)): return {"username":u.username,"role":u.role}

@app.get("/api/dashboard")
def dashboard(s:Session=Depends(db),u=Depends(current)):
    today=date.today(); soon=today+timedelta(days=90)
    meds=s.query(Medicine).all(); sales=s.query(Sale).all(); expenses=s.query(Expense).all()
    return {"medicines":len(meds),"stock":sum(m.quantity for m in meds),"low_stock":sum(m.quantity<=10 for m in meds),"expiring":sum(today<=m.expiry<=soon for m in meds),"expired":sum(m.expiry<today for m in meds),"sales":round(sum(x.total for x in sales),2),"profit":round(sum((i.price-(s.get(Medicine,i.medicine_id).purchase_price if s.get(Medicine,i.medicine_id) else 0))*i.qty for i in s.query(SaleItem).all())-sum(e.amount for e in expenses),2),"recent_sales":[out(x) for x in s.query(Sale).order_by(Sale.id.desc()).limit(5)]}

def crud_routes(path,model,schema):
    @app.get(f"/api/{path}",name=f"list_{path}")
    def listing(s:Session=Depends(db),u=Depends(current)): return [out(x) for x in s.query(model).order_by(model.id.desc()).all()]
    @app.post(f"/api/{path}",name=f"create_{path}")
    def create(x:schema,s:Session=Depends(db),u=Depends(current)):
        obj=model(**x.model_dump()); s.add(obj); audit(s,u,f"Created {path}: {getattr(obj,'name',getattr(obj,'title','record'))}"); s.commit(); s.refresh(obj); return out(obj)
    @app.delete(f"/api/{path}/{{item_id}}",name=f"delete_{path}")
    def delete(item_id:int,s:Session=Depends(db),u=Depends(current)):
        obj=s.get(model,item_id)
        if not obj: raise HTTPException(404,"Not found")
        s.delete(obj); audit(s,u,f"Deleted {path} #{item_id}"); s.commit(); return {"ok":True}
crud_routes("suppliers",Supplier,SupplierIn); crud_routes("customers",Customer,CustomerIn); crud_routes("expenses",Expense,ExpenseIn)

@app.get("/api/medicines")
def medicines(q:str="",s:Session=Depends(db),u=Depends(current)):
    query=s.query(Medicine)
    if q: query=query.filter((Medicine.name.ilike(f"%{q}%"))|(Medicine.barcode.ilike(f"%{q}%"))|(Medicine.generic.ilike(f"%{q}%")))
    return [out(x) for x in query.order_by(Medicine.expiry).all()]
@app.post("/api/medicines")
def add_med(x:MedicineIn,s:Session=Depends(db),u=Depends(current)):
    if s.query(Medicine).filter(Medicine.barcode==x.barcode).first(): raise HTTPException(400,"Barcode already exists")
    m=Medicine(**x.model_dump()); s.add(m); audit(s,u,f"Added medicine {m.name}"); s.commit(); s.refresh(m); return out(m)
@app.put("/api/medicines/{mid}")
def update_med(mid:int,x:MedicineIn,s:Session=Depends(db),u=Depends(current)):
    m=s.get(Medicine,mid)
    if not m: raise HTTPException(404,"Medicine not found")
    for k,v in x.model_dump().items(): setattr(m,k,v)
    audit(s,u,f"Updated medicine {m.name}"); s.commit(); return out(m)
@app.delete("/api/medicines/{mid}")
def del_med(mid:int,s:Session=Depends(db),u=Depends(current)):
    m=s.get(Medicine,mid)
    if not m: raise HTTPException(404,"Medicine not found")
    s.delete(m); audit(s,u,f"Deleted medicine {m.name}"); s.commit(); return {"ok":True}

@app.post("/api/sales")
def create_sale(x:SaleIn,s:Session=Depends(db),u=Depends(current)):
    total=0; rows=[]
    for it in x.items:
        m=s.get(Medicine,it.medicine_id)
        if not m or it.qty<1 or m.quantity<it.qty: raise HTTPException(400,"Insufficient medicine stock")
        if m.expiry<date.today(): raise HTTPException(400,f"{m.name} is expired")
        total+=m.sale_price*it.qty; rows.append((m,it.qty))
    sale=Sale(customer_id=x.customer_id,total=total,paid=x.paid); s.add(sale); s.flush()
    for m,qty in rows: m.quantity-=qty; s.add(SaleItem(sale_id=sale.id,medicine_id=m.id,name=m.name,qty=qty,price=m.sale_price,subtotal=m.sale_price*qty))
    if x.customer_id and x.paid<total:
        c=s.get(Customer,x.customer_id)
        if c: c.credit+=total-x.paid
    audit(s,u,f"Created sale invoice #{sale.id}"); s.commit(); return {"id":sale.id,"total":total}
@app.get("/api/sales")
def sales(s:Session=Depends(db),u=Depends(current)): return [out(x) for x in s.query(Sale).order_by(Sale.id.desc()).all()]
@app.get("/api/sales/{sid}/invoice")
def invoice(sid:int,s:Session=Depends(db),u=Depends(current)):
    sale=s.get(Sale,sid)
    if not sale: raise HTTPException(404,"Invoice not found")
    items=s.query(SaleItem).filter(SaleItem.sale_id==sid).all(); b=BytesIO(); c=canvas.Canvas(b,pagesize=A4); y=800
    c.setFont("Helvetica-Bold",18); c.drawString(50,y,"UQ PHARMACY - SALES INVOICE"); y-=35; c.setFont("Helvetica",10); c.drawString(50,y,f"Invoice #{sale.id}   Date: {sale.created_at:%Y-%m-%d %H:%M}"); y-=30
    for i in items: c.drawString(50,y,f"{i.name}  x{i.qty}"); c.drawRightString(540,y,f"PKR {i.subtotal:.2f}"); y-=22
    c.line(50,y,540,y); y-=25; c.setFont("Helvetica-Bold",12); c.drawRightString(540,y,f"Total: PKR {sale.total:.2f}"); c.save(); b.seek(0)
    return StreamingResponse(b,media_type="application/pdf",headers={"Content-Disposition":f"attachment; filename=invoice-{sid}.pdf"})
@app.get("/api/reports/excel")
def excel(s:Session=Depends(db),u=Depends(current)):
    wb=Workbook(); ws=wb.active; ws.title="Medicines"; ws.append(["ID","Medicine","Generic","Barcode","Batch","Expiry","Qty","Purchase","Sale"])
    for m in s.query(Medicine).all(): ws.append([m.id,m.name,m.generic,m.barcode,m.batch,str(m.expiry),m.quantity,m.purchase_price,m.sale_price])
    ws2=wb.create_sheet("Sales"); ws2.append(["Invoice","Date","Total","Paid"])
    for x in s.query(Sale).all(): ws2.append([x.id,str(x.created_at),x.total,x.paid])
    b=BytesIO(); wb.save(b); b.seek(0); return StreamingResponse(b,media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",headers={"Content-Disposition":"attachment; filename=pharmacy-report.xlsx"})
@app.get("/api/audit")
def logs(s:Session=Depends(db),u=Depends(current)): return [out(x) for x in s.query(Audit).order_by(Audit.id.desc()).limit(100)]

ROOT=Path(__file__).parent; app.mount("/assets",StaticFiles(directory=ROOT/"frontend"/"assets"),name="assets")
@app.get("/{path:path}",include_in_schema=False)
def spa(path:str): return FileResponse(ROOT/"frontend"/"index.html")
