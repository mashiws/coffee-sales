from fastapi import FastAPI, HTTPException, UploadFile, File
from fastapi.responses import FileResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel
from typing import Optional
from datetime import datetime, date
from pathlib import Path
import os, json, io, zipfile, tempfile
from sqlalchemy import create_engine, text
from openpyxl import Workbook
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas

BASE_DIR = Path(__file__).resolve().parent.parent
FRONTEND_DIR = BASE_DIR / "frontend"
DATABASE_URL = os.getenv("DATABASE_URL")
if DATABASE_URL:
    DATABASE_URL = DATABASE_URL.replace("postgres://", "postgresql://", 1)
else:
    DATABASE_URL = f"sqlite:///{BASE_DIR / 'coffee_sales.db'}"
engine = create_engine(DATABASE_URL, future=True)
app = FastAPI(title="Coffee Sales")

def pg(): return engine.url.get_backend_name().startswith("postgresql")
def rows(conn, sql, p=None): return [dict(r._mapping) for r in conn.execute(text(sql), p or {}).fetchall()]
def one(conn, sql, p=None):
    r=conn.execute(text(sql), p or {}).fetchone(); return dict(r._mapping) if r else None

def init_db():
    with engine.begin() as c:
        if pg():
            c.execute(text("""CREATE TABLE IF NOT EXISTS drinks(id SERIAL PRIMARY KEY,name VARCHAR(100) UNIQUE NOT NULL,price NUMERIC(10,2) NOT NULL CHECK(price>0),created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP);"""))
            c.execute(text("""CREATE TABLE IF NOT EXISTS cup_sizes(id SERIAL PRIMARY KEY,name VARCHAR(50) UNIQUE NOT NULL,created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP);"""))
            c.execute(text("""CREATE TABLE IF NOT EXISTS sales(id SERIAL PRIMARY KEY,drink_id INTEGER NOT NULL REFERENCES drinks(id),cup_size_id INTEGER NOT NULL REFERENCES cup_sizes(id),quantity INTEGER NOT NULL CHECK(quantity>0),price NUMERIC(10,2) NOT NULL CHECK(price>0),sold_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,is_deleted BOOLEAN NOT NULL DEFAULT FALSE,created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP);"""))
        else:
            c.execute(text("""CREATE TABLE IF NOT EXISTS drinks(id INTEGER PRIMARY KEY AUTOINCREMENT,name TEXT UNIQUE NOT NULL,price REAL NOT NULL CHECK(price>0),created_at TEXT DEFAULT CURRENT_TIMESTAMP,updated_at TEXT DEFAULT CURRENT_TIMESTAMP);"""))
            c.execute(text("""CREATE TABLE IF NOT EXISTS cup_sizes(id INTEGER PRIMARY KEY AUTOINCREMENT,name TEXT UNIQUE NOT NULL,created_at TEXT DEFAULT CURRENT_TIMESTAMP,updated_at TEXT DEFAULT CURRENT_TIMESTAMP);"""))
            c.execute(text("""CREATE TABLE IF NOT EXISTS sales(id INTEGER PRIMARY KEY AUTOINCREMENT,drink_id INTEGER NOT NULL,cup_size_id INTEGER NOT NULL,quantity INTEGER NOT NULL CHECK(quantity>0),price REAL NOT NULL CHECK(price>0),sold_at TEXT DEFAULT CURRENT_TIMESTAMP,is_deleted INTEGER NOT NULL DEFAULT 0,created_at TEXT DEFAULT CURRENT_TIMESTAMP,updated_at TEXT DEFAULT CURRENT_TIMESTAMP,FOREIGN KEY(drink_id) REFERENCES drinks(id),FOREIGN KEY(cup_size_id) REFERENCES cup_sizes(id));"""))
init_db()

class DrinkIn(BaseModel): name:str; price:float
class CupIn(BaseModel): name:str
class SaleIn(BaseModel): drink_id:int; cup_size_id:int; quantity:int

def dsql(col): return f"DATE({col})=CAST(:date AS DATE)" if pg() else f"date({col})=date(:date)"
def deleted(v): return bool(v) if pg() else int(bool(v))
def used_drink(c,id): return one(c,"SELECT id FROM sales WHERE drink_id=:id LIMIT 1",{"id":id}) is not None
def used_cup(c,id): return one(c,"SELECT id FROM sales WHERE cup_size_id=:id LIMIT 1",{"id":id}) is not None

@app.get('/api/health')
def health(): return {'ok':True,'database':'postgresql' if pg() else 'sqlite'}

@app.get('/api/drinks')
def get_drinks():
    with engine.begin() as c: return rows(c,"SELECT * FROM drinks ORDER BY name")
@app.post('/api/drinks')
def add_drink(x:DrinkIn):
    if not x.name.strip() or x.price<=0: raise HTTPException(400,'Название и цена обязательны, цена > 0')
    try:
        with engine.begin() as c:
            if pg(): return one(c,"INSERT INTO drinks(name,price) VALUES(:n,:p) RETURNING *",{'n':x.name.strip(),'p':x.price})
            c.execute(text("INSERT INTO drinks(name,price) VALUES(:n,:p)"),{'n':x.name.strip(),'p':x.price}); return one(c,"SELECT * FROM drinks WHERE name=:n",{'n':x.name.strip()})
    except Exception as e:
        if 'unique' in str(e).lower() or 'duplicate' in str(e).lower(): raise HTTPException(400,'Такой напиток уже существует')
        raise
@app.put('/api/drinks/{id}')
def upd_drink(id:int,x:DrinkIn):
    if not x.name.strip() or x.price<=0: raise HTTPException(400,'Название и цена обязательны, цена > 0')
    with engine.begin() as c:
        old=one(c,"SELECT * FROM drinks WHERE id=:id",{'id':id})
        if not old: raise HTTPException(404,'Напиток не найден')
        if used_drink(c,id) and float(x.price)!=float(old['price']): raise HTTPException(400,'Цена заблокирована: напиток уже использовался в продажах')
        try:
            c.execute(text("UPDATE drinks SET name=:n, price=:p, updated_at=CURRENT_TIMESTAMP WHERE id=:id"),{'n':x.name.strip(),'p':x.price,'id':id})
        except Exception as e:
            if 'unique' in str(e).lower() or 'duplicate' in str(e).lower(): raise HTTPException(400,'Такой напиток уже существует')
            raise
        return one(c,"SELECT * FROM drinks WHERE id=:id",{'id':id})
@app.delete('/api/drinks/{id}')
def del_drink(id:int):
    with engine.begin() as c:
        if used_drink(c,id): raise HTTPException(400,'Нельзя удалить: напиток используется в продажах')
        c.execute(text("DELETE FROM drinks WHERE id=:id"),{'id':id}); return {'ok':True}

@app.get('/api/cup-sizes')
def get_cups():
    with engine.begin() as c: return rows(c,"SELECT * FROM cup_sizes ORDER BY name")
@app.post('/api/cup-sizes')
def add_cup(x:CupIn):
    if not x.name.strip(): raise HTTPException(400,'Объём обязателен')
    try:
        with engine.begin() as c:
            if pg(): return one(c,"INSERT INTO cup_sizes(name) VALUES(:n) RETURNING *",{'n':x.name.strip()})
            c.execute(text("INSERT INTO cup_sizes(name) VALUES(:n)"),{'n':x.name.strip()}); return one(c,"SELECT * FROM cup_sizes WHERE name=:n",{'n':x.name.strip()})
    except Exception as e:
        if 'unique' in str(e).lower() or 'duplicate' in str(e).lower(): raise HTTPException(400,'Такой объём уже существует')
        raise
@app.put('/api/cup-sizes/{id}')
def upd_cup(id:int,x:CupIn):
    if not x.name.strip(): raise HTTPException(400,'Объём обязателен')
    with engine.begin() as c:
        try: c.execute(text("UPDATE cup_sizes SET name=:n, updated_at=CURRENT_TIMESTAMP WHERE id=:id"),{'n':x.name.strip(),'id':id})
        except Exception as e:
            if 'unique' in str(e).lower() or 'duplicate' in str(e).lower(): raise HTTPException(400,'Такой объём уже существует')
            raise
        return one(c,"SELECT * FROM cup_sizes WHERE id=:id",{'id':id})
@app.delete('/api/cup-sizes/{id}')
def del_cup(id:int):
    with engine.begin() as c:
        if used_cup(c,id): raise HTTPException(400,'Нельзя удалить: объём используется в продажах')
        c.execute(text("DELETE FROM cup_sizes WHERE id=:id"),{'id':id}); return {'ok':True}

@app.get('/api/sales')
def get_sales(date_:Optional[str]=None, include_deleted:int=0):
    date_=date_ or date.today().isoformat(); cond='TRUE' if include_deleted else ('s.is_deleted=FALSE' if pg() else 's.is_deleted=0')
    with engine.begin() as c:
        return rows(c,f"""SELECT s.*,d.name drink_name,cups.name cup_name,(s.quantity*s.price) total FROM sales s JOIN drinks d ON d.id=s.drink_id JOIN cup_sizes cups ON cups.id=s.cup_size_id WHERE {dsql('s.sold_at')} AND {cond} ORDER BY s.sold_at DESC,s.id DESC""",{'date':date_})
@app.post('/api/sales')
def add_sale(x:SaleIn):
    if x.quantity<=0: raise HTTPException(400,'Количество должно быть положительным')
    with engine.begin() as c:
        dr=one(c,"SELECT * FROM drinks WHERE id=:id",{'id':x.drink_id}); cup=one(c,"SELECT * FROM cup_sizes WHERE id=:id",{'id':x.cup_size_id})
        if not dr or not cup: raise HTTPException(400,'Напиток и объём обязательны')
        if pg(): return one(c,"INSERT INTO sales(drink_id,cup_size_id,quantity,price) VALUES(:d,:cu,:q,:p) RETURNING *",{'d':x.drink_id,'cu':x.cup_size_id,'q':x.quantity,'p':dr['price']})
        c.execute(text("INSERT INTO sales(drink_id,cup_size_id,quantity,price) VALUES(:d,:cu,:q,:p)"),{'d':x.drink_id,'cu':x.cup_size_id,'q':x.quantity,'p':dr['price']}); return {'ok':True}
@app.put('/api/sales/{id}')
def upd_sale(id:int,x:SaleIn):
    if x.quantity<=0: raise HTTPException(400,'Количество должно быть положительным')
    with engine.begin() as c:
        dr=one(c,"SELECT * FROM drinks WHERE id=:id",{'id':x.drink_id}); cup=one(c,"SELECT * FROM cup_sizes WHERE id=:id",{'id':x.cup_size_id})
        if not dr or not cup: raise HTTPException(400,'Напиток и объём обязательны')
        c.execute(text("UPDATE sales SET drink_id=:d,cup_size_id=:cu,quantity=:q,price=:p,updated_at=CURRENT_TIMESTAMP WHERE id=:id"),{'d':x.drink_id,'cu':x.cup_size_id,'q':x.quantity,'p':dr['price'],'id':id}); return {'ok':True}
@app.delete('/api/sales/{id}')
def soft_del(id:int):
    with engine.begin() as c: c.execute(text("UPDATE sales SET is_deleted=:v,updated_at=CURRENT_TIMESTAMP WHERE id=:id"),{'v':deleted(True),'id':id}); return {'ok':True}
@app.post('/api/sales/{id}/restore')
def restore(id:int):
    with engine.begin() as c: c.execute(text("UPDATE sales SET is_deleted=:v,updated_at=CURRENT_TIMESTAMP WHERE id=:id"),{'v':deleted(False),'id':id}); return {'ok':True}
@app.delete('/api/sales/{id}/force')
def force(id:int):
    with engine.begin() as c: c.execute(text("DELETE FROM sales WHERE id=:id"),{'id':id}); return {'ok':True}

@app.get('/api/reports/daily')
def daily_report(date_:Optional[str]=None):
    date_=date_ or date.today().isoformat(); delv=deleted(False)
    with engine.begin() as c:
        items=rows(c,f"""SELECT d.name drink_name,cups.name cup_name,SUM(s.quantity) quantity,s.price price,SUM(s.quantity*s.price) total FROM sales s JOIN drinks d ON d.id=s.drink_id JOIN cup_sizes cups ON cups.id=s.cup_size_id WHERE {dsql('s.sold_at')} AND s.is_deleted=:delv GROUP BY d.name,cups.name,s.price ORDER BY d.name,cups.name""",{'date':date_,'delv':delv})
        totals=one(c,f"SELECT COALESCE(SUM(quantity),0) cups,COALESCE(SUM(quantity*price),0) revenue FROM sales WHERE {dsql('sold_at')} AND is_deleted=:delv",{'date':date_,'delv':delv})
        leader=one(c,f"""SELECT d.name drink_name,SUM(s.quantity) qty FROM sales s JOIN drinks d ON d.id=s.drink_id WHERE {dsql('s.sold_at')} AND s.is_deleted=:delv GROUP BY d.name ORDER BY qty DESC LIMIT 1""",{'date':date_,'delv':delv})
        return {'date':date_,'total_cups':int(totals['cups'] or 0),'revenue':float(totals['revenue'] or 0),'leader':leader['drink_name'] if leader else 'Нет данных','items':items}
@app.get('/api/reports/daily/export.xlsx')
def xlsx(date_:Optional[str]=None):
    r=daily_report(date_); wb=Workbook(); ws=wb.active; ws.title='Daily report'
    for row in [['Дата',r['date']],['Всего чашек',r['total_cups']],['Выручка',r['revenue']],['Лидер продаж',r['leader']],[],['Напиток','Объём','Количество','Цена','Сумма']]: ws.append(row)
    for i in r['items']: ws.append([i['drink_name'],i['cup_name'],i['quantity'],float(i['price']),float(i['total'])])
    tmp=tempfile.NamedTemporaryFile(delete=False,suffix='.xlsx'); wb.save(tmp.name); return FileResponse(tmp.name,filename=f"daily_report_{r['date']}.xlsx")
@app.get('/api/reports/daily/export.pdf')
def pdf(date_:Optional[str]=None):
    r=daily_report(date_); tmp=tempfile.NamedTemporaryFile(delete=False,suffix='.pdf'); c=canvas.Canvas(tmp.name,pagesize=A4); w,h=A4; y=h-50
    c.setFont('Helvetica-Bold',16); c.drawString(50,y,'Coffee Sales - Daily Report'); y-=30; c.setFont('Helvetica',11)
    for line in [f"Date: {r['date']}",f"Total cups: {r['total_cups']}",f"Revenue: {r['revenue']}",f"Leader: {r['leader']}"]: c.drawString(50,y,line); y-=20
    y-=10; c.setFont('Helvetica-Bold',10); c.drawString(50,y,'Drink'); c.drawString(180,y,'Cup'); c.drawString(260,y,'Qty'); c.drawString(320,y,'Price'); c.drawString(390,y,'Total'); y-=18; c.setFont('Helvetica',10)
    for i in r['items']:
        if y<60: c.showPage(); y=h-50
        c.drawString(50,y,str(i['drink_name'])[:20]); c.drawString(180,y,str(i['cup_name'])[:12]); c.drawString(260,y,str(i['quantity'])); c.drawString(320,y,str(i['price'])); c.drawString(390,y,str(i['total'])); y-=16
    c.save(); return FileResponse(tmp.name,filename=f"daily_report_{r['date']}.pdf")

@app.get('/api/backups/export')
def backup_export():
    with engine.begin() as c:
        data={'created_at':datetime.now().isoformat(),'drinks':rows(c,'SELECT name,price FROM drinks'),'cup_sizes':rows(c,'SELECT name FROM cup_sizes'),'sales':rows(c,"""SELECT s.sold_at,d.name drink_name,cups.name cup_name,s.quantity,s.price,s.is_deleted FROM sales s JOIN drinks d ON d.id=s.drink_id JOIN cup_sizes cups ON cups.id=s.cup_size_id""")}
    mem=io.BytesIO(); zipfile.ZipFile(mem,'w',zipfile.ZIP_DEFLATED).writestr('coffee_backup.json',json.dumps(data,ensure_ascii=False,indent=2,default=str)); mem.seek(0)
    return StreamingResponse(mem,media_type='application/zip',headers={'Content-Disposition':f'attachment; filename=coffee_backup_{date.today().isoformat()}.zip'})
@app.post('/api/backups/import')
async def backup_import(file:UploadFile=File(...)):
    raw=await file.read(); data=json.loads(zipfile.ZipFile(io.BytesIO(raw)).read('coffee_backup.json').decode('utf-8')); res={'drinks_added':0,'cups_added':0,'sales_added':0,'sales_skipped':0}
    with engine.begin() as c:
        for d in data.get('drinks',[]):
            if not one(c,'SELECT id FROM drinks WHERE name=:n',{'n':d['name']}): c.execute(text('INSERT INTO drinks(name,price) VALUES(:n,:p)'),{'n':d['name'],'p':d['price']}); res['drinks_added']+=1
        for cup in data.get('cup_sizes',[]):
            if not one(c,'SELECT id FROM cup_sizes WHERE name=:n',{'n':cup['name']}): c.execute(text('INSERT INTO cup_sizes(name) VALUES(:n)'),{'n':cup['name']}); res['cups_added']+=1
        for s in data.get('sales',[]):
            dr=one(c,'SELECT id FROM drinks WHERE name=:n',{'n':s['drink_name']}); cu=one(c,'SELECT id FROM cup_sizes WHERE name=:n',{'n':s['cup_name']})
            if not dr or not cu: continue
            exists=one(c,'SELECT id FROM sales WHERE CAST(sold_at AS TEXT)=:sold AND drink_id=:d AND cup_size_id=:cu AND quantity=:q AND price=:p LIMIT 1',{'sold':str(s['sold_at']),'d':dr['id'],'cu':cu['id'],'q':s['quantity'],'p':s['price']})
            if exists: res['sales_skipped']+=1
            else: c.execute(text('INSERT INTO sales(drink_id,cup_size_id,quantity,price,sold_at,is_deleted) VALUES(:d,:cu,:q,:p,:sold,:delv)'),{'d':dr['id'],'cu':cu['id'],'q':s['quantity'],'p':s['price'],'sold':s['sold_at'],'delv':deleted(s.get('is_deleted',False))}); res['sales_added']+=1
    return res

app.mount('/', StaticFiles(directory=FRONTEND_DIR, html=True), name='frontend')
