
from fastapi import FastAPI, HTTPException, UploadFile, File
from fastapi.responses import FileResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel
from typing import Optional
from datetime import datetime, date
from pathlib import Path
import sqlite3, json, io, zipfile, tempfile
from openpyxl import Workbook
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas

BASE_DIR = Path(__file__).resolve().parent.parent
DB_PATH = BASE_DIR / "coffee_sales.db"
FRONTEND_DIR = BASE_DIR / "frontend"

app = FastAPI(title="Coffee Sales")

def db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    return conn

def init_db():
    with db() as conn:
        conn.executescript("""
        CREATE TABLE IF NOT EXISTS drinks (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL UNIQUE,
            price REAL NOT NULL CHECK(price > 0),
            created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
            updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
        );

        CREATE TABLE IF NOT EXISTS cup_sizes (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL UNIQUE,
            created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
            updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
        );

        CREATE TABLE IF NOT EXISTS sales (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            drink_id INTEGER NOT NULL,
            cup_size_id INTEGER NOT NULL,
            quantity INTEGER NOT NULL CHECK(quantity > 0),
            price REAL NOT NULL CHECK(price > 0),
            sold_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
            is_deleted INTEGER NOT NULL DEFAULT 0,
            created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
            updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY(drink_id) REFERENCES drinks(id),
            FOREIGN KEY(cup_size_id) REFERENCES cup_sizes(id)
        );
        """)
init_db()

def rows(cur):
    return [dict(r) for r in cur.fetchall()]

def one(cur):
    r = cur.fetchone()
    return dict(r) if r else None

class DrinkIn(BaseModel):
    name: str
    price: float

class CupIn(BaseModel):
    name: str

class SaleIn(BaseModel):
    drink_id: int
    cup_size_id: int
    quantity: int

def used_drink(conn, drink_id: int) -> bool:
    return conn.execute("SELECT 1 FROM sales WHERE drink_id=? LIMIT 1", (drink_id,)).fetchone() is not None

def used_cup(conn, cup_id: int) -> bool:
    return conn.execute("SELECT 1 FROM sales WHERE cup_size_id=? LIMIT 1", (cup_id,)).fetchone() is not None

@app.get("/api/health")
def health():
    return {"ok": True}

@app.get("/api/drinks")
def get_drinks():
    with db() as conn:
        return rows(conn.execute("SELECT * FROM drinks ORDER BY name"))

@app.post("/api/drinks")
def create_drink(data: DrinkIn):
    if not data.name.strip() or data.price <= 0:
        raise HTTPException(400, "Название и цена обязательны, цена > 0")
    with db() as conn:
        try:
            cur = conn.execute("INSERT INTO drinks(name, price) VALUES(?,?)", (data.name.strip(), data.price))
            return one(conn.execute("SELECT * FROM drinks WHERE id=?", (cur.lastrowid,)))
        except sqlite3.IntegrityError:
            raise HTTPException(400, "Такой напиток уже существует")

@app.put("/api/drinks/{drink_id}")
def update_drink(drink_id: int, data: DrinkIn):
    if not data.name.strip() or data.price <= 0:
        raise HTTPException(400, "Название и цена обязательны, цена > 0")
    with db() as conn:
        old = one(conn.execute("SELECT * FROM drinks WHERE id=?", (drink_id,)))
        if not old:
            raise HTTPException(404, "Напиток не найден")
        if used_drink(conn, drink_id) and float(data.price) != float(old["price"]):
            raise HTTPException(400, "Цена заблокирована: напиток уже использовался в продажах")
        try:
            conn.execute("UPDATE drinks SET name=?, price=?, updated_at=CURRENT_TIMESTAMP WHERE id=?",
                         (data.name.strip(), data.price, drink_id))
            return one(conn.execute("SELECT * FROM drinks WHERE id=?", (drink_id,)))
        except sqlite3.IntegrityError:
            raise HTTPException(400, "Такой напиток уже существует")

@app.delete("/api/drinks/{drink_id}")
def delete_drink(drink_id: int):
    with db() as conn:
        if used_drink(conn, drink_id):
            raise HTTPException(400, "Нельзя удалить: напиток используется в продажах")
        conn.execute("DELETE FROM drinks WHERE id=?", (drink_id,))
        return {"ok": True}

@app.get("/api/cup-sizes")
def get_cups():
    with db() as conn:
        return rows(conn.execute("SELECT * FROM cup_sizes ORDER BY name"))

@app.post("/api/cup-sizes")
def create_cup(data: CupIn):
    if not data.name.strip():
        raise HTTPException(400, "Объём обязателен")
    with db() as conn:
        try:
            cur = conn.execute("INSERT INTO cup_sizes(name) VALUES(?)", (data.name.strip(),))
            return one(conn.execute("SELECT * FROM cup_sizes WHERE id=?", (cur.lastrowid,)))
        except sqlite3.IntegrityError:
            raise HTTPException(400, "Такой объём уже существует")

@app.put("/api/cup-sizes/{cup_id}")
def update_cup(cup_id: int, data: CupIn):
    if not data.name.strip():
        raise HTTPException(400, "Объём обязателен")
    with db() as conn:
        try:
            conn.execute("UPDATE cup_sizes SET name=?, updated_at=CURRENT_TIMESTAMP WHERE id=?", (data.name.strip(), cup_id))
            return one(conn.execute("SELECT * FROM cup_sizes WHERE id=?", (cup_id,)))
        except sqlite3.IntegrityError:
            raise HTTPException(400, "Такой объём уже существует")

@app.delete("/api/cup-sizes/{cup_id}")
def delete_cup(cup_id: int):
    with db() as conn:
        if used_cup(conn, cup_id):
            raise HTTPException(400, "Нельзя удалить: объём используется в продажах")
        conn.execute("DELETE FROM cup_sizes WHERE id=?", (cup_id,))
        return {"ok": True}

@app.get("/api/sales")
def get_sales(date_: Optional[str] = None, include_deleted: int = 0):
    if date_ is None:
        date_ = date.today().isoformat()
    with db() as conn:
        return rows(conn.execute("""
        SELECT s.*, d.name AS drink_name, c.name AS cup_name, (s.quantity*s.price) AS total
        FROM sales s
        JOIN drinks d ON d.id=s.drink_id
        JOIN cup_sizes c ON c.id=s.cup_size_id
        WHERE date(s.sold_at)=date(?) AND (?=1 OR s.is_deleted=0)
        ORDER BY s.sold_at DESC, s.id DESC
        """, (date_, include_deleted)))

@app.post("/api/sales")
def create_sale(data: SaleIn):
    if data.quantity <= 0:
        raise HTTPException(400, "Количество должно быть положительным")
    with db() as conn:
        drink = one(conn.execute("SELECT * FROM drinks WHERE id=?", (data.drink_id,)))
        cup = one(conn.execute("SELECT * FROM cup_sizes WHERE id=?", (data.cup_size_id,)))
        if not drink or not cup:
            raise HTTPException(400, "Напиток и объём обязательны")
        cur = conn.execute(
            "INSERT INTO sales(drink_id, cup_size_id, quantity, price) VALUES(?,?,?,?)",
            (data.drink_id, data.cup_size_id, data.quantity, drink["price"])
        )
        return one(conn.execute("SELECT * FROM sales WHERE id=?", (cur.lastrowid,)))

@app.put("/api/sales/{sale_id}")
def update_sale(sale_id: int, data: SaleIn):
    if data.quantity <= 0:
        raise HTTPException(400, "Количество должно быть положительным")
    with db() as conn:
        drink = one(conn.execute("SELECT * FROM drinks WHERE id=?", (data.drink_id,)))
        cup = one(conn.execute("SELECT * FROM cup_sizes WHERE id=?", (data.cup_size_id,)))
        if not drink or not cup:
            raise HTTPException(400, "Напиток и объём обязательны")
        conn.execute("""
            UPDATE sales SET drink_id=?, cup_size_id=?, quantity=?, price=?, updated_at=CURRENT_TIMESTAMP
            WHERE id=?
        """, (data.drink_id, data.cup_size_id, data.quantity, drink["price"], sale_id))
        return {"ok": True}

@app.delete("/api/sales/{sale_id}")
def soft_delete_sale(sale_id: int):
    with db() as conn:
        conn.execute("UPDATE sales SET is_deleted=1, updated_at=CURRENT_TIMESTAMP WHERE id=?", (sale_id,))
        return {"ok": True}

@app.post("/api/sales/{sale_id}/restore")
def restore_sale(sale_id: int):
    with db() as conn:
        conn.execute("UPDATE sales SET is_deleted=0, updated_at=CURRENT_TIMESTAMP WHERE id=?", (sale_id,))
        return {"ok": True}

@app.delete("/api/sales/{sale_id}/force")
def force_delete_sale(sale_id: int):
    with db() as conn:
        conn.execute("DELETE FROM sales WHERE id=?", (sale_id,))
        return {"ok": True}

@app.get("/api/reports/daily")
def daily_report(date_: Optional[str] = None):
    if date_ is None:
        date_ = date.today().isoformat()
    with db() as conn:
        items = rows(conn.execute("""
        SELECT d.name AS drink_name, c.name AS cup_name, SUM(s.quantity) AS quantity,
               s.price AS price, SUM(s.quantity*s.price) AS total
        FROM sales s
        JOIN drinks d ON d.id=s.drink_id
        JOIN cup_sizes c ON c.id=s.cup_size_id
        WHERE date(s.sold_at)=date(?) AND s.is_deleted=0
        GROUP BY d.name, c.name, s.price
        ORDER BY d.name, c.name
        """, (date_,)))
        totals = one(conn.execute("""
        SELECT COALESCE(SUM(quantity),0) AS cups, COALESCE(SUM(quantity*price),0) AS revenue
        FROM sales WHERE date(sold_at)=date(?) AND is_deleted=0
        """, (date_,)))
        leader = one(conn.execute("""
        SELECT d.name AS drink_name, SUM(s.quantity) AS qty
        FROM sales s JOIN drinks d ON d.id=s.drink_id
        WHERE date(s.sold_at)=date(?) AND s.is_deleted=0
        GROUP BY d.name ORDER BY qty DESC LIMIT 1
        """, (date_,)))
        return {"date": date_, "total_cups": totals["cups"], "revenue": totals["revenue"],
                "leader": leader["drink_name"] if leader else "Нет данных", "items": items}

@app.get("/api/reports/daily/export.xlsx")
def export_xlsx(date_: Optional[str] = None):
    report = daily_report(date_)
    wb = Workbook()
    ws = wb.active
    ws.title = "Daily report"
    ws.append(["Дата", report["date"]])
    ws.append(["Всего чашек", report["total_cups"]])
    ws.append(["Выручка", report["revenue"]])
    ws.append(["Лидер продаж", report["leader"]])
    ws.append([])
    ws.append(["Напиток", "Объём", "Количество", "Цена", "Сумма"])
    for i in report["items"]:
        ws.append([i["drink_name"], i["cup_name"], i["quantity"], i["price"], i["total"]])
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    wb.save(tmp.name)
    return FileResponse(tmp.name, filename=f"daily_report_{report['date']}.xlsx")

@app.get("/api/reports/daily/export.pdf")
def export_pdf(date_: Optional[str] = None):
    report = daily_report(date_)
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
    c = canvas.Canvas(tmp.name, pagesize=A4)
    w, h = A4
    y = h - 50
    c.setFont("Helvetica-Bold", 16)
    c.drawString(50, y, "Coffee Sales - Daily Report")
    y -= 30
    c.setFont("Helvetica", 11)
    for line in [f"Date: {report['date']}", f"Total cups: {report['total_cups']}", f"Revenue: {report['revenue']}", f"Leader: {report['leader']}"]:
        c.drawString(50, y, line); y -= 20
    y -= 10
    c.setFont("Helvetica-Bold", 10)
    c.drawString(50, y, "Drink"); c.drawString(180, y, "Cup"); c.drawString(260, y, "Qty"); c.drawString(320, y, "Price"); c.drawString(390, y, "Total")
    y -= 18
    c.setFont("Helvetica", 10)
    for i in report["items"]:
        if y < 60:
            c.showPage(); y = h - 50
        c.drawString(50, y, str(i["drink_name"])[:20])
        c.drawString(180, y, str(i["cup_name"])[:12])
        c.drawString(260, y, str(i["quantity"]))
        c.drawString(320, y, str(i["price"]))
        c.drawString(390, y, str(i["total"]))
        y -= 16
    c.save()
    return FileResponse(tmp.name, filename=f"daily_report_{report['date']}.pdf")

@app.get("/api/backups/export")
def backup_export():
    with db() as conn:
        data = {
            "created_at": datetime.now().isoformat(),
            "drinks": rows(conn.execute("SELECT name, price FROM drinks")),
            "cup_sizes": rows(conn.execute("SELECT name FROM cup_sizes")),
            "sales": rows(conn.execute("""
                SELECT s.sold_at, d.name AS drink_name, c.name AS cup_name, s.quantity, s.price, s.is_deleted
                FROM sales s
                JOIN drinks d ON d.id=s.drink_id
                JOIN cup_sizes c ON c.id=s.cup_size_id
            """))
        }
    mem = io.BytesIO()
    with zipfile.ZipFile(mem, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("coffee_backup.json", json.dumps(data, ensure_ascii=False, indent=2))
    mem.seek(0)
    return StreamingResponse(mem, media_type="application/zip",
        headers={"Content-Disposition": f"attachment; filename=coffee_backup_{date.today().isoformat()}.zip"})

@app.post("/api/backups/import")
async def backup_import(file: UploadFile = File(...)):
    raw = await file.read()
    with zipfile.ZipFile(io.BytesIO(raw), "r") as z:
        data = json.loads(z.read("coffee_backup.json").decode("utf-8"))
    result = {"drinks_added":0, "cups_added":0, "sales_added":0, "sales_skipped":0}
    with db() as conn:
        for d in data.get("drinks", []):
            exists = one(conn.execute("SELECT * FROM drinks WHERE name=?", (d["name"],)))
            if not exists:
                conn.execute("INSERT INTO drinks(name,price) VALUES(?,?)", (d["name"], d["price"]))
                result["drinks_added"] += 1
        for c in data.get("cup_sizes", []):
            exists = one(conn.execute("SELECT * FROM cup_sizes WHERE name=?", (c["name"],)))
            if not exists:
                conn.execute("INSERT INTO cup_sizes(name) VALUES(?)", (c["name"],))
                result["cups_added"] += 1
        for s in data.get("sales", []):
            drink = one(conn.execute("SELECT id FROM drinks WHERE name=?", (s["drink_name"],)))
            cup = one(conn.execute("SELECT id FROM cup_sizes WHERE name=?", (s["cup_name"],)))
            exists = one(conn.execute("""
                SELECT id FROM sales WHERE sold_at=? AND drink_id=? AND cup_size_id=? AND quantity=? AND price=?
            """, (s["sold_at"], drink["id"], cup["id"], s["quantity"], s["price"])))
            if exists:
                result["sales_skipped"] += 1
            else:
                conn.execute("""
                INSERT INTO sales(drink_id,cup_size_id,quantity,price,sold_at,is_deleted)
                VALUES(?,?,?,?,?,?)
                """, (drink["id"], cup["id"], s["quantity"], s["price"], s["sold_at"], s.get("is_deleted",0)))
                result["sales_added"] += 1
    return result

app.mount("/", StaticFiles(directory=FRONTEND_DIR, html=True), name="frontend")
