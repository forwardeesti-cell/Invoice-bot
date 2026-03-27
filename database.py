"""
База данных SQLite — хранит объекты, счета, позиции, историю цен
"""

import sqlite3
import json
import logging
from datetime import datetime
from pathlib import Path

logger = logging.getLogger(__name__)

DB_PATH = "invoices.db"


class Database:
    def __init__(self):
        self.conn = sqlite3.connect(DB_PATH, check_same_thread=False)
        self.conn.row_factory = sqlite3.Row
        self._init_tables()

    def _init_tables(self):
        self.conn.executescript("""
        CREATE TABLE IF NOT EXISTS objects (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT UNIQUE NOT NULL,
            xlsx_path TEXT,
            created_at TEXT DEFAULT (datetime('now'))
        );

        CREATE TABLE IF NOT EXISTS invoices (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            object_id INTEGER REFERENCES objects(id),
            number TEXT,
            date TEXT,
            supplier TEXT,
            total REAL,
            section TEXT,
            raw_json TEXT,
            created_at TEXT DEFAULT (datetime('now'))
        );

        CREATE TABLE IF NOT EXISTS invoice_items (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            invoice_id INTEGER REFERENCES invoices(id),
            name TEXT,
            diameter TEXT,
            quantity REAL,
            unit TEXT,
            price REAL,
            amount REAL,
            category TEXT,
            section TEXT
        );

        CREATE TABLE IF NOT EXISTS price_history (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            item_name TEXT NOT NULL,
            diameter TEXT,
            price REAL NOT NULL,
            supplier TEXT,
            object_name TEXT,
            invoice_date TEXT,
            created_at TEXT DEFAULT (datetime('now'))
        );

        CREATE TABLE IF NOT EXISTS categories (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT UNIQUE NOT NULL
        );

        INSERT OR IGNORE INTO categories (name) VALUES
            ('Kanalisatsioon'),('Vesi'),('MÄRG TORU'),('Sadevee'),
            ('Материалы'),('Услуги'),('Аренда'),('Транспорт'),('Прочее');
        """)
        self.conn.commit()

    def get_or_create_object(self, name: str, xlsx_path: str = None) -> int:
        cur = self.conn.execute("SELECT id FROM objects WHERE name=?", (name,))
        row = cur.fetchone()
        if row:
            if xlsx_path:
                self.conn.execute("UPDATE objects SET xlsx_path=? WHERE id=?", (xlsx_path, row['id']))
                self.conn.commit()
            return row['id']
        cur = self.conn.execute(
            "INSERT INTO objects (name, xlsx_path) VALUES (?,?)", (name, xlsx_path)
        )
        self.conn.commit()
        return cur.lastrowid

    def save_invoice(self, invoice_data: dict) -> int:
        object_id = self.get_or_create_object(invoice_data.get('object', 'Без объекта'))
        cur = self.conn.execute(
            """INSERT INTO invoices (object_id, number, date, supplier, total, section, raw_json)
               VALUES (?,?,?,?,?,?,?)""",
            (object_id, invoice_data.get('number'), invoice_data.get('date'),
             invoice_data.get('supplier'),
             sum(i.get('amount', 0) for i in invoice_data.get('items', [])),
             invoice_data.get('section'), json.dumps(invoice_data, ensure_ascii=False))
        )
        invoice_id = cur.lastrowid

        for item in invoice_data.get('items', []):
            self.conn.execute(
                """INSERT INTO invoice_items 
                   (invoice_id, name, diameter, quantity, unit, price, amount, category, section)
                   VALUES (?,?,?,?,?,?,?,?,?)""",
                (invoice_id, item.get('name'), item.get('diameter'), item.get('quantity'),
                 item.get('unit'), item.get('price'), item.get('amount'),
                 item.get('category'), item.get('section'))
            )
            # Сохранить историю цен
            if item.get('price'):
                self.conn.execute(
                    """INSERT INTO price_history (item_name, diameter, price, supplier, object_name, invoice_date)
                       VALUES (?,?,?,?,?,?)""",
                    (item['name'], item.get('diameter'), item['price'],
                     invoice_data.get('supplier'), invoice_data.get('object'),
                     invoice_data.get('date'))
                )
        self.conn.commit()
        return invoice_id

    def check_price_changes(self, invoice_data: dict) -> list:
        """Проверить изменение цен по сравнению с предыдущими закупками"""
        changes = []
        for item in invoice_data.get('items', []):
            if not item.get('price'):
                continue
            cur = self.conn.execute(
                """SELECT price, invoice_date, supplier FROM price_history
                   WHERE item_name=? ORDER BY created_at DESC LIMIT 1""",
                (item['name'],)
            )
            row = cur.fetchone()
            if row and abs(row['price'] - item['price']) > 0.01:
                changes.append({
                    'item_name': item['name'],
                    'old_price': row['price'],
                    'new_price': item['price'],
                    'old_date': row['invoice_date'],
                    'old_supplier': row['supplier']
                })
        return changes

    def get_all_objects(self) -> list:
        cur = self.conn.execute("""
            SELECT o.name, o.xlsx_path,
                   COUNT(DISTINCT i.id) as invoice_count,
                   COALESCE(SUM(i.total), 0) as total
            FROM objects o
            LEFT JOIN invoices i ON i.object_id = o.id
            GROUP BY o.id ORDER BY o.name
        """)
        return [dict(r) for r in cur.fetchall()]

    def get_object_by_name(self, name: str) -> dict:
        cur = self.conn.execute("SELECT * FROM objects WHERE name=?", (name,))
        row = cur.fetchone()
        return dict(row) if row else None

    def get_object_report(self, object_name: str) -> dict:
        obj = self.get_object_by_name(object_name)
        if not obj:
            return None
        cur = self.conn.execute(
            "SELECT COUNT(*) as cnt, COALESCE(SUM(total),0) as total FROM invoices WHERE object_id=?",
            (obj['id'],)
        )
        summary = dict(cur.fetchone())

        cur = self.conn.execute(
            """SELECT category, COALESCE(SUM(amount),0) as amount
               FROM invoice_items ii JOIN invoices i ON ii.invoice_id=i.id
               WHERE i.object_id=? GROUP BY category ORDER BY amount DESC""",
            (obj['id'],)
        )
        categories = {r['category']: r['amount'] for r in cur.fetchall()}

        cur = self.conn.execute(
            """SELECT number, date, total FROM invoices
               WHERE object_id=? ORDER BY created_at DESC LIMIT 5""",
            (obj['id'],)
        )
        recent = [dict(r) for r in cur.fetchall()]

        return {
            'invoice_count': summary['cnt'],
            'total': summary['total'],
            'categories': categories,
            'recent_invoices': recent
        }

    def get_categories(self) -> list:
        cur = self.conn.execute("SELECT name FROM categories ORDER BY name")
        return [r['name'] for r in cur.fetchall()]

    def add_category(self, name: str):
        self.conn.execute("INSERT OR IGNORE INTO categories (name) VALUES (?)", (name,))
        self.conn.commit()

    def get_object_xlsx(self, object_name: str) -> str:
        cur = self.conn.execute("SELECT xlsx_path FROM objects WHERE name=?", (object_name,))
        row = cur.fetchone()
        return row['xlsx_path'] if row else None

    def set_object_xlsx(self, object_name: str, xlsx_path: str):
        self.conn.execute("UPDATE objects SET xlsx_path=? WHERE name=?", (xlsx_path, object_name))
        self.conn.commit()
