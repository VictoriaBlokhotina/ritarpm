from datetime import datetime
from io import BytesIO
import openpyxl
import pandas as pd
import streamlit as st

# ==================== БИЗНЕС-ЛОГИКА ====================


class Product:

    def __init__(self, pid, name, artist, genre, media, price, stock):
        self.id = pid
        self.name = name
        self.artist = artist
        self.genre = genre
        self.media = media
        self.price = price
        self.stock = stock


class MusicStore:

    def __init__(self):
        self.products = []
        self.sales = []
        self.orders = []
        self.total_revenue = 0.0
        self._init_demo_data()

    def _init_demo_data(self):
        demo_products = [
            (1, "Thriller", "Michael Jackson", "Pop", "CD", 15.99, 10),
            (2, "Back in Black", "AC/DC", "Rock", "Vinyl", 22.50, 5),
            (
                3,
                "Dark Side of the Moon",
                "Pink Floyd",
                "Rock",
                "DVD",
                18.75,
                0,
            ),
            (4, "Lemonade", "Beyoncé", "R&B", "CD", 14.99, 7),
            (5, "Abbey Road", "The Beatles", "Rock", "Vinyl", 25.00, 3),
            (6, "Nevermind", "Nirvana", "Grunge", "CD", 16.50, 8),
            (7, "The Eminem Show", "Eminem", "Hip-Hop", "CD", 13.99, 4),
            (8, "Rumours", "Fleetwood Mac", "Rock", "Vinyl", 27.99, 2),
            (9, "21", "Adele", "Pop", "DVD", 12.50, 6),
            (10, "Hotel California", "Eagles", "Rock", "CD", 19.99, 0),
        ]
        for p in demo_products:
            self.products.append(Product(*p))

    def find_product(self, pid):
        return next((p for p in self.products if p.id == pid), None)

    def sell_product(self, pid, quantity, customer_name=""):
        product = self.find_product(pid)
        if not product:
            return {"success": False, "message": "❌ Товар не найден"}
        if product.stock < quantity:
            return {
                "success": False,
                "message": f"⚠️ Недостаточно товара. В наличии: {product.stock} шт.",
            }

        total = product.price * quantity
        product.stock -= quantity
        self.total_revenue += total
        self.sales.append(
            {
                "product_id": pid,
                "product_name": product.name,
                "artist": product.artist,
                "quantity": quantity,
                "total": total,
                "customer": customer_name or "Без имени",
                "date": datetime.now().strftime("%d.%m.%Y %H:%M"),
            }
        )
        return {
            "success": True,
            "message": "✅ Продажа выполнена!",
            "total": total,
            "remaining_stock": product.stock,
        }

    def create_order(self, pid, quantity, customer_name):
        product = self.find_product(pid)
        if not product:
            return {"success": False, "message": "❌ Товар не найден"}

        for order in self.orders:
            if (
                order["product_id"] == pid
                and order["customer"] == customer_name
                and order["status"] == "⏳ Ожидает поставки"
            ):
                order["quantity"] += quantity
                return {
                    "success": True,
                    "message": f"📦 Заказ обновлён! Теперь {order['quantity']} шт. '{product.name}'",
                }

        self.orders.append(
            {
                "product_id": pid,
                "product_name": product.name,
                "artist": product.artist,
                "quantity": quantity,
                "customer": customer_name,
                "status": "⏳ Ожидает поставки",
                "date": datetime.now().strftime("%d.%m.%Y %H:%M"),
            }
        )
        return {
            "success": True,
            "message": f"📦 Заказ на '{product.name}' оформлен!",
        }

    def complete_order(self, order_index):
        if 0 <= order_index < len(self.orders):
            order = self.orders[order_index]
            product = self.find_product(order["product_id"])
            if product:
                product.stock += order["quantity"]
                self.orders.pop(order_index)
                return {
                    "success": True,
                    "message": f"✅ Заказ на '{order['product_name']}' выполнен!",
                }
        return {
            "success": False,
            "message": "❌ Ошибка при выполнении заказа",
        }

    def get_stock_dataframe(self):
        data = [
            {
                "ID": p.id,
                "Название": p.name,
                "Исполнитель": p.artist,
                "Жанр": p.genre,
                "Носитель": p.media,
                "Цена (руб)": p.price,
                "Остаток (шт)": p.stock,
                "Статус": "В наличии" if p.stock > 0 else "Нет в наличии",
            }
            # Сортировка по ID для наглядности
            for p in sorted(self.products, key=lambda x: x.id)
        ]
        return pd.DataFrame(data)

    def get_sales_dataframe(self):
        return pd.DataFrame(self.sales) if self.sales else pd.DataFrame()

    def get_orders_dataframe(self):
        return pd.DataFrame(self.orders) if self.orders else pd.DataFrame()

    def get_revenue(self):
        return self.total_revenue


# ==================== ЭКСПОРТ В EXCEL ====================


def export_to_excel(store):
    """Экспорт всех данных в Excel-файл с тремя листами"""
    output = BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        # Лист 1: Склад
        store.get_stock_dataframe().to_excel(
            writer, sheet_name="Склад", index=False
        )

        # Лист 2: Заказы
        orders_df = store.get_orders_dataframe()
        if not orders_df.empty:
            orders_df.to_excel(writer, sheet_name="Заказы", index=False)
        else:
            pd.DataFrame({"Сообщение": ["Нет оформленных заказов"]}).to_excel(
                writer, sheet_name="Заказы", index=False
            )

        # Лист 3: Отчёт о продажах (Красивое форматирование через pandas)
        sales_df = store.get_sales_dataframe()
        workbook = writer.book
        sheet = workbook.create_sheet("Отчёт о продажах", 0)

        # Стилизация заголовков
        sheet["A1"] = "ОТЧЁТ О ПРОДАЖАХ МУЗЫКАЛЬНОГО МАГАЗИНА"
        sheet["A1"].font = openpyxl.styles.Font(bold=True, size=14)
        sheet["A2"] = (
            f"Дата формирования: {datetime.now().strftime('%d.%m.%Y %H:%M')}"
        )
        sheet["A3"] = f"Общая выручка: {store.get_revenue():.2f} руб."
        sheet["A3"].font = openpyxl.styles.Font(
            bold=True, size=12, color="27ae60"
        )

        # Выгрузка самих продаж ниже шапки
        if not sales_df.empty:
            # Переименовываем для красивого отображения в Excel
            sales_export = sales_df[
                [
                    "date",
                    "product_name",
                    "artist",
                    "quantity",
                    "total",
                    "customer",
                ]
            ].copy()
            sales_export.columns = [
                "Дата",
                "Товар",
                "Исполнитель",
                "Кол-во",
                "Сумма (руб)",
                "Клиент",
            ]
            sales_export.to_excel(
                writer, sheet_name="Отчёт о продажах", startrow=4, index=False
            )
        else: