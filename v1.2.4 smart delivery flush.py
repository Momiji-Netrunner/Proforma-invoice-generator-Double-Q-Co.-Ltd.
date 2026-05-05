import tkinter as tk
from tkinter import ttk, messagebox
from tkcalendar import DateEntry
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_COLOR_INDEX
from docx.shared import Pt, RGBColor
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.table import _Row
from copy import deepcopy
from datetime import datetime, timedelta
import os
import sys
import re

if getattr(sys, "frozen", False):
    # When packed as an EXE, use the executable's folder for output files,
    # instead of the temporary runtime extraction location.
    BASE_DIR = os.path.dirname(os.path.abspath(sys.argv[0]))
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))


def format_currency(value):
    return f"{value:,.2f}"


def get_currency_prefix(currency=None):
    currency = (currency or "USD").strip().upper()
    return currency if currency in ("USD", "UGX") else "USD"


def smart_linebreak_specs(text):
    keywords = [
        "Front Axle", "Inner dimensions", "Column", "Loading", "Towing",
        "Model:CPCD35-XC5K2", "Capacity 3500kg@500mm load center", 
        "Without", "Overall", "Tires:", "Tire:", "Locks:", "Fifth",
        "2 stage 3meter standard lift mast", "3piece valve with side shift",
        "1070mm fork", "solid tires", "comfortable seat", "other all standard",
        "Steering:", "Gearbox:", "Engine:", "Cabin:", "XINCHAI C490engine",
        "Automatic transimission",
        "Cargo", "Color:", "Load:", "Fuel", "Rear Axle", "With", "Volume of the compartment", 
        "Axle: FUWA",
        "Traction pin", "Brake chamber"
    ]
    
        
    for keyword in keywords:
        # Replace space + keyword with newline + keyword
        text = re.sub(r'(?<=[^\n])' + re.escape(keyword), '\n' + keyword, text)
    
    # Clean up and trim
    text = text.strip()
    return text


def parse_specs_text(spec_text):
    parts = [p.strip() for p in spec_text.replace("\n", "|").split("|") if p.strip()]
    return parts


def set_cell_text(cell, text, bold=False, underline=False, font_name=None, font_size=None, font_color=None, highlight_color=None):
    cell.text = ""
    paragraph = cell.paragraphs[0]
    run = paragraph.add_run(text)
    run.bold = bold
    run.underline = underline
    if font_name is not None:
        run.font.name = font_name
    if font_size is not None:
        run.font.size = Pt(font_size)
    if font_color is not None:
        run.font.color.rgb = font_color
    if highlight_color is not None:
        run.font.highlight_color = highlight_color


def set_cell_text_centered(cell, text, bold=False, underline=False, font_name=None, font_size=None, highlight_color=None):
    set_cell_text(cell, text, bold=bold, underline=underline, font_name=font_name, font_size=font_size, highlight_color=highlight_color)
    cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER


def set_cell_text_right(cell, text, bold=False, underline=False, font_name=None, font_size=None, highlight_color=None):
    set_cell_text(cell, text, bold=bold, underline=underline, font_name=font_name, font_size=font_size, highlight_color=highlight_color)
    cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT


def emphasize_in_paragraph(paragraph, target, bold=False, underline=False):
    if not target:
        return

    for run in list(paragraph.runs):
        text = run.text
        if target not in text:
            continue

        parts = text.split(target)
        run.text = parts[0]

        # Insert alternating emphasized target and normal tail segments.
        for idx, tail in enumerate(parts[1:]):
            target_run = paragraph.add_run(target)
            target_run.bold = bold
            target_run.underline = underline
            if tail:
                paragraph.add_run(tail)


def style_text_in_paragraph(paragraph, target, font_name=None, bold=None, underline=None, font_size=None):
    if not target:
        return

    for run in list(paragraph.runs):
        text = run.text
        if target not in text:
            continue

        parts = text.split(target)
        run.text = parts[0]

        for tail in parts[1:]:
            target_run = paragraph.add_run(target)
            if font_name is not None:
                target_run.font.name = font_name
            if bold is not None:
                target_run.bold = bold
            if underline is not None:
                target_run.underline = underline
            if font_size is not None:
                target_run.font.size = Pt(font_size)
            if tail:
                paragraph.add_run(tail)


def set_spec_cell(cell, model, spec_text):
    lines = parse_specs_text(spec_text)
    cell.text = ""
    paragraph = cell.paragraphs[0]
    model_run = paragraph.add_run(model.strip())
    model_run.bold = True
    model_run.font.name = "Times New Roman"
    model_run.font.size = Pt(9)

    for line in lines:
        line_run = paragraph.add_run("\n" + line)
        line_run.font.name = "Times New Roman"
        line_run.font.size = Pt(9)


def get_template_path():
    if getattr(sys, "frozen", False):
        base_dir = getattr(sys, "_MEIPASS", BASE_DIR)
    else:
        base_dir = BASE_DIR

    candidates = [
        os.path.join(base_dir, "quotation_template.docx"),
        os.path.join(base_dir, "quotation_template.docx.docx"),
    ]
    for path in candidates:
        if os.path.exists(path):
            return path
    raise FileNotFoundError("quotation_template.docx")


def find_main_table(doc):
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                if "Model & Specification of Commodities" in cell.text:
                    return table
    raise ValueError("Main proforma table not found in template")


def find_row_index_by_text(table, marker):
    for idx, row in enumerate(table.rows):
        if any(marker in cell.text for cell in row.cells):
            return idx
    raise ValueError(f"Row containing '{marker}' not found")


def clone_row_before(table, source_row, row_before):
    new_tr = deepcopy(source_row._tr)
    row_before._tr.addprevious(new_tr)
    return _Row(new_tr, table)


def populate_item_row(row, item, item_no, currency):
    amount = item["unit_price"] * item["qty"]
    currency = get_currency_prefix(currency)
    set_cell_text_centered(row.cells[0], f"{item_no:02d}", bold=True, font_name="Times New Roman", font_size=9)
    set_spec_cell(row.cells[1], item["model"], item["spec"])
    set_cell_text_centered(row.cells[8], f"{currency}\n{format_currency(item['unit_price'])}\n{item['trade_terms']}", font_name="Times New Roman", font_size=9)
    set_cell_text_centered(row.cells[13], str(item["qty"]), bold=True, font_name="Times New Roman", font_size=9)
    set_cell_text_centered(row.cells[16], f"{currency}\n{format_currency(amount)}", font_name="Times New Roman", font_size=9)


def populate_notice_row(row, model, item_no, delivery_days, warranty_hours, warranty_months, unit):
    warranty_value = f"{warranty_hours:,}" if unit == "km" else str(warranty_hours)
    set_cell_text_centered(row.cells[0], f"{item_no:02d}", bold=True, font_name="Times New Roman", font_size=10.5)
    set_cell_text_centered(row.cells[2], model, font_name="Times New Roman", font_size=9)
    if delivery_days <= 30:
        delivery_text = f"{delivery_days} working days after full payment"
    else:
        delivery_text = f"{delivery_days} days after full payment"
    set_cell_text_centered(row.cells[6], delivery_text, font_name="Times New Roman", font_size=9)
    set_cell_text_centered(
        row.cells[14],
        f"Within {warranty_months} months or {warranty_value} {unit}, whichever comes first",
        font_name="Times New Roman",
        font_size=9,
    )


def normalize_products():
    for idx, product in enumerate(products, start=1):
        product["no"] = idx
        product["amount"] = round(product["unit_price"] * product["qty"], 2)


def get_income_advance_base(model_text, currency):
    model_text = model_text.upper()
    if "4X2" in model_text:
        return 200 if currency == "USD" else 700000
    if "6X4" in model_text:
        return 350 if currency == "USD" else 1300000
    if "8X4" in model_text:
        return 400 if currency == "USD" else 1500000
    if "3-AXLE" in model_text:
        return 450 if currency == "USD" else 1650000
    if "8X8" in model_text:
        return 400 if currency == "USD" else 1500000
    return 0


def compute_income_advance_tax():
    currency = get_currency_prefix(currency_var.get())
    total_tax = 0
    for product in products:
        base = get_income_advance_base(product["model"], currency)
        if base:
            total_tax += base * product["qty"]
    return round(total_tax, 2)


def refresh_income_advance_tax():
    if brand_var.get() == "Sinotruk":
        income_tax = compute_income_advance_tax()
        extra_line_label_var.set("The Income Advance Tax")
        extra_line_amount_var.set(income_tax)
        extra_line_enabled_var.set(income_tax > 0)
    else:
        # keep manual extra line settings for other brands
        pass


def refresh_treeview():
    normalize_products()
    currency = get_currency_prefix(currency_var.get())
    for item_id in tree.get_children():
        tree.delete(item_id)

    for product in products:
        tree.insert(
            "",
            "end",
            values=(
                product["no"],
                product["model"],
                product["spec"],
                f"{currency} {format_currency(product['unit_price'])}",
                product["trade_terms"],
                product.get("delivery_days", 45),
                product["qty"],
                f"{currency} {format_currency(product['amount'])}",
            ),
        )


def clear_item_inputs():
    model_var.set("")
    spec_var.set("")
    terms_var.set("DUTY INCLUSIVE")
    qty_var.set(1)
    price_var.set("")
    delivery_days_var.set(45)
    warranty_hours_var.set("2000")
    warranty_months_var.set("12")


def clear_all_inputs():
    products.clear()
    refresh_treeview()
    pi_var.set("")
    sales_var.set("")
    client_var.set("")
    tel_var.set("")
    email_var.set("")
    issue_date.set_date(datetime.now().date())
    valid_days_var.set(30)
    valid_date.set_date(issue_date.get_date() + timedelta(days=30))
    brand_var.set("XCMG")
    delivery_days_var.set(45)
    clear_item_inputs()
    update_warranty()

# ====================== Main Window ======================
root = tk.Tk()
root.title("Proforma Invoice Generator - XCMG, Sinotruk, Heli")
root.geometry("900x720")
root.resizable(True, True)

main_container = ttk.Frame(root)
main_container.pack(fill="both", expand=True)

app_canvas = tk.Canvas(main_container, highlightthickness=0)
app_scrollbar = ttk.Scrollbar(main_container, orient="vertical", command=app_canvas.yview)
app_canvas.configure(yscrollcommand=app_scrollbar.set)

app_canvas.pack(side="left", fill="both", expand=True)
app_scrollbar.pack(side="right", fill="y")

content_frame = ttk.Frame(app_canvas)
content_window_id = app_canvas.create_window((0, 0), window=content_frame, anchor="nw")


def on_content_configure(_event):
    app_canvas.configure(scrollregion=app_canvas.bbox("all"))


def on_canvas_configure(event):
    app_canvas.itemconfigure(content_window_id, width=event.width)


content_frame.bind("<Configure>", on_content_configure)
app_canvas.bind("<Configure>", on_canvas_configure)

# Variables
products = []  # list of dicts for dynamic rows
editing_index = None
brand_var = tk.StringVar(value="XCMG")
unit_var = tk.StringVar(value="hours")
valid_days_var = tk.IntVar(value=30)
delivery_days_var = tk.IntVar(value=45)
warranty_hours_var = tk.StringVar(value="2000")
warranty_months_var = tk.StringVar(value="12")

def update_warranty(*args):
    brand = brand_var.get()
    if brand == "XCMG":
        warranty_months_var.set("12")
        warranty_hours_var.set("2000")
        unit_var.set("hours")
    elif brand == "Sinotruk":
        warranty_months_var.set("12")
        warranty_hours_var.set("60000")
        unit_var.set("km")
    elif brand == "Heli":
        warranty_months_var.set("12")
        warranty_hours_var.set("2000")
        unit_var.set("hours")
    refresh_income_advance_tax()

def update_valid_date(*args):
    try:
        issue = issue_date.get_date()
        days = valid_days_var.get()
        valid = issue + timedelta(days=days)
        valid_date.set_date(valid)
    except Exception:
        pass

brand_var.trace_add("write", update_warranty)
valid_days_var.trace_add("write", update_valid_date)

# ==================== Input Fields ====================
header_buttons_frame = ttk.Frame(content_frame)
header_buttons_frame.pack(fill="x", padx=15, pady=(10, 0))

ttk.Button(header_buttons_frame, text="Flush All", command=clear_all_inputs).pack(side="right")

info_frame = ttk.LabelFrame(content_frame, text="Proforma Information", padding=10)
info_frame.pack(fill="x", padx=15, pady=10)

ttk.Label(info_frame, text="PI Number:").grid(row=0, column=0, sticky="e", pady=4, padx=(0,10))
pi_var = tk.StringVar(value="")
ttk.Entry(info_frame, textvariable=pi_var, width=25).grid(row=0, column=1, padx=(0,5), sticky="w")

ttk.Label(info_frame, text="Sales Officer:").grid(row=1, column=0, sticky="e", pady=4, padx=(0,10))
sales_var = tk.StringVar(value="")
ttk.Entry(info_frame, textvariable=sales_var, width=25).grid(row=1, column=1, padx=(0,5), sticky="w")

ttk.Label(info_frame, text="Client (To):").grid(row=2, column=0, sticky="e", pady=4, padx=(0,10))
client_var = tk.StringVar(value="")
ttk.Entry(info_frame, textvariable=client_var, width=40).grid(row=2, column=1, padx=(0,5), columnspan=2, sticky="w")

ttk.Label(info_frame, text="TEL:").grid(row=3, column=0, sticky="e", pady=4, padx=(0,10))
tel_var = tk.StringVar(value="")
ttk.Entry(info_frame, textvariable=tel_var, width=25).grid(row=3, column=1, padx=(0,5), sticky="w")

ttk.Label(info_frame, text="E-mail:").grid(row=4, column=0, sticky="e", pady=4, padx=(0,10))
email_var = tk.StringVar()
ttk.Entry(info_frame, textvariable=email_var, width=40).grid(row=4, column=1, padx=(0,5), columnspan=2, sticky="w")

# Dates
ttk.Label(info_frame, text="Date of Issue:").grid(row=5, column=0, sticky="e", pady=4, padx=(0,10))
issue_date = DateEntry(info_frame, width=15, date_pattern="yyyy-MM-dd")
issue_date.grid(row=5, column=1, padx=(0,5), sticky="w")

ttk.Label(info_frame, text="Valid To:").grid(row=6, column=0, sticky="e", pady=4, padx=(0,10))
valid_date = DateEntry(info_frame, width=15, date_pattern="yyyy-MM-dd")
valid_date.grid(row=6, column=1, padx=(0,5), sticky="w")

ttk.Label(info_frame, text="Valid Days:").grid(row=7, column=0, sticky="e", pady=4, padx=(0,10))
ttk.Spinbox(info_frame, from_=1, to=365, textvariable=valid_days_var, width=10).grid(row=7, column=1, padx=(0,5), sticky="w")

ttk.Label(info_frame, text="Brand:").grid(row=8, column=0, sticky="e", pady=4, padx=(0,10))
brand_box = ttk.Combobox(info_frame, textvariable=brand_var, values=("XCMG", "Sinotruk", "Heli"), width=12, state="readonly")
brand_box.grid(row=8, column=1, padx=(0,5), sticky="w")

issue_date.bind("<<DateEntrySelected>>", update_valid_date)
update_warranty()
update_valid_date()

# ==================== Products Section ====================
prod_frame = ttk.LabelFrame(content_frame, text="Products (Click + to add more)", padding=10)
prod_frame.pack(fill="both", expand=True, padx=15, pady=10)

# Treeview to show added products
columns = ("No", "Model", "Specification", "Unit Price", "Trade Terms", "Delivery", "Qty", "Amount")
tree_container = ttk.Frame(prod_frame)
tree_container.pack(fill="both", expand=True, pady=5)

tree = ttk.Treeview(tree_container, columns=columns, show="headings", height=8)
for col in columns:
    tree.heading(col, text=col)
    tree.column(col, width=120, anchor="center")

tree.column("Specification", width=280, anchor="w")

tree_scroll_y = ttk.Scrollbar(tree_container, orient="vertical", command=tree.yview)
tree_scroll_x = ttk.Scrollbar(tree_container, orient="horizontal", command=tree.xview)
tree.configure(yscrollcommand=tree_scroll_y.set, xscrollcommand=tree_scroll_x.set)

tree.grid(row=0, column=0, sticky="nsew")
tree_scroll_y.grid(row=0, column=1, sticky="ns")
tree_scroll_x.grid(row=1, column=0, sticky="ew")
tree_container.grid_rowconfigure(0, weight=1)
tree_container.grid_columnconfigure(0, weight=1)

# Add Product Form
add_frame = ttk.Frame(prod_frame)
add_frame.pack(fill="x", pady=8)
add_frame.grid_columnconfigure(1, weight=1)
add_frame.grid_columnconfigure(3, weight=1)

ttk.Label(add_frame, text="Model:").grid(row=0, column=0, sticky="e", padx=(0,10), pady=4)
model_var = tk.StringVar()
ttk.Entry(add_frame, textvariable=model_var, width=25).grid(row=0, column=1, padx=(0,5), pady=4, sticky="w")

ttk.Label(add_frame, text="Specification:").grid(row=1, column=0, sticky="e", padx=(0,10), pady=4)
spec_var = tk.StringVar()
ttk.Entry(add_frame, textvariable=spec_var, width=50).grid(row=1, column=1, padx=(0,5), pady=4, columnspan=3, sticky="w")

ttk.Label(add_frame, text="Currency:").grid(row=2, column=0, sticky="e", padx=(0,10), pady=4)
currency_var = tk.StringVar(value="USD")
currency_box = ttk.Combobox(add_frame, textvariable=currency_var, values=("USD", "UGX"), width=12, state="readonly")
currency_box.grid(row=2, column=1, padx=(0,5), sticky="w")

unit_price_label = ttk.Label(add_frame, text="Unit Price (USD):")
unit_price_label.grid(row=2, column=2, sticky="e", padx=(0,10), pady=4)
price_var = tk.StringVar(value="")
ttk.Entry(add_frame, textvariable=price_var, width=15).grid(row=2, column=3, padx=(0,5), pady=4, sticky="w")

ttk.Label(add_frame, text="Trade Terms:").grid(row=3, column=0, sticky="e", padx=(0,10), pady=4)
terms_var = tk.StringVar(value="DUTY INCLUSIVE")
ttk.Entry(add_frame, textvariable=terms_var, width=20).grid(row=3, column=1, padx=(0,5), sticky="w")

tk.Label(add_frame, text="Qty:").grid(row=3, column=2, sticky="e", padx=(0,10), pady=4)
qty_var = tk.IntVar(value=1)
ttk.Spinbox(add_frame, from_=1, to=100, textvariable=qty_var, width=10).grid(row=3, column=3, padx=(0,5), sticky="w")

ttk.Label(add_frame, text="Delivery Period (days):").grid(row=4, column=0, sticky="e", padx=(0,10), pady=4)
ttk.Spinbox(add_frame, from_=1, to=365, textvariable=delivery_days_var, width=10).grid(row=4, column=1, padx=(0,5), sticky="w")

ttk.Label(add_frame, text="Warranty (hours/km):").grid(row=5, column=0, sticky="e", padx=(0,10), pady=4)
warranty_hours_var = tk.StringVar(value="")
ttk.Entry(add_frame, textvariable=warranty_hours_var, width=15).grid(row=5, column=1, padx=(0,5), pady=4, sticky="w")

ttk.Label(add_frame, text="Warranty (months):").grid(row=5, column=2, sticky="e", padx=(0,10), pady=4)
warranty_months_var = tk.StringVar(value="")
ttk.Entry(add_frame, textvariable=warranty_months_var, width=10).grid(row=5, column=3, padx=(0,5), pady=4, sticky="w")

extra_line_enabled_var = tk.BooleanVar(value=False)
extra_line_label_var = tk.StringVar(value="The Income Advance Tax")
extra_line_amount_var = tk.DoubleVar(value=0.00)

ttk.Checkbutton(add_frame, text="The Income Advance Tax", variable=extra_line_enabled_var).grid(row=6, column=0, columnspan=2, sticky="w", padx=5, pady=8)

ttk.Label(add_frame, text="Extra line title:").grid(row=7, column=0, sticky="e", padx=(0,10), pady=4)
ttk.Entry(add_frame, textvariable=extra_line_label_var, width=35).grid(row=7, column=1, columnspan=3, padx=(0,5), pady=4, sticky="w")

ttk.Label(add_frame, text="Extra amount:").grid(row=8, column=0, sticky="e", padx=(0,10), pady=4)
ttk.Entry(add_frame, textvariable=extra_line_amount_var, width=15).grid(row=8, column=1, padx=(0,5), pady=4, sticky="w")

def update_unit_price_label(*_):
    currency = get_currency_prefix(currency_var.get())
    unit_price_label.config(text=f"Unit Price ({currency}):")
    refresh_treeview()
    refresh_income_advance_tax()

currency_var.trace_add("write", update_unit_price_label)

def add_product():
    global editing_index
    try:
        price_str = price_var.get().replace(',', '')
        price = float(price_str)
        qty = int(qty_var.get())
        model = model_var.get().strip()
        spec = smart_linebreak_specs(spec_var.get().strip())
        terms = terms_var.get().strip()
        warranty_hours = int(warranty_hours_var.get())
        warranty_months = int(warranty_months_var.get())
        unit = "hours" if brand_var.get() in ("XCMG", "Heli") else "km"
        delivery_days = int(delivery_days_var.get())

        if not model or not spec:
            messagebox.showwarning("Missing Info", "Please enter both model and specification.")
            return

        amount = round(price * qty, 2)
        
        no = len(products) + 1
        prod = {
            "no": no,
            "model": model,
            "spec": spec,
            "unit_price": price,
            "trade_terms": terms,
            "qty": qty,
            "delivery_days": delivery_days,
            "amount": amount,
            "warranty_hours": warranty_hours,
            "warranty_months": warranty_months,
            "unit": unit,
        }
        products.append(prod)
        refresh_treeview()
        refresh_income_advance_tax()
        editing_index = None
        
        # Clear inputs for next item
        clear_item_inputs()
    except Exception as e:
        messagebox.showerror("Error", f"Invalid input: {e}")


def remove_selected_product():
    global editing_index
    selected = tree.selection()
    if not selected:
        messagebox.showwarning("Warning", "Select at least one row to remove.")
        return

    selected_indexes = sorted((tree.index(item_id) for item_id in selected), reverse=True)
    for idx in selected_indexes:
        if 0 <= idx < len(products):
            products.pop(idx)

    refresh_treeview()
    refresh_income_advance_tax()
    editing_index = None


def load_selected_product():
    global editing_index
    selected = tree.selection()
    if len(selected) != 1:
        messagebox.showwarning("Warning", "Select exactly one row to edit.")
        return

    selected_idx = tree.index(selected[0])
    if not (0 <= selected_idx < len(products)):
        messagebox.showerror("Error", "Selected item index is invalid.")
        return

    product = products[selected_idx]
    model_var.set(product["model"])
    spec_var.set(product["spec"])
    price_var.set(str(product["unit_price"]))
    terms_var.set(product["trade_terms"])
    qty_var.set(product["qty"])
    delivery_days_var.set(product.get("delivery_days", 45))
    warranty_hours_var.set(str(product["warranty_hours"]))
    warranty_months_var.set(str(product["warranty_months"]))
    editing_index = selected_idx


def update_selected_product():
    global editing_index
    if editing_index is None:
        messagebox.showwarning("Warning", "Load a row first using 'Edit Selected'.")
        return

    try:
        price_str = price_var.get().replace(',', '')
        price = float(price_str)
        qty = int(qty_var.get())
        model = model_var.get().strip()
        spec = smart_linebreak_specs(spec_var.get().strip())
        terms = terms_var.get().strip()
        warranty_hours = int(warranty_hours_var.get())
        warranty_months = int(warranty_months_var.get())
        unit = "hours" if brand_var.get() in ("XCMG", "Heli") else "km"

        if not model or not spec:
            messagebox.showwarning("Missing Info", "Please enter both model and specification.")
            return

        if not (0 <= editing_index < len(products)):
            messagebox.showerror("Error", "The selected row is no longer available.")
            editing_index = None
            return

        products[editing_index]["model"] = model
        products[editing_index]["spec"] = spec
        products[editing_index]["unit_price"] = price
        products[editing_index]["trade_terms"] = terms
        products[editing_index]["qty"] = qty
        products[editing_index]["delivery_days"] = int(delivery_days_var.get())
        products[editing_index]["amount"] = round(price * qty, 2)
        products[editing_index]["warranty_hours"] = warranty_hours
        products[editing_index]["warranty_months"] = warranty_months
        products[editing_index]["unit"] = unit

        refresh_treeview()
        refresh_income_advance_tax()
        editing_index = None
        clear_item_inputs()
    except Exception as e:
        messagebox.showerror("Error", f"Invalid input: {e}")


buttons_frame = ttk.Frame(add_frame)
buttons_frame.grid(row=9, column=0, columnspan=4, pady=12)

ttk.Button(buttons_frame, text="+ Add Product", command=add_product, width=14).pack(side="left", padx=8, pady=4)
ttk.Button(buttons_frame, text="Edit Selected", command=load_selected_product, width=14).pack(side="left", padx=8, pady=4)
ttk.Button(buttons_frame, text="Update Selected", command=update_selected_product, width=14).pack(side="left", padx=8, pady=4)
ttk.Button(buttons_frame, text="- Remove Selected", command=remove_selected_product, width=14).pack(side="left", padx=8, pady=4)

# ==================== Generate Button ====================
def generate_proforma():
    if not products:
        messagebox.showwarning("Warning", "Please add at least one product!")
        return
    
    try:
        normalize_products()
        template_path = get_template_path()
        doc = Document(template_path)
        main_table = find_main_table(doc)

        # Header replacements
        replacements = {
            "DQME2025100202M": pi_var.get().strip(),
            "Henock Sean": sales_var.get().strip(),
            "Aahil Ventures LTD": client_var.get().strip(),
            "+256-788801299": tel_var.get().strip(),
            "henock@doubleq.co.ug": email_var.get().strip() or "henock@doubleq.co.ug",
            "2025-OCTOBER-02": issue_date.get_date().strftime("%Y-%B-%d").upper(),
            "2025-NOVEMBER-01": valid_date.get_date().strftime("%Y-%B-%d").upper(),
        }

        for paragraph in doc.paragraphs:
            for old, new in replacements.items():
                if old in paragraph.text:
                    paragraph.text = paragraph.text.replace(old, new)

        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        for old, new in replacements.items():
                            if old in paragraph.text:
                                paragraph.text = paragraph.text.replace(old, new)

        # Keep all filled-in input information in Times New Roman.
        styled_values = [
            (pi_var.get().strip(), True, True, 11),
            (sales_var.get().strip(), False, False, None),
            (client_var.get().strip(), False, False, None),
            (tel_var.get().strip(), False, False, None),
            (email_var.get().strip() or "henock@doubleq.co.ug", False, False, None),
            (issue_date.get_date().strftime("%Y-%B-%d").upper(), False, False, None),
            (valid_date.get_date().strftime("%Y-%B-%d").upper(), False, False, None),
        ]

        for paragraph in doc.paragraphs:
            for value, bold, underline, size in styled_values:
                style_text_in_paragraph(paragraph, value, font_name="Times New Roman", bold=bold, underline=underline, font_size=size)

        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        for value, bold, underline, size in styled_values:
                            style_text_in_paragraph(paragraph, value, font_name="Times New Roman", bold=bold, underline=underline, font_size=size)

        # Dynamic items section
        item_row_idx = find_row_index_by_text(main_table, "XCMG EXCAVATOR")
        total_row_idx = find_row_index_by_text(main_table, "The Total Amount")
        total_row = main_table.rows[total_row_idx]
        template_item_row = main_table.rows[item_row_idx]

        currency = get_currency_prefix(currency_var.get())

        for i, product in enumerate(products, start=1):
            if i == 1:
                current_row = template_item_row
            else:
                current_row = clone_row_before(main_table, template_item_row, total_row)
            populate_item_row(current_row, product, i, currency)

        total_amount = sum(p["amount"] for p in products)

        if brand_var.get() == "Sinotruk":
            income_tax_amount = compute_income_advance_tax()
            if income_tax_amount > 0:
                extra_tax_row = clone_row_before(main_table, total_row, total_row)
                for cell in extra_tax_row.cells:
                    cell.text = ""
                set_cell_text_right(
                    extra_tax_row.cells[1],
                    "The Income Advance Tax",
                    bold=True,
                    underline=True,
                    font_name="Times New Roman",
                    font_size=10,
                    highlight_color=WD_COLOR_INDEX.GRAY_25,
                )
                set_cell_text(
                    extra_tax_row.cells[16],
                    f"{currency} {format_currency(income_tax_amount)}",
                    bold=True,
                    underline=True,
                    font_name="Times New Roman",
                    font_size=10,
                    font_color=RGBColor(0, 0, 0),
                    highlight_color=WD_COLOR_INDEX.GRAY_25,
                )
                extra_tax_row.cells[16].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                total_amount += income_tax_amount

        if brand_var.get() != "Sinotruk" and extra_line_enabled_var.get() and extra_line_label_var.get().strip():
            extra_amount = round(float(extra_line_amount_var.get()), 2)
            if extra_amount:
                total_amount += extra_amount
                extra_row = clone_row_before(main_table, total_row, total_row)
                for cell in extra_row.cells:
                    cell.text = ""
                set_cell_text_right(
                    extra_row.cells[1],
                    extra_line_label_var.get().strip(),
                    bold=True,
                    underline=True,
                    font_name="Times New Roman",
                    font_size=10,
                )
                set_cell_text(
                    extra_row.cells[16],
                    f"{currency} {format_currency(extra_amount)}",
                    bold=True,
                    underline=True,
                    font_name="Times New Roman",
                    font_size=10,
                )
                extra_row.cells[16].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

        total_cell = main_table.rows[find_row_index_by_text(main_table, "The Total Amount")].cells[16]
        set_cell_text(
            total_cell,
            f"{currency} {format_currency(total_amount)}",
            bold=True,
            underline=True,
            font_name="Times New Roman",
            font_size=10,
            font_color=RGBColor(0, 0, 0),
            highlight_color=WD_COLOR_INDEX.GRAY_25,
        )
        total_cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

        total_row = main_table.rows[find_row_index_by_text(main_table, "The Total Amount")]
        for cell in total_row.cells:
            if "The Total Amount" in cell.text:
                set_cell_text_right(
                    cell,
                    "The Total Amount",
                    bold=True,
                    underline=True,
                    font_name="Times New Roman",
                    font_size=10,
                    highlight_color=WD_COLOR_INDEX.GRAY_25,
                )
                break

        # Dynamic notice section (model, delivery, warranty)
        notice_row_idx = find_row_index_by_text(main_table, "Delivery Period")
        notice_header_row = main_table.rows[notice_row_idx]
        notice_detail_row = main_table.rows[notice_row_idx + 1]
        notice_note_row = main_table.rows[notice_row_idx + 2]

        for i, product in enumerate(products, start=1):
            if i == 1:
                notice_row = notice_detail_row
            else:
                notice_row = clone_row_before(main_table, notice_detail_row, notice_note_row)

            delivery_days = int(product.get("delivery_days", 45))

            populate_notice_row(
                notice_row,
                product["model"],
                i,
                delivery_days,
                product["warranty_hours"],
                product["warranty_months"],
                product["unit"],
            )

        output_name = f"{pi_var.get().strip()}.docx"
        output_dir = os.path.join(BASE_DIR, "output")
        os.makedirs(output_dir, exist_ok=True)
        output_path = os.path.join(output_dir, output_name)
        doc.save(output_path)
        
        messagebox.showinfo("Success", f"Proforma saved as:\n{output_name}\n\nLocation: {output_dir}")
        
    except Exception as e:
        messagebox.showerror("Error", f"Failed to generate document:\n{str(e)}")

ttk.Button(content_frame, text="Generate", command=generate_proforma, style="Accent.TButton").pack(pady=20)

# Style
style = ttk.Style()
style.configure("Accent.TButton", font=("Helvetica", 12, "bold"))

root.mainloop()