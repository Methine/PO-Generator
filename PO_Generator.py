import tkinter as tk
from tkinter import ttk, messagebox
from tkcalendar import DateEntry
import sqlite3, os
import webbrowser
from datetime import datetime
import base64
import mimetypes
import ctypes

# ================= Basic Function =================
def fmt_date(d):
    day = d.day
    if 11 <= day <= 13:
        suffix = "th"
    else:
        suffix = {1: "st", 2: "nd", 3: "rd"}.get(day % 10, "th")

    month = d.strftime("%b")  
    year = d.strftime("%Y")  

    return f"{day}{suffix} {month}. {year}"

def parse_amount(s):
    """把 '1,234.56' → 1234.56"""
    try:
        return float(s.replace(",", "").strip())
    except:
        return 0.0


def format_amount(v):
    """把 1234.56 → '1,234.56'"""
    try:
        return f"{v:,.2f}"
    except:
        return "0.00"

os.makedirs("output", exist_ok=True)

# ================= Database =================
conn = sqlite3.connect("po.db")
cur = conn.cursor()

cur.execute("""CREATE TABLE IF NOT EXISTS supplier(
id INTEGER PRIMARY KEY,
supplier_name TEXT,
supplier_address TEXT,
supplier_attn TEXT,
supplier_tel TEXT
)""")

cur.execute("""CREATE TABLE IF NOT EXISTS ship_to(
id INTEGER PRIMARY KEY,
ship_to TEXT,
ship_address TEXT,
ship_attn TEXT,
ship_tel TEXT
)""")

cur.execute("""CREATE TABLE IF NOT EXISTS trade_terms(
id INTEGER PRIMARY KEY,
delivery_terms TEXT,
forwarder TEXT,
payment_terms TEXT,
currency TEXT
)""")

cur.execute("""CREATE TABLE IF NOT EXISTS footer_terms(
id INTEGER PRIMARY KEY,
packing TEXT,
remark TEXT,
issuer_name TEXT
)""")

conn.commit()

# ================= UI =================
ctypes.windll.shcore.SetProcessDpiAwareness(1)
root = tk.Tk()
root.title("PO Generator")
root.geometry("1050x600")
nb = ttk.Notebook(root)
nb.pack(fill="both", expand=True)

widgets = {}

# ---------- General ----------
def add_entry(f, r, label, key, w=60):
    ttk.Label(f, text=label).grid(row=r, column=0, sticky="e", padx=6, pady=4)
    e = ttk.Entry(f, width=w)
    e.grid(row=r, column=1, sticky="w")
    widgets[key] = e

def add_text(f, r, label, key, h=4):
    ttk.Label(f, text=label).grid(row=r, column=0, sticky="ne", padx=6, pady=4)
    t = tk.Text(f, width=60, height=h)
    t.grid(row=r, column=1, sticky="w")
    widgets[key] = t

def img_to_data_uri(path: str) -> str:
        """
        将图片文件转成 data URI: data:image/png;base64,xxxx
        找不到文件时返回空字符串（也可以改成 raise）
        """
        if not os.path.exists(path):
            return ""  # 或者：raise FileNotFoundError(path)

        mime, _ = mimetypes.guess_type(path)
        if not mime:
            mime = "application/octet-stream"

        with open(path, "rb") as fimg:
            b64 = base64.b64encode(fimg.read()).decode("utf-8")
        return f"data:{mime};base64,{b64}"
        
# ================= Supplier =================
f = ttk.Frame(nb); nb.add(f, text="Supplier")
cb = ttk.Combobox(f, width=45, state="readonly")
cb.grid(row=0, column=1, sticky="w")
ttk.Label(f, text="Load Supplier").grid(row=0, column=0)

add_entry(f,1,"Name","supplier_name")
add_text(f,2,"Address","supplier_address")
add_entry(f,3,"Attn","supplier_attn")
add_entry(f,4,"Tel","supplier_tel")

def load_supplier():
    cur.execute("SELECT id,supplier_name FROM supplier")
    cb["values"]=[f"{i} | {n}" for i,n in cur.fetchall()]

def pick_supplier(e):
    i=cb.get().split("|")[0].strip()
    cur.execute("SELECT supplier_name,supplier_address,supplier_attn,supplier_tel FROM supplier WHERE id=?",(i,))
    r=cur.fetchone()
    for k,v in zip(["supplier_name","supplier_address","supplier_attn","supplier_tel"],r):
        if isinstance(widgets[k],tk.Text):
            widgets[k].delete("1.0","end"); widgets[k].insert("1.0",v)
        else:
            widgets[k].delete(0,"end"); widgets[k].insert(0,v)

def save_supplier():
    cur.execute("INSERT INTO supplier VALUES(NULL,?,?,?,?)",
        (widgets["supplier_name"].get(),
         widgets["supplier_address"].get("1.0","end").strip(),
         widgets["supplier_attn"].get(),
         widgets["supplier_tel"].get()))
    conn.commit(); load_supplier()

cb.bind("<<ComboboxSelected>>",pick_supplier)
ttk.Button(f,text="Save Supplier",command=save_supplier).grid(row=5,column=1,sticky="w")
load_supplier()

# ================= PO Info =================
f = ttk.Frame(nb); nb.add(f, text="PO Info")
add_entry(f,0,"PO Number","po_number")
ttk.Label(f,text="PO Date").grid(row=1,column=0,sticky="e")
d1=DateEntry(f); d1.grid(row=1,column=1,sticky="w")
d1.insert(0,fmt_date(datetime.today()))
d1.bind("<<DateEntrySelected>>",lambda e:(d1.delete(0,"end"),d1.insert(0,fmt_date(d1.get_date()))))
widgets["po_date"]=d1
add_entry(f,2,"Internal No","internal_no")
add_entry(f,3,"ETD Port","etd_port")
ttk.Label(f, text="ETD Date").grid(row=4, column=0, sticky="e")

d2 = ttk.Entry(f, width=30)
d2.grid(row=4, column=1, sticky="w")

# 可选：默认值
d2.insert(0, "ASAP")

widgets["etd_date"] = d2

# ================= Ship To =================
f = ttk.Frame(nb); nb.add(f, text="Ship To")
cb2 = ttk.Combobox(f,width=45,state="readonly")
cb2.grid(row=0,column=1,sticky="w")
ttk.Label(f,text="Load Ship To").grid(row=0,column=0)

add_entry(f,1,"Ship To","ship_to")
add_text(f,2,"Address","ship_address")
add_entry(f,3,"Attn","ship_attn")
add_entry(f,4,"Tel","ship_tel")

def load_ship():
    cur.execute("SELECT id,ship_to FROM ship_to")
    cb2["values"]=[f"{i} | {n}" for i,n in cur.fetchall()]

def pick_ship(e):
    i=cb2.get().split("|")[0].strip()
    cur.execute("SELECT ship_to,ship_address,ship_attn,ship_tel FROM ship_to WHERE id=?",(i,))
    r=cur.fetchone()
    for k,v in zip(["ship_to","ship_address","ship_attn","ship_tel"],r):
        if isinstance(widgets[k],tk.Text):
            widgets[k].delete("1.0","end"); widgets[k].insert("1.0",v)
        else:
            widgets[k].delete(0,"end"); widgets[k].insert(0,v)

def save_ship():
    cur.execute("INSERT INTO ship_to VALUES(NULL,?,?,?,?)",
        (widgets["ship_to"].get(),
         widgets["ship_address"].get("1.0","end").strip(),
         widgets["ship_attn"].get(),
         widgets["ship_tel"].get()))
    conn.commit(); load_ship()

cb2.bind("<<ComboboxSelected>>",pick_ship)
ttk.Button(f,text="Save Ship To",command=save_ship).grid(row=5,column=1,sticky="w")
load_ship()

# ================= Trade Terms =================
f = ttk.Frame(nb); nb.add(f, text="Trade Terms")
cb3 = ttk.Combobox(f,width=45,state="readonly")
cb3.grid(row=0,column=1,sticky="w")
ttk.Label(f,text="Load Terms").grid(row=0,column=0)

add_entry(f,1,"Delivery Terms","delivery_terms")
add_entry(f,2,"Forwarder","forwarder")
add_text(f,3,"Payment Terms","payment_terms",h=5)
add_entry(f,4,"Currency","currency")

def load_terms():
    cur.execute("SELECT id,delivery_terms FROM trade_terms")
    cb3["values"]=[f"{i} | {n}" for i,n in cur.fetchall()]

def pick_terms(e):
    i=cb3.get().split("|")[0].strip()
    cur.execute("SELECT delivery_terms,forwarder,payment_terms,currency FROM trade_terms WHERE id=?",(i,))
    r=cur.fetchone()
    for k,v in zip(["delivery_terms","forwarder","payment_terms","currency"],r):
        if isinstance(widgets[k],tk.Text):
            widgets[k].delete("1.0","end"); widgets[k].insert("1.0",v)
        else:
            widgets[k].delete(0,"end"); widgets[k].insert(0,v)

def save_terms():
    cur.execute("INSERT INTO trade_terms VALUES(NULL,?,?,?,?)",
        (widgets["delivery_terms"].get(),
         widgets["forwarder"].get(),
         widgets["payment_terms"].get("1.0","end").strip(),
         widgets["currency"].get()))
    conn.commit(); load_terms()

cb3.bind("<<ComboboxSelected>>",pick_terms)
ttk.Button(f,text="Save Terms",command=save_terms).grid(row=5,column=1,sticky="w")
load_terms()

# ================= Items =================
items_f = ttk.Frame(nb); nb.add(items_f, text="Items")
headers=["Item No","Description","Qty","Unit Price","Total",""]
for i,h in enumerate(headers):
    ttk.Label(items_f,text=h).grid(row=0,column=i)

rows=[]
total_var=tk.StringVar(value="0.00")

def recalc():
    total = 0.0
    for r in rows:
        q = parse_amount(r[2].get())
        p = parse_amount(r[3].get())
        t = q * p

        r[4].configure(state="normal")
        r[4].delete(0, "end")
        r[4].insert(0, format_amount(t))
        r[4].configure(state="readonly")

        total += t

    total_var.set(format_amount(total))



def _bind_recalc(entry):
    entry.bind("<KeyRelease>", lambda event: recalc())

def refresh_rows():
    """把 rows 里现存控件全部重新排版，并重绑删除按钮索引"""
    for idx, line in enumerate(rows):
        grid_row = idx + 1  # header 在 row=0
        for c in range(5):
            line[c].grid(row=grid_row, column=c)
        # line[5] 是删除按钮
        line[5].configure(command=lambda i=idx: remove(i))
        line[5].grid(row=grid_row, column=5)

def add_row():
    line = []
    for c in range(5):
        if c == 4:
            e = tk.Entry(items_f, width=18, state="readonly")
        else:
            e = tk.Entry(items_f, width=18)

        if c in (2, 3):
            _bind_recalc(e)
        line.append(e)

    b = ttk.Button(items_f, text="X")
    line.append(b)

    rows.append(line)
    refresh_rows()

def remove(i):
    for w in rows[i]:
        w.destroy()
    rows.pop(i)
    refresh_rows()
    recalc()

ttk.Button(items_f,text="+ Add Item",command=add_row).grid(row=999,column=0)
ttk.Label(items_f,text="Total Amount").grid(row=999,column=3)
ttk.Entry(items_f,textvariable=total_var,state="readonly").grid(row=999,column=4)
add_row()

# ================= Footer / Issued =================
f = ttk.Frame(nb); nb.add(f, text="Footer / Issued")
cb4 = ttk.Combobox(f,width=45,state="readonly")
cb4.grid(row=0,column=1,sticky="w")
ttk.Label(f,text="Load Footer").grid(row=0,column=0)

add_entry(f,1,"Packing","packing")
add_text(f,2,"Remark","remark",h=4)
add_entry(f,3,"Issuer Name","issuer_name")

def load_footer():
    cur.execute("SELECT id,issuer_name FROM footer_terms")
    cb4["values"]=[f"{i} | {n}" for i,n in cur.fetchall()]

def pick_footer(e):
    i=cb4.get().split("|")[0].strip()
    cur.execute("SELECT packing,remark,issuer_name FROM footer_terms WHERE id=?",(i,))
    r=cur.fetchone()
    for k,v in zip(["packing","remark","issuer_name"],r):
        if isinstance(widgets[k],tk.Text):
            widgets[k].delete("1.0","end"); widgets[k].insert("1.0",v)
        else:
            widgets[k].delete(0,"end"); widgets[k].insert(0,v)

def save_footer():
    cur.execute("INSERT INTO footer_terms VALUES(NULL,?,?,?)",
        (widgets["packing"].get(),
         widgets["remark"].get("1.0","end").strip(),
         widgets["issuer_name"].get()))
    conn.commit(); load_footer()

cb4.bind("<<ComboboxSelected>>",pick_footer)
ttk.Button(f,text="Save Footer",command=save_footer).grid(row=4,column=1,sticky="w")
load_footer()

# ================= Generate =================
def generate():
    with open("template.html","r",encoding="utf-8") as f:
        html=f.read()

    def gv(k):
        w=widgets[k]
        return w.get("1.0","end").strip() if isinstance(w,tk.Text) else w.get()
        
    html = html.replace("{{logo_b64}}", img_to_data_uri("logo.png"))
    html = html.replace("{{stamp_b64}}", img_to_data_uri("stamp.png"))
    html = html.replace("{{sales_rep_stamp_b64}}", img_to_data_uri("sales_rep_stamp.png"))
    
    for k in widgets:
        if k == "po_date":
            v = fmt_date(widgets["po_date"].get_date())
        else:
            w = widgets[k]
            if isinstance(w, tk.Text):
                v = w.get("1.0", "end").strip()
                v = v.replace("\n", "<br>")
            else:
                v = w.get()

        html = html.replace(f"{{{{{k}}}}}", v)


    item_html=""
    for r in rows:
        item_html+=f"<tr><td>{r[0].get()}</td><td>{r[1].get()}</td><td>{r[2].get()}</td><td>{r[3].get()}</td><td>{r[4].get()}</td></tr>\n"

    html=html.replace("{{items_html}}",item_html)
    html=html.replace("{{total_amount}}",total_var.get())

    fn=f"output/PO_{datetime.now().strftime('%Y%m%d_%H%M%S')}.html"
    open(fn,"w",encoding="utf-8").write(html)
    webbrowser.open(f"file:///{os.path.abspath(fn)}")
    messagebox.showinfo("Done",f"Generated:\n{fn}")

ttk.Button(root,text="Generate PO HTML, open in browser",command=generate).pack(pady=10)
root.mainloop()
