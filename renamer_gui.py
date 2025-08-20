# -*- coding: utf-8 -*-
import os, re, json, sys
from pathlib import Path
from typing import Optional, List, Tuple, Dict

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

try:
    from PIL import Image, ImageOps, ImageFilter, ImageTk
except Exception:
    Image = None; ImageOps = None; ImageFilter = None; ImageTk = None

try:
    import pytesseract
except Exception:
    pytesseract = None

APP_TITLE = "Переименование отсканированных файлов — Ella Renamer v3.2 (зоны+якоря, ручное выделение)"
SUPPORTED_EXTS = (".jpg",".jpeg",".png",".webp",".tif",".tiff",".bmp",".jfif")
PATTERN = re.compile(r"(2711\d{4}|20\d{6})")
CONFIG_NAME = "config.json"

def app_dir()->Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).parent
    return Path(__file__).parent

def load_config()->dict:
    cfg_path = app_dir() / CONFIG_NAME
    if cfg_path.exists():
        try:
            return json.loads(cfg_path.read_text(encoding="utf-8"))
        except Exception:
            pass
    return {"tesseract_path": r"C:\Program Files\Tesseract-OCR\tesseract.exe",
            "use_rotations": True}

def save_config(cfg: dict):
    (app_dir() / CONFIG_NAME).write_text(json.dumps(cfg, ensure_ascii=False, indent=2), encoding="utf-8")

def ensure_tess(cfg: dict):
    global pytesseract
    if pytesseract is None: return
    path = cfg.get("tesseract_path") or ""
    if os.name=="nt" and path and os.path.exists(path):
        pytesseract.pytesseract.tesseract_cmd = path

def parse_date_from_folder(folder: Path)->Optional[str]:
    name = folder.name
    m = re.search(r"(\d{2})[.\-_/](\d{2})[.\-_/](\d{4})", name)
    if m: return f"{m.group(1)}.{m.group(2)}.{m.group(3)}"
    m = re.search(r"(\d{4})[.\-_/](\d{2})[.\-_/](\d{2})", name)
    if m: return f"{m.group(3)}.{m.group(2)}.{m.group(1)}"
    m = re.search(r"(\d{8})", name)
    if m:
        raw = m.group(1)
        if raw.startswith("20"):
            y,mm,d = raw[:4], raw[4:6], raw[6:]
        else:
            d,mm,y = raw[:2], raw[2:4], raw[4:]
        return f"{d}.{mm}.{y}"
    return None

def pil_open(path: Path):
    img = Image.open(path)
    try: img = ImageOps.exif_transpose(img)
    except Exception: pass
    return img

def preprocess(img):
    img = img.convert("L")
    try: img = ImageOps.autocontrast(img)
    except Exception: pass
    try: img = img.filter(ImageFilter.SHARPEN)
    except Exception: pass
    return img

# --- OCR helpers ---
def ocr_data_words(pil, lang="rus+eng")->List[Dict]:
    if pytesseract is None: return []
    try:
        d = pytesseract.image_to_data(pil, lang=lang, config="--psm 6", output_type=pytesseract.Output.DICT)
        out=[]; n=len(d["text"])
        for i in range(n):
            txt=(d["text"][i] or "").strip()
            if not txt: continue
            out.append({"text": txt, "left": d["left"][i], "top": d["top"][i], "width": d["width"][i], "height": d["height"][i]})
        return out
    except Exception:
        return []

def ocr_text_digits(pil)->str:
    if pytesseract is None: return ""
    try:
        return pytesseract.image_to_string(pil, lang="eng",
            config="--psm 6 -c tessedit_char_whitelist=0123456789")
    except Exception:
        return ""

def ocr_data_digits(pil)->List[Dict]:
    if pytesseract is None: return []
    try:
        d = pytesseract.image_to_data(pil, lang="eng",
            config="--psm 6 -c tessedit_char_whitelist=0123456789",
            output_type=pytesseract.Output.DICT)
        out=[]; n=len(d["text"])
        for i in range(n):
            txt=(d["text"][i] or "").strip().replace(" ","")
            if not txt: continue
            out.append({"text": txt, "left": d["left"][i], "top": d["top"][i],
                        "width": d["width"][i], "height": d["height"][i]})
        return out
    except Exception:
        return []

# --- зоны поиска (стандарт/широкие) ---
def search_regions(img, wide=False):
    w,h=img.size
    if not wide:
        return [
            ((int(w*0.55), 0, int(w*0.995), int(h*0.40)), "правый верх"),
            ((0, 0, int(w*0.22), h), "левая вертикаль"),
        ]
    else:
        return [
            ((int(w*0.50), 0, int(w*0.995), int(h*0.55)), "правый верх (шир.)"),
            ((0, 0, int(w*0.28), h), "левая вертикаль (шир.)"),
        ]

def orientations(img, use_rot=True):
    lst=[(0, img)]
    if not use_rot: return lst
    try:
        lst += [(90, img.rotate(90,expand=True)),
                (180, img.rotate(180,expand=True)),
                (270, img.rotate(270,expand=True))]
    except Exception: pass
    return lst

# --- Поиск рядом с якорями «отправка №», «№» ---
def find_near_anchor(region_img)->Tuple[Optional[str], Optional[Tuple[int,int,int,int]], str]:
    words = ocr_data_words(preprocess(region_img))
    W,H = region_img.size
    norm = [(w["text"].lower().replace("ё","е"), w) for w in words]
    candidates = []
    for i,(t,w) in enumerate(norm):
        has_no = ("№" in w["text"]) or (t=="no") or (t=="n°") or (t=="№")
        has_otpr = ("отправк" in t)
        if has_no and i>0 and "отправк" in norm[i-1][0]:
            anchor = w
            l = anchor["left"] + anchor["width"]
            t0 = max(0, anchor["top"] - int(anchor["height"]*0.5))
            r = min(W, l + int(W*0.40))
            b = min(H, anchor["top"] + int(anchor["height"]*1.8))
            candidates.append((l,t0,r,b,"по якорю: отправка №"))
        elif has_no:
            anchor = w
            l = anchor["left"] + anchor["width"]
            t0 = max(0, anchor["top"] - int(anchor["height"]*0.5))
            r = min(W, l + int(W*0.40))
            b = min(H, anchor["top"] + int(anchor["height"]*1.8))
            candidates.append((l,t0,r,b,"по якорю: №"))
        elif has_otpr:
            anchor = w
            l = anchor["left"] + anchor["width"]
            t0 = max(0, anchor["top"] - int(anchor["height"]*0.5))
            r = min(W, l + int(W*0.45))
            b = min(H, anchor["top"] + int(anchor["height"]*1.8))
            candidates.append((l,t0,r,b,"по якорю: отправка"))
    for (l,t0,r,b,source) in candidates:
        roi = region_img.crop((l,t0,r,b))
        for item in ocr_data_digits(preprocess(roi)):
            m = PATTERN.fullmatch(item["text"].replace(" ", ""))
            if m:
                L = l + item["left"]; T = t0 + item["top"]
                R = L + item["width"]; B = T + item["height"]
                return m.group(1), (L,T,R,B), source
        text = ocr_text_digits(preprocess(roi)).replace(" ", "")
        m = PATTERN.search(text)
        if m:
            return m.group(1), (l,t0,r,b), source
    return None, None, ""

def extract_number_debug(path: Path, cfg: dict, wide=False):
    if Image is None:
        return {"serial": None, "angle":0, "bbox": None, "source": None, "img": None,
                "regions": []}
    try: base = pil_open(path)
    except Exception:
        return {"serial": None, "angle":0, "bbox": None, "source": None, "img": None,
                "regions": []}

    best = {"serial": None, "angle":0, "bbox": None, "source": None, "img": base,
            "regions": []}
    for angle, img in orientations(base, cfg.get("use_rotations", True)):
        regs = search_regions(img, wide)
        if angle==0: best["regions"] = regs
        # 1) По якорям
        for bbox_reg, lbl in regs:
            region = img.crop(bbox_reg)
            serial, box, why = find_near_anchor(region)
            if serial:
                if box:
                    L = bbox_reg[0] + box[0]; T = bbox_reg[1] + box[1]
                    R = bbox_reg[0] + box[2]; B = bbox_reg[1] + box[3]
                    gb = (L,T,R,B)
                else:
                    gb = bbox_reg
                return {"serial": serial, "angle": angle, "bbox": gb,
                        "source": f"{lbl}, {why}", "img": img, "regions": regs}
        # 2) Фолбэк: цифры внутри зон
        for bbox_reg, lbl in regs:
            region = img.crop(bbox_reg)
            for item in ocr_data_digits(preprocess(region)):
                m = PATTERN.fullmatch(item["text"].replace(" ", ""))
                if m:
                    L = bbox_reg[0] + item["left"]; T = bbox_reg[1] + item["top"]
                    R = L + item["width"]; B = T + item["height"]
                    return {"serial": m.group(1), "angle": angle, "bbox": (L,T,R,B),
                            "source": f"{lbl}, без якоря", "img": img, "regions": regs}
    return best

def propose(date_str:str, serial:str, ext:str, existing:set, seen:dict)->str:
    base = f"{date_str} {serial}"
    seen[base] = seen.get(base,0)+1
    cand = base if seen[base]==1 else f"{base} 1"
    name = f"{cand}{ext}"
    i=2
    while name in existing:
        name = f"{base} {i-1}{ext}"; i+=1
    return name

class App(ttk.Frame):
    def __init__(self, master):
        super().__init__(master); self.pack(fill="both", expand=True)
        master.title(APP_TITLE); master.geometry("1220x680")
        self.cfg = load_config(); ensure_tess(self.cfg)
        self.dir_var=tk.StringVar(); self.date_var=tk.StringVar()
        self.tess_var=tk.StringVar(value=self.cfg.get("tesseract_path",""))
        self.status=tk.StringVar(value="Готово"); self.rows=[]; self.preview_cache={}
        self._photo=None
        self.wide_var = tk.BooleanVar(value=False)   # «Широкие зоны»
        # для ручного выделения
        self.sel_start=None; self.sel_rect=None
        self.disp_scale=1.0; self.disp_off=(0,0); self.last_img=None
        self.build()
        # хоткеи
        master.bind("<F5>", lambda e: self.preview())
        master.bind("<F6>", lambda e: self.apply())
        master.bind("<F9>", lambda e: self.recognize_in_selection())
        master.bind("<Control-o>", lambda e: self.choose_dir())

    def build(self):
        top=ttk.Frame(self); top.pack(fill="x", padx=10, pady=8)
        ttk.Label(top,text="Папка с файлами:").grid(row=0,column=0,sticky="w")
        ttk.Entry(top,textvariable=self.dir_var,width=76).grid(row=0,column=1,sticky="we",padx=6)
        ttk.Button(top,text="Выбрать...",command=self.choose_dir).grid(row=0,column=2,padx=4)

        ttk.Label(top,text="Дата (ДД.ММ.ГГГГ):").grid(row=1,column=0,sticky="w")
        ttk.Entry(top,textvariable=self.date_var,width=20).grid(row=1,column=1,sticky="w")
        ttk.Button(top,text="Из имени папки",command=self.fill_date).grid(row=1,column=2,padx=4,sticky="w")

        ttk.Label(top,text="Путь к tesseract.exe:").grid(row=2,column=0,sticky="w")
        ttk.Entry(top,textvariable=self.tess_var,width=76).grid(row=2,column=1,sticky="we",padx=6)
        ttk.Button(top,text="Найти...",command=self.choose_tess).grid(row=2,column=2,padx=4)
        ttk.Button(top,text="Сохранить",command=self.save_cfg).grid(row=2,column=3,padx=4)
        ttk.Checkbutton(top,text="Широкие зоны",variable=self.wide_var).grid(row=2,column=4,padx=(10,0))

        toolbar=ttk.Frame(top); toolbar.grid(row=3, column=0, columnspan=5, sticky="we", pady=(8,0))
        ttk.Button(toolbar,text="Предпросмотр (F5)",command=self.preview).pack(side="left",padx=4)
        ttk.Button(toolbar,text="Переименовать (F6)",command=self.apply).pack(side="left",padx=4)
        ttk.Button(toolbar,text="Открыть папку",command=self.open_folder).pack(side="left",padx=4)
        ttk.Button(toolbar,text="Распознать в выделенной области (F9)",command=self.recognize_in_selection).pack(side="left",padx=10)

        manual=ttk.Frame(top); manual.grid(row=4, column=0, columnspan=5, sticky="we", pady=(6,0))
        ttk.Label(manual, text="Ручной ввод номера (для выбранного файла):").pack(side="left")
        self.manual_var = tk.StringVar()
        ttk.Entry(manual, textvariable=self.manual_var, width=22).pack(side="left", padx=6)
        ttk.Button(manual, text="Применить к выбранному", command=self.apply_manual).pack(side="left")

        main=ttk.Frame(self); main.pack(fill="both", expand=True, padx=10, pady=6)
        left=ttk.Frame(main); left.pack(side="left", fill="both", expand=True)
        self.tree=ttk.Treeview(left,columns=("old","serial","new","status"),show="headings",height=22)
        for k,t,w in [("old","Старое имя",380),("serial","Номер (8 цифр)",140),("new","Новое имя",320),("status","Статус",110)]:
            self.tree.heading(k,text=t); self.tree.column(k,width=w,anchor="center" if k!="old" else "w")
        self.tree.pack(side="left",fill="both",expand=True)
        vsb=ttk.Scrollbar(left,orient="vertical",command=self.tree.yview); vsb.pack(side="right",fill="y")
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.bind("<<TreeviewSelect>>", self.on_select_row)

        right=ttk.Frame(main); right.pack(side="left", fill="both", expand=False, padx=(10,0))
        self.canvas = tk.Canvas(right, width=560, height=560, bg="#f6f6f6",
                                highlightthickness=1, highlightbackground="#cccccc")
        self.canvas.pack(fill="both", expand=False)
        # события для ручного выделения
        self.canvas.bind("<Button-1>", self.on_canvas_down)
        self.canvas.bind("<B1-Motion>", self.on_canvas_drag)
        self.canvas.bind("<ButtonRelease-1>", self.on_canvas_up)

        ttk.Label(self,textvariable=self.status,anchor="w").pack(fill="x", padx=10, pady=(0,8))
        top.columnconfigure(1, weight=1)

    # ==== утилиты ====
    def choose_dir(self):
        p=filedialog.askdirectory(title="Выберите папку с файлами")
        if p: self.dir_var.set(p); self.fill_date()

    def choose_tess(self):
        p=filedialog.askopenfilename(title="Выберите tesseract.exe",
            filetypes=[("tesseract.exe","tesseract.exe"),("Все файлы","*.*")])
        if p: self.tess_var.set(p)

    def save_cfg(self):
        cfg = load_config()
        cfg["tesseract_path"]=self.tess_var.get().strip()
        save_config(cfg); ensure_tess(cfg); self.status.set("Путь к Tesseract сохранён.")

    def fill_date(self):
        f=Path(self.dir_var.get().strip())
        if f.exists():
            ds=parse_date_from_folder(f)
            if ds: self.date_var.set(ds)

    def files(self)->List[Path]:
        f=Path(self.dir_var.get().strip())
        if not f.exists(): return []
        return [f/n for n in sorted(os.listdir(f)) if n.lower().endswith(SUPPORTED_EXTS)]

    # ==== основной поток ====
    def preview(self):
        self.tree.delete(*self.tree.get_children()); self.preview_cache.clear()
        self.canvas.delete("all"); self._photo=None; self.sel_clear()
        mode = "Широкие" if self.wide_var.get() else "Стандартные"
        self.status.set(f"Распознавание ({mode} зоны; якоря «отправка №», «№»)..."); self.update_idletasks()

        fs=self.files()
        if not fs: self.status.set("Файлы не найдены."); return
        date=self.date_var.get().strip()
        if not re.fullmatch(r"\d{2}\.\d{2}\.\d{4}", date):
            messagebox.showwarning("Дата","Введите ДД.ММ.ГГГГ или нажмите «Из имени папки»."); self.status.set("Ожидает даты."); return
        folder=Path(self.dir_var.get().strip()); existing=set(os.listdir(folder)); seen={}; self.rows=[]

        for p in fs:
            dbg = extract_number_debug(p, {"use_rotations": True}, wide=self.wide_var.get()) if Image else \
                  {"serial": None, "img": None, "bbox": None, "angle": 0, "source": None, "regions":[]}
            self.preview_cache[p.name] = dbg
            serial = dbg["serial"]
            if not serial: st="НЕ НАЙДЕНО"; newname=""
            else: st="OK"; newname=propose(date,serial,p.suffix.lower(),existing,seen)
            row={"old":p.name,"serial":serial or "","new":newname,"status":st}
            self.rows.append(row); self.tree.insert("", "end", values=(row["old"],row["serial"],row["new"],row["status"]))
        self.status.set("Готово. Выберите строку — справа подсветятся зоны и найденное место.")

    def apply(self):
        if not self.rows: messagebox.showinfo("Нет данных","Сначала выполните предпросмотр."); return
        folder=Path(self.dir_var.get().strip()); ok=0
        for r in self.rows:
            if r["status"]!="OK" or not r["new"]: continue
            src=folder/r["old"]; dst=folder/r["new"]
            try:
                if src.resolve()!=dst.resolve(): src.rename(dst)
                ok+=1; r["status"]="ПЕРЕИМ."
            except FileExistsError: r["status"]="СУЩЕСТВУЕТ"
            except Exception: r["status"]="ОШИБКА"
        self.refresh_tree(); self.status.set(f"Готово. Успешно переименовано: {ok}")

    def refresh_tree(self):
        self.tree.delete(*self.tree.get_children())
        for r in self.rows: self.tree.insert("", "end", values=(r["old"],r["serial"],r["new"],r["status"]))

    def open_folder(self):
        p=self.dir_var.get().strip()
        if not p: return
        try:
            if sys.platform.startswith("win"): os.startfile(p)
            elif sys.platform=="darwin": os.system(f'open "{p}"')
            else: os.system(f'xdg-open "{p}"')
        except Exception: pass

    # ==== правая панель ====
    def on_select_row(self, event=None):
        sel = self.tree.selection()
        if not sel: return
        vals = self.tree.item(sel[0],"values")
        if not vals: return
        filename = vals[0]
        dbg = self.preview_cache.get(filename)
        folder=Path(self.dir_var.get().strip())
        path = folder/filename
        self.show_image_with_overlays(path, dbg)

    def show_image_with_overlays(self, path: Path, dbg: dict):
        self.canvas.delete("all"); self._photo=None; self.sel_clear()
        if Image is None: return
        try:
            img = dbg["img"] if (dbg and dbg.get("img") is not None) else pil_open(path)
        except Exception:
            return
        self.last_img = img
        cw = self.canvas.winfo_width() or 560
        ch = self.canvas.winfo_height() or 560
        w,h = img.size
        scale = min(cw/w, ch/h)
        nw, nh = int(w*scale), int(h*scale)
        disp = img.resize((nw,nh))
        self._photo = ImageTk.PhotoImage(disp)
        x0 = (cw-nw)//2; y0 = (ch-nh)//2
        self.disp_scale = scale
        self.disp_off = (x0,y0)
        self.canvas.create_image(x0, y0, image=self._photo, anchor="nw")

        # подсветка "зон поиска"
        regs = dbg.get("regions") or search_regions(img, self.wide_var.get())
        for (L,T,R,B), _lbl in regs:
            l = int(L*scale)+x0; r=int(R*scale)+x0; t=int(T*scale)+y0; b=int(B*scale)+y0
            self.canvas.create_rectangle(l,t,r,b, outline="#ffb000", dash=(5,3), width=2)

        # найденный bbox
        bbox = dbg.get("bbox") if dbg else None
        if bbox:
            l,t,r,b = bbox
            l = int(l*scale)+x0; r=int(r*scale)+x0; t=int(t*scale)+y0; b=int(b*scale)+y0
            self.canvas.create_rectangle(l,t,r,b, outline="red", width=3)

        cap = ""
        if dbg and dbg.get("source"): cap += dbg["source"]
        if dbg and dbg.get("angle"): cap += f" | поворот {dbg['angle']}°"
        if cap: self.canvas.create_text(10, ch-10, text=cap, anchor="sw", fill="#222222")

    # ====== выделение мышью на canvas ======
    def on_canvas_down(self, e):
        self.sel_clear()
        self.sel_start = (e.x, e.y)
        self.sel_rect = self.canvas.create_rectangle(e.x, e.y, e.x, e.y,
                                                     outline="#0066ff", width=2)

    def on_canvas_drag(self, e):
        if self.sel_rect and self.sel_start:
            x0,y0 = self.sel_start
            self.canvas.coords(self.sel_rect, x0, y0, e.x, e.y)

    def on_canvas_up(self, e):
        # по отпусканию ЛКМ можно сразу OCRнуть выбранную область
        self.recognize_in_selection(auto=True)

    def sel_clear(self):
        if self.sel_rect:
            self.canvas.delete(self.sel_rect)
        self.sel_rect=None; self.sel_start=None

    def selection_bbox_image_coords(self)->Optional[Tuple[int,int,int,int]]:
        """Вернёт прямоугольник выделения в координатах исходного изображения."""
        if not self.sel_rect or not self.last_img: return None
        x0,y0,x1,y1 = self.canvas.coords(self.sel_rect)
        if x0==x1 or y0==y1: return None
        if x0>x1: x0,x1 = x1,x0
        if y0>y1: y0,y1 = y1,y0
        sx, sy = self.disp_off
        scale = self.disp_scale
        # пересечение с областью изображения на canvas
        ix0 = max(x0, sx); iy0 = max(y0, sy)
        ix1 = min(x1, sx + int(self.last_img.size[0]*scale))
        iy1 = min(y1, sy + int(self.last_img.size[1]*scale))
        if ix0>=ix1 or iy0>=iy1: return None
        # перевод в координаты исходного изображения
        L = int((ix0 - sx) / scale)
        T = int((iy0 - sy) / scale)
        R = int((ix1 - sx) / scale)
        B = int((iy1 - sy) / scale)
        # небольшая подушка
        padx = max(2, int((R-L)*0.03)); pady = max(2, int((B-T)*0.03))
        L = max(0, L-padx); T = max(0, T-pady)
        R = min(self.last_img.size[0], R+padx); B = min(self.last_img.size[1], B+pady)
        return (L,T,R,B)

    def recognize_in_selection(self, auto=False):
        """OCR в выделенной области → заполнить номер у выбранного файла."""
        sel = self.tree.selection()
        if not sel:
            if not auto:
                messagebox.showinfo("Нет выбора","Слева выберите файл.")
            return
        if not self.last_img:
            if not auto: messagebox.showinfo("Нет изображения","Нечего распознавать.")
            return
        box = self.selection_bbox_image_coords()
        if not box:
            if not auto: messagebox.showinfo("Нет выделения","Выделите прямоугольник на изображении.")
            return
        L,T,R,B = box
        roi = self.last_img.crop((L,T,R,B))
        # OCR цифр
        serial = None
        for item in ocr_data_digits(preprocess(roi)):
            m = PATTERN.fullmatch(item["text"].replace(" ",""))
            if m: serial = m.group(1); break
        if not serial:
            txt = ocr_text_digits(preprocess(roi)).replace(" ","")
            m = PATTERN.search(txt)
            if m: serial = m.group(1)
        if not serial:
            if not auto:
                messagebox.showwarning("Не распозналось","В этой области не нашёл 2711xxxx или 20xxxxxx. Введите вручную ниже и нажмите «Применить к выбранному».")
            return

        # записать в таблицу + предложить имя
        vals = self.tree.item(sel[0],"values"); fname = vals[0]
        for r in self.rows:
            if r["old"]==fname:
                r["serial"]=serial; r["status"]="OK"
                date=self.date_var.get().strip()
                folder=Path(self.dir_var.get().strip())
                existing=set(os.listdir(folder)); seen={}
                r["new"] = propose(date, serial, Path(r["old"]).suffix.lower(), existing, seen)
                break
        self.refresh_tree()
        # показать красную рамку
        dbg = self.preview_cache.get(fname) or {}
        dbg["img"]=self.last_img; dbg["bbox"]=box; self.preview_cache[fname]=dbg
        self.show_image_with_overlays(Path(self.dir_var.get())/fname, dbg)
        self.status.set(f"Найден номер {serial} в выделенной области.")

    def apply_manual(self):
        sel = self.tree.selection()
        if not sel: messagebox.showinfo("Нет выбора","Выберите строку слева."); return
        serial = self.manual_var.get().strip()
        if not PATTERN.fullmatch(serial):
            messagebox.showwarning("Неверный формат","Введите 8 цифр: 2711xxxx или 20xxxxxx."); return
        vals = self.tree.item(sel[0],"values"); fname = vals[0]
        for r in self.rows:
            if r["old"]==fname: r["serial"]=serial; r["status"]="OK"; break
        date=self.date_var.get().strip(); folder=Path(self.dir_var.get().strip())
        existing=set(os.listdir(folder)); seen={}
        for r in self.rows:
            r["new"] = propose(date, r["serial"], Path(r["old"]).suffix.lower(), existing, seen) if r["serial"] else ""
        self.refresh_tree(); self.status.set("Ручной номер применён.")

def main():
    root=tk.Tk(); App(root); root.mainloop()

if __name__=="__main__":
    main()
