# -*- coding: utf-8 -*-
import os, re, json, sys, threading, queue
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

APP_TITLE = "Переименование отсканированных файлов — Ella Renamer v3.2 (фоновые задачи)"
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
        return pytesseract.image_to_string(pil, lang="eng", config="--psm 6 -c tessedit_char_whitelist=0123456789")
    except Exception:
        return ""

def ocr_data_digits(pil)->List[Dict]:
    if pytesseract is None: return []
    try:
        d = pytesseract.image_to_data(pil, lang="eng", config="--psm 6 -c tessedit_char_whitelist=0123456789", output_type=pytesseract.Output.DICT)
        out=[]; n=len(d["text"])
        for i in range(n):
            txt=(d["text"][i] or "").strip().replace(" ","")
            if not txt: continue
            out.append({"text": txt, "left": d["left"][i], "top": d["top"][i], "width": d["width"][i], "height": d["height"][i]})
        return out
    except Exception:
        return []

def search_regions(img):
    w,h=img.size
    return [
        ((int(w*0.55), 0, int(w*0.995), int(h*0.40)), "правый верх"),
        ((0, 0, int(w*0.22), h), "левая вертикаль"),
    ]

def orientations(img, use_rot=True):
    lst=[(0, img)]
    if not use_rot: return lst
    try:
        lst += [(90, img.rotate(90,expand=True)), (180, img.rotate(180,expand=True)), (270, img.rotate(270,expand=True))]
    except Exception: pass
    return lst

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

def extract_number_debug(path: Path, cfg: dict):
    if Image is None:
        return {"serial": None, "angle":0, "bbox": None, "source": None, "img": None}
    try:
        base = pil_open(path)
    except Exception:
        return {"serial": None, "angle":0, "bbox": None, "source": None, "img": None}

    # Ускорение: ограничим максимальную ширину изображения (OCR хватает)
    max_w = 2400
    if base.size[0] > max_w:
        h = int(base.size[1] * (max_w / base.size[0]))
        base = base.resize((max_w, h))

    for angle, img in orientations(base, cfg.get("use_rotations", True)):
        # 1) По якорям
        for bbox_reg, lbl in search_regions(img):
            region = img.crop(bbox_reg)
            serial, box, why = find_near_anchor(region)
            if serial:
                if box:
                    L = bbox_reg[0] + box[0]; T = bbox_reg[1] + box[1]
                    R = bbox_reg[0] + box[2]; B = bbox_reg[1] + box[3]
                    gb = (L,T,R,B)
                else:
                    gb = bbox_reg
                return {"serial": serial, "angle": angle, "bbox": gb, "source": f"{lbl}, {why}", "img": img}
        # 2) Фолбэк: только в зонах
        for bbox_reg, lbl in search_regions(img):
            region = img.crop(bbox_reg)
            for item in ocr_data_digits(preprocess(region)):
                m = PATTERN.fullmatch(item["text"].replace(" ", ""))
                if m:
                    L = bbox_reg[0] + item["left"]; T = bbox_reg[1] + item["top"]
                    R = L + item["width"]; B = T + item["height"]
                    return {"serial": m.group(1), "angle": angle, "bbox": (L,T,R,B), "source": f"{lbl}, без якоря", "img": img}
    return {"serial": None, "angle":0, "bbox": None, "source": None, "img": base}

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
        master.title(APP_TITLE); master.geometry("1180x660")
        self.cfg = load_config(); ensure_tess(self.cfg)
        self.dir_var=tk.StringVar(); self.date_var=tk.StringVar()
        self.tess_var=tk.StringVar(value=self.cfg.get("tesseract_path",""))
        self.status=tk.StringVar(value="Готово")

        # worker
        self.worker=None
        self.cancel_flag=False
        self.q = queue.Queue()
        self.total = 0
        self.done = 0

        self.rows=[]; self.preview_cache={}
        self._photo=None
        self.build()

        master.bind("<F5>", lambda e: self.start_preview())
        master.bind("<F6>", lambda e: self.apply())
        master.bind("<Control-o>", lambda e: self.choose_dir())

    def build(self):
        top=ttk.Frame(self); top.pack(fill="x", padx=10, pady=8)
        ttk.Label(top,text="Папка с файлами:").grid(row=0,column=0,sticky="w")
        ttk.Entry(top,textvariable=self.dir_var,width=70).grid(row=0,column=1,sticky="we",padx=6)
        ttk.Button(top,text="Выбрать...",command=self.choose_dir).grid(row=0,column=2,padx=4)
        ttk.Label(top,text="Дата (ДД.ММ.ГГГГ):").grid(row=1,column=0,sticky="w")
        ttk.Entry(top,textvariable=self.date_var,width=20).grid(row=1,column=1,sticky="w")
        ttk.Button(top,text="Из имени папки",command=self.fill_date).grid(row=1,column=2,padx=4,sticky="w")
        ttk.Label(top,text="Путь к tesseract.exe:").grid(row=2,column=0,sticky="w")
        ttk.Entry(top,textvariable=self.tess_var,width=70).grid(row=2,column=1,sticky="we",padx=6)
        ttk.Button(top,text="Найти...",command=self.choose_tess).grid(row=2,column=2,padx=4)
        ttk.Button(top,text="Сохранить",command=self.save_cfg).grid(row=2,column=3,padx=4)

        # панель действий
        toolbar=ttk.Frame(top); toolbar.grid(row=3,column=0,columnspan=4,sticky="we",pady=(8,4))
        self.btn_preview = ttk.Button(toolbar,text="Предпросмотр (F5)",command=self.start_preview)
        self.btn_preview.pack(side="left",padx=4)
        self.btn_cancel = ttk.Button(toolbar,text="Отмена",command=self.cancel_preview,state="disabled")
        self.btn_cancel.pack(side="left",padx=4)
        self.btn_rename = ttk.Button(toolbar,text="Переименовать (F6)",command=self.apply)
        self.btn_rename.pack(side="left",padx=4)
        ttk.Button(toolbar,text="Открыть папку",command=self.open_folder).pack(side="left",padx=4)

        # прогресс
        prog=ttk.Frame(top); prog.grid(row=4,column=0,columnspan=4,sticky="we")
        self.pb = ttk.Progressbar(prog,mode="determinate",maximum=100)
        self.pb.pack(fill="x",expand=True)

        # ручной ввод
        manual=ttk.Frame(top); manual.grid(row=5,column=0,columnspan=4,sticky="we",pady=(6,0))
        ttk.Label(manual, text="Ручной ввод номера (для выбранного файла):").pack(side="left")
        self.manual_var = tk.StringVar()
        ttk.Entry(manual, textvariable=self.manual_var, width=20).pack(side="left", padx=6)
        ttk.Button(manual, text="Применить к выбранному", command=self.apply_manual).pack(side="left")

        # основная область
        main=ttk.Frame(self); main.pack(fill="both", expand=True, padx=10, pady=6)
        left=ttk.Frame(main); left.pack(side="left", fill="both", expand=True)
        self.tree=ttk.Treeview(left,columns=("old","serial","new","status"),show="headings",height=18)
        for k,t,w in [("old","Старое имя",360),("serial","Номер (8 цифр)",120),("new","Новое имя",300),("status","Статус",100)]:
            self.tree.heading(k,text=t); self.tree.column(k,width=w,anchor="center" if k!="old" else "w")
        self.tree.pack(side="left",fill="both",expand=True)
        vsb=ttk.Scrollbar(left,orient="vertical",command=self.tree.yview); vsb.pack(side="right",fill="y")
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.bind("<<TreeviewSelect>>", self.on_select_row)

        right=ttk.Frame(main); right.pack(side="left", fill="both", expand=False, padx=(10,0))
        self.canvas = tk.Canvas(right, width=520, height=520, bg="#f6f6f6", highlightthickness=1, highlightbackground="#cccccc")
        self.canvas.pack(fill="both", expand=False)

        ttk.Label(self,textvariable=self.status,anchor="w").pack(fill="x", padx=10, pady=(0,8))

        top.columnconfigure(1, weight=1)

    # ---------- UI actions ----------
    def choose_dir(self):
        p=filedialog.askdirectory(title="Выберите папку с файлами")
        if p: self.dir_var.set(p); self.fill_date()

    def choose_tess(self):
        p=filedialog.askopenfilename(title="Выберите tesseract.exe",filetypes=[("tesseract.exe","tesseract.exe"),("Все файлы","*.*")])
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

    # ---------- threaded preview ----------
    def start_preview(self):
        if self.worker and self.worker.is_alive():
            return
        self.tree.delete(*self.tree.get_children()); self.preview_cache.clear()
        self.pb["value"]=0; self.status.set("Подготовка...")
        fs=self.files()
        if not fs:
            self.status.set("Файлы не найдены."); return
        date=self.date_var.get().strip()
        if not re.fullmatch(r"\d{2}\.\d{2}\.\d{4}", date):
            messagebox.showwarning("Дата","Введите ДД.ММ.ГГГГ или нажмите «Из имени папки»."); self.status.set("Ожидает даты."); return

        ensure_tess(load_config())
        self.rows=[]; self.cancel_flag=False; self.total=len(fs); self.done=0
        self.btn_preview.config(state="disabled"); self.btn_cancel.config(state="normal"); self.btn_rename.config(state="disabled")
        # пустим воркер
        self.q = queue.Queue()
        self.worker = threading.Thread(target=self._worker_preview, args=(fs, date), daemon=True)
        self.worker.start()
        self.after(100, self._poll_queue)

    def cancel_preview(self):
        self.cancel_flag=True
        self.status.set("Отмена...")

    def _worker_preview(self, files: List[Path], date: str):
        folder = Path(self.dir_var.get().strip())
        existing=set(os.listdir(folder)); seen={}
        for p in files:
            if self.cancel_flag: break
            try:
                dbg = extract_number_debug(p, {"use_rotations": True}) if Image else {"serial": None, "img": None, "bbox": None, "angle": 0, "source": None}
                serial = dbg["serial"]
                if serial:
                    st="OK"; newname=propose(date,serial,p.suffix.lower(),existing,seen)
                else:
                    st="НЕ НАЙДЕНО"; newname=""
                row={"old":p.name,"serial":serial or "","new":newname,"status":st}
                self.q.put(("row", row, p.name, dbg))
            except Exception as e:
                row={"old":p.name,"serial":"","new":"","status":"ОШИБКА"}
                self.q.put(("row", row, p.name, {"img": None, "bbox": None, "source": str(e), "angle": 0}))
            self.q.put(("progress", None, None, None))
        self.q.put(("done", None, None, None))

    def _poll_queue(self):
        changed=False
        while True:
            try:
                kind, row, fname, dbg = self.q.get_nowait()
            except queue.Empty:
                break
            if kind=="row":
                self.rows.append(row)
                self.tree.insert("", "end", values=(row["old"],row["serial"],row["new"],row["status"]))
                self.preview_cache[fname]=dbg
                changed=True
            elif kind=="progress":
                self.done += 1
                self.pb["value"] = int(100*self.done/max(1,self.total))
                self.status.set(f"Обработка {self.done}/{self.total}...")
            elif kind=="done":
                self.btn_preview.config(state="normal"); self.btn_cancel.config(state="disabled"); self.btn_rename.config(state="normal")
                if self.cancel_flag:
                    self.status.set(f"Отменено. Готово {self.done} из {self.total}.")
                else:
                    self.status.set("Готово. Выберите файл — справа подсветится зона.")
                return
        # продолжаем опрос
        self.after(100, self._poll_queue)

    # ---------- остальные действия ----------
    def apply(self):
        if self.worker and self.worker.is_alive():
            messagebox.showinfo("Подождите","Предпросмотр ещё выполняется или отменяется.")
            return
        if not self.rows:
            messagebox.showinfo("Нет данных","Сначала выполните предпросмотр.")
            return
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

    def on_select_row(self, event=None):
        sel = self.tree.selection()
        if not sel: return
        vals = self.tree.item(sel[0],"values")
        if not vals: return
        filename = vals[0]
        dbg = self.preview_cache.get(filename)
        folder=Path(self.dir_var.get().strip())
        path = folder/filename
        self.show_image_with_bbox(path, dbg)

    def show_image_with_bbox(self, path: Path, dbg: dict):
        self.canvas.delete("all"); self._photo=None
        if Image is None: return
        try:
            img = dbg["img"] if (dbg and dbg.get("img") is not None) else pil_open(path)
        except Exception:
            return
        cw = self.canvas.winfo_width() or 520
        ch = self.canvas.winfo_height() or 520
        w,h = img.size
        scale = min(cw/w, ch/h)
        nw, nh = int(w*scale), int(h*scale)
        disp = img.resize((nw,nh))
        self._photo = ImageTk.PhotoImage(disp)
        x0 = (cw-nw)//2; y0 = (ch-nh)//2
        self.canvas.create_image(x0, y0, image=self._photo, anchor="nw")
        bbox = dbg.get("bbox") if dbg else None
        if bbox:
            l,t,r,b = bbox
            l = int(l*scale)+x0; r=int(r*scale)+x0; t=int(t*scale)+y0; b=int(b*scale)+y0
            self.canvas.create_rectangle(l,t,r,b, outline="red", width=3)
        cap = ""
        if dbg and dbg.get("source"): cap += dbg["source"]
        if dbg and dbg.get("angle"): cap += f" | поворот {dbg['angle']}°"
        if cap: self.canvas.create_text(10, ch-10, text=cap, anchor="sw", fill="#222222")

    def apply_manual(self):
        sel = self.tree.selection()
        if not sel: messagebox.showinfo("Нет выбора","Выберите строку слева."); return
        serial = self.manual_var.get().strip()
        if not PATTERN.fullmatch(serial): messagebox.showwarning("Неверный формат","Введите 8 цифр: 2711xxxx или 20xxxxxx."); return
        vals = self.tree.item(sel[0],"values"); fname = vals[0]
        for r in self.rows:
            if r["old"]==fname: r["serial"]=serial; r["status"]="OK"; break
        date=self.date_var.get
