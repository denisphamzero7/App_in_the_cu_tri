import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from PIL import Image, ImageDraw, ImageFont, ImageTk
import pandas as pd
import os
import json
import platform
import win32api
import time
from copy import deepcopy

# ========================================================================================
# KHU VỰC 1: CẤU HÌNH & HẰNG SỐ
# ========================================================================================
CONFIG_FILE = "cau_hinh_v12_final.json"

COLORS = {
    "primary": "#3498db", "success": "#2ecc71", "danger": "#e74c3c",
    "warning": "#f39c12", "purple": "#9b59b6", "dark": "#2c3e50",
    "light": "#ecf0f1", "grey": "#bdc3c7", "text": "#2c3e50", "white": "#ffffff"
}

FONT_MAP = {
    "Arial": {"normal": "arial.ttf", "bold": "arialbd.ttf"},
    "Times New Roman": {"normal": "times.ttf", "bold": "timesbd.ttf"},
    "Calibri": {"normal": "calibri.ttf", "bold": "calibrib.ttf"}
}

# ========================================================================================
# KHU VỰC 2: UI COMPONENTS (BUTTON BO TRÒN HIỆN ĐẠI)
# ========================================================================================
class RoundedButton(tk.Canvas):
    """
    Nút bấm bo tròn sử dụng Canvas.
    Hỗ trợ: Bo góc (radius), Hiệu ứng Hover, Tự động resize.
    """
    def __init__(self, master, text="", command=None, width=120, height=35, radius=15, bg=COLORS["primary"], fg="white", font=("Segoe UI", 10, "bold"), **kwargs):
        # Lấy màu nền của cha để làm viền cho Canvas (tạo cảm giác trong suốt)
        try:
            parent_bg = master.cget("bg")
        except:
            parent_bg = COLORS["light"]

        super().__init__(master, width=width, height=height, bg=parent_bg, highlightthickness=0, **kwargs)
        
        self.command = command
        self.radius = radius
        self.text_str = text
        self.bg_color = bg
        self.fg_color = fg
        self.font = font
        
        # Tính toán màu hover (sáng hơn màu gốc)
        self.hover_color = self.adjust_color_lightness(bg, 1.15)
        # Nếu màu nền là xám/trắng thì chữ màu đen, ngược lại chữ trắng
        if bg in [COLORS["white"], COLORS["grey"], COLORS["light"]]:
            self.fg_color = "#2c3e50"
            self.hover_color = "#95a5a6"

        self.rect_id = None
        self.text_id = None

        # Bind sự kiện
        self.bind("<Configure>", self.resize) # Vẽ lại khi thay đổi kích thước
        self.bind("<Enter>", self.on_enter)
        self.bind("<Leave>", self.on_leave)
        self.bind("<Button-1>", self.on_click)

    def adjust_color_lightness(self, color_hex, factor):
        """Hàm phụ trợ làm sáng màu Hex"""
        try:
            r, g, b = int(color_hex[1:3], 16), int(color_hex[3:5], 16), int(color_hex[5:7], 16)
            r = min(int(r * factor), 255)
            g = min(int(g * factor), 255)
            b = min(int(b * factor), 255)
            return f"#{r:02x}{g:02x}{b:02x}"
        except:
            return color_hex

    def resize(self, event):
        """Vẽ lại nút khi kích thước thay đổi"""
        self.delete("all")
        w, h = event.width, event.height
        
        # Vẽ hình bo góc bằng kỹ thuật polygon smooth
        # Tạo toạ độ cho hình chữ nhật bo góc
        r = self.radius
        if r > h/2: r = h/2
        
        # Vẽ hình chữ nhật bo góc
        self.round_rect(0, 0, w, h, r, fill=self.bg_color, tags="bg")
        
        # Vẽ chữ ở giữa
        self.create_text(w/2, h/2, text=self.text_str, fill=self.fg_color, font=self.font, tags="text")

    def round_rect(self, x1, y1, x2, y2, r, **kwargs):
        points = (x1+r, y1, x1+r, y1, x2-r, y1, x2-r, y1, x2, y1, x2, y1+r, x2, y1+r, x2, y2-r, x2, y2-r, x2, y2, x2-r, y2, x2-r, y2, x1+r, y2, x1+r, y2, x1, y2, x1, y2-r, x1, y2-r, x1, y1+r, x1, y1+r, x1, y1)
        return self.create_polygon(points, **kwargs, smooth=True)

    def on_enter(self, event):
        self.itemconfig("bg", fill=self.hover_color)

    def on_leave(self, event):
        self.itemconfig("bg", fill=self.bg_color)

    def on_click(self, event):
        if self.command:
            self.command()

# ========================================================================================
# KHU VỰC 3: LỚP CHÍNH (MAIN APP)
# ========================================================================================
class VoterAppV12Final:
    def __init__(self, root):
        self.root = root
        self.root.title("HỆ THỐNG IN THẺ CỬ TRI - V12 FINAL (ROUNDED UI)")
        self.root.geometry("1600x950")
        self.root.configure(bg=COLORS["light"])
        
        # --- Cấu hình Style cho bảng (Treeview) ---
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Treeview.Heading", font=("Segoe UI", 10, "bold"), background=COLORS["dark"], foreground="white", relief="flat")
        style.map("Treeview.Heading", background=[('active', COLORS["primary"])])
        style.configure("Treeview", font=("Segoe UI", 10), rowheight=30, background="white", fieldbackground="white")
        style.map("Treeview", background=[('selected', COLORS["primary"])], foreground=[('selected', 'white')])

        # --- Khởi tạo các biến dữ liệu ---
        self.df = None; self.template_path = None; self.signature_folder = None
        self.pil_image = None; self.tk_image = None; self.tk_sig_ref = None 
        self.current_idx = 0; self.scale_factor = 1.0 
        self.drag_data = {"x": 0, "y": 0, "item": None}
        self.global_config = {}; self.custom_configs = {}
        self.chk_field_vars = {}; self.field_labels = {}; self.selected_field_name = None
        
        self.load_config_file()
        self.setup_ui_layout()

    def setup_ui_layout(self):
        main_container = tk.Frame(self.root, bg=COLORS["light"])
        main_container.pack(fill=tk.BOTH, expand=True)
        
        # Chia 3 cột
        left_panel = tk.Frame(main_container, bg=COLORS["light"], padx=15, pady=15)
        left_panel.place(relx=0, rely=0, relwidth=0.22, relheight=1)
        
        mid_panel = tk.Frame(main_container, bg="white", padx=10, pady=10, bd=1, relief="solid")
        mid_panel.place(relx=0.22, rely=0, relwidth=0.43, relheight=1)
        
        right_panel = tk.Frame(main_container, bg=COLORS["dark"])
        right_panel.place(relx=0.65, rely=0, relwidth=0.35, relheight=1)

        # --- PANEL TRÁI: Dữ liệu & Cấu hình ---
        tk.Label(left_panel, text="1. DỮ LIỆU ĐẦU VÀO", font=("Segoe UI", 12, "bold"), bg=COLORS["light"], fg=COLORS["dark"]).pack(anchor="w", pady=(0, 5))
        
        # Khu vực các nút chọn file
        btn_frame = tk.Frame(left_panel, bg=COLORS["light"])
        btn_frame.pack(fill=tk.X, pady=5)
        
        # SỬ DỤNG RoundedButton THAY VÌ ModernButton
        # Lưu ý: Canvas cần set chiều cao (height) cụ thể
        RoundedButton(btn_frame, text="📂 Chọn Ảnh Phôi", command=self.select_template, bg=COLORS["primary"]).pack(fill=tk.X, pady=3)
        RoundedButton(btn_frame, text="📊 Chọn File Excel", command=self.select_excel, bg=COLORS["success"]).pack(fill=tk.X, pady=3)
        RoundedButton(btn_frame, text="📂 Folder Chữ Ký (Auto)", command=self.select_signature_folder, bg=COLORS["purple"]).pack(fill=tk.X, pady=3)

        tk.Label(left_panel, text="2. CẤU HÌNH TRƯỜNG", font=("Segoe UI", 12, "bold"), bg=COLORS["light"], fg=COLORS["dark"]).pack(anchor="w", pady=(20, 5))
        list_container = tk.Frame(left_panel, bg="white", bd=1, relief="solid")
        list_container.pack(fill=tk.BOTH, expand=True, pady=5)
        self.canvas_list = tk.Canvas(list_container, bg="white", highlightthickness=0)
        scrollbar = ttk.Scrollbar(list_container, orient="vertical", command=self.canvas_list.yview)
        self.scrollable_frame = tk.Frame(self.canvas_list, bg="white")
        self.scrollable_frame.bind("<Configure>", lambda e: self.canvas_list.configure(scrollregion=self.canvas_list.bbox("all")))
        self.canvas_list.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas_list.configure(yscrollcommand=scrollbar.set)
        self.canvas_list.pack(side="left", fill="both", expand=True, padx=2, pady=2)
        scrollbar.pack(side="right", fill="y")

        # --- STYLE CONTROL PANEL ---
        style_frame = tk.LabelFrame(left_panel, text="3. TÙY CHỈNH STYLE", font=("Segoe UI", 11, "bold"), bg=COLORS["grey"], fg=COLORS["dark"], padx=5, pady=5)
        style_frame.pack(fill=tk.X, pady=10, side=tk.BOTTOM)

        self.edit_mode = tk.StringVar(value="global")
        tk.Radiobutton(style_frame, text="Chỉnh cho TẤT CẢ", variable=self.edit_mode, value="global", bg=COLORS["grey"], font=("Segoe UI", 9)).pack(anchor="w")
        tk.Radiobutton(style_frame, text="Chỉnh RIÊNG người này", variable=self.edit_mode, value="individual", bg=COLORS["grey"], font=("Segoe UI", 9, "bold"), fg="red").pack(anchor="w")
        tk.Frame(style_frame, height=1, bg="white").pack(fill=tk.X, pady=5)

        tk.Label(style_frame, text="Đang chọn:", bg=COLORS["grey"], font=("Segoe UI", 9)).pack(anchor="w")
        self.lbl_current_field = tk.Label(style_frame, text="(Chưa chọn)", fg="blue", bg=COLORS["grey"], font=("Segoe UI", 10, "bold"))
        self.lbl_current_field.pack(anchor="w", pady=(0, 5))
        
        self.btn_manual_sig = RoundedButton(style_frame, text="📂 Chọn ảnh chữ ký", command=self.pick_manual_signature, bg=COLORS["warning"], height=30)
        
        # Control Ảnh (Resize)
        self.row_img_size = tk.Frame(style_frame, bg=COLORS["grey"])
        tk.Label(self.row_img_size, text="Rộng:", bg=COLORS["grey"], font=("Segoe UI", 9)).pack(side=tk.LEFT, padx=(0,5))
        self.spin_img_w = tk.Spinbox(self.row_img_size, from_=1, to=1000, width=5, font=("Segoe UI", 10), command=self.apply_image_size)
        self.spin_img_w.pack(side=tk.LEFT)
        self.spin_img_w.bind("<Return>", self.apply_image_size)

        tk.Label(self.row_img_size, text="Cao:", bg=COLORS["grey"], font=("Segoe UI", 9)).pack(side=tk.LEFT, padx=(10,5))
        self.spin_img_h = tk.Spinbox(self.row_img_size, from_=1, to=1000, width=5, font=("Segoe UI", 10), command=self.apply_image_size)
        self.spin_img_h.pack(side=tk.LEFT)
        self.spin_img_h.bind("<Return>", self.apply_image_size)

        # Control Text (Font, Size, Color)
        self.row_text_font = tk.Frame(style_frame, bg=COLORS["grey"])
        self.combo_font = ttk.Combobox(self.row_text_font, values=["Arial", "Times New Roman"], width=13, state="readonly", font=("Segoe UI", 9))
        self.combo_font.pack(side=tk.LEFT)
        self.combo_font.bind("<<ComboboxSelected>>", self.apply_text_properties)
        self.spin_size = tk.Spinbox(self.row_text_font, from_=5, to=300, width=5, command=self.apply_text_properties, font=("Segoe UI", 10))
        self.spin_size.pack(side=tk.LEFT, padx=5)
        self.spin_size.bind("<Return>", self.apply_text_properties)
        
        self.row_text_style = tk.Frame(style_frame, bg=COLORS["grey"])
        self.chk_bold_var = tk.BooleanVar(); self.chk_upper_var = tk.BooleanVar()
        tk.Checkbutton(self.row_text_style, text="B", variable=self.chk_bold_var, bg=COLORS["grey"], font="Arial 9 bold", command=self.apply_text_properties).pack(side=tk.LEFT)
        tk.Checkbutton(self.row_text_style, text="AA", variable=self.chk_upper_var, bg=COLORS["grey"], command=self.apply_text_properties).pack(side=tk.LEFT)
        self.combo_color = ttk.Combobox(self.row_text_style, values=["Black", "Red", "Blue"], width=8, state="readonly", font=("Segoe UI", 9))
        self.combo_color.pack(side=tk.LEFT, padx=5)
        self.combo_color.bind("<<ComboboxSelected>>", self.apply_text_properties)
        
        RoundedButton(style_frame, text="↺ Reset Default", command=self.reset_current_custom, bg=COLORS["danger"], height=30).pack(fill=tk.X, pady=10)

        # --- PANEL GIỮA: List & Toolbar ---
        toolbar_frame = tk.Frame(mid_panel, bg="white")
        toolbar_frame.pack(side=tk.TOP, fill=tk.X, pady=(0, 10))
        
        left_tool = tk.Frame(toolbar_frame, bg="white")
        left_tool.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        # Nút nhỏ hơn 1 chút
        RoundedButton(left_tool, text="Chọn Tất Cả", command=self.select_all, bg=COLORS["primary"], width=90, height=30).pack(side=tk.LEFT, padx=(0, 5))
        RoundedButton(left_tool, text="Bỏ Chọn", command=self.deselect_all, bg=COLORS["grey"], width=90, height=30).pack(side=tk.LEFT)
        
        RoundedButton(toolbar_frame, text="🖨️ IN NGAY", command=self.start_batch_print, bg=COLORS["danger"], width=100, height=30).pack(side=tk.RIGHT)
        
        self.lbl_count = tk.Label(mid_panel, text="Đã chọn: 0", font=("Segoe UI", 10, "bold"), fg=COLORS["danger"], bg="white")
        self.lbl_count.pack(anchor="e", pady=(0, 5))

        self.tree = ttk.Treeview(mid_panel, columns=("stt", "name", "gender", "cccd", "area"), show="headings", selectmode="extended")
        self.tree.heading("stt", text="STT"); self.tree.column("stt", width=40, anchor="center")
        self.tree.heading("name", text="Họ Tên"); self.tree.column("name", width=180)
        self.tree.heading("gender", text="Giới tính"); self.tree.column("gender", width=60, anchor="center")
        self.tree.heading("cccd", text="CCCD"); self.tree.column("cccd", width=100)
        self.tree.heading("area", text="Khu vực"); self.tree.column("area", width=120)
        self.tree.tag_configure('custom', foreground='red', font=('Segoe UI', 10, 'bold'))
        
        ysb = ttk.Scrollbar(mid_panel, orient="vertical", command=self.tree.yview)
        xsb = ttk.Scrollbar(mid_panel, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=ysb.set, xscrollcommand=xsb.set)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        ysb.pack(side=tk.RIGHT, fill=tk.Y); xsb.pack(side=tk.BOTTOM, fill=tk.X)
        self.tree.bind("<<TreeviewSelect>>", self.on_tree_select_change)

        # --- PANEL PHẢI: Canvas ---
        self.canvas = tk.Canvas(right_panel, bg="#95a5a6", cursor="fleur")
        self.canvas.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        self.canvas.tag_bind("draggable", "<ButtonPress-1>", self.on_drag_start)
        self.canvas.tag_bind("draggable", "<B1-Motion>", self.on_drag_motion)
        self.canvas.tag_bind("draggable", "<ButtonRelease-1>", self.on_drag_end)

    # ========================================================================================
    # KHU VỰC 4: QUẢN LÝ CẤU HÌNH & DATA (LOGIC)
    # ========================================================================================
    def load_config_file(self):
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, "r", encoding="utf-8") as f: 
                    data = json.load(f)
                    self.global_config = data.get("global", {})
                    self.custom_configs = {int(k): v for k, v in data.get("custom", {}).items()}
            except: pass
        if "signature_img" not in self.global_config:
            self.global_config["signature_img"] = {"x": 300, "y": 300, "w": 150, "h": 80, "enable": True, "type": "image"}
    
    def save_config_file(self):
        data = {"global": self.global_config, "custom": self.custom_configs}
        with open(CONFIG_FILE, "w", encoding="utf-8") as f: json.dump(data, f, ensure_ascii=False, indent=4)

    def get_current_config(self, idx):
        config = deepcopy(self.global_config)
        if idx in self.custom_configs:
            for col, props in self.custom_configs[idx].items():
                if col in config: config[col].update(props)
                else: config[col] = props
        return config

    def update_config_value(self, col, key, value):
        mode = self.edit_mode.get()
        if mode == "global":
            if col in self.global_config:
                self.global_config[col][key] = value
                self.save_config_file()
        else:
            idx = self.current_idx
            if idx not in self.custom_configs: self.custom_configs[idx] = {}
            if col not in self.custom_configs[idx]:
                self.custom_configs[idx][col] = deepcopy(self.global_config.get(col, {}))
            self.custom_configs[idx][col][key] = value
            self.save_config_file()
            self.tree.item(idx, tags=('custom',))

    def reset_current_custom(self):
        idx = self.current_idx
        if idx in self.custom_configs:
            del self.custom_configs[idx]
            self.save_config_file()
            self.tree.item(idx, tags=())
            self.render_canvas()
            self.load_props(self.selected_field_name)
            messagebox.showinfo("Reset", "Đã xóa chỉnh sửa riêng.")

    # ========================================================================================
    # KHU VỰC 5: XỬ LÝ FILE INPUT (EXCEL, ẢNH)
    # ========================================================================================
    def select_template(self):
        path = filedialog.askopenfilename(filetypes=[("Image", "*.jpg;*.png")])
        if path: self.template_path = path; self.pil_image = Image.open(path); self.render_canvas()

    def select_signature_folder(self):
        folder = filedialog.askdirectory()
        if folder: self.signature_folder = folder; self.refresh_field_list(); messagebox.showinfo("OK", f"Đã chọn folder: {folder}")

    def select_excel(self):
        path = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx;*.xls")])
        if path:
            try:
                self.df = pd.read_excel(path).fillna(""); self.df.columns = self.df.columns.str.strip()
                self.refresh_field_list(); self.populate_treeview(); self.render_canvas()
            except Exception as e: messagebox.showerror("Lỗi", str(e))

    def refresh_field_list(self):
        for w in self.scrollable_frame.winfo_children(): w.destroy()
        self.chk_field_vars = {}; self.field_labels = {}
        cols = self.df.columns.tolist() if self.df is not None else []
        if "signature_img" not in cols: cols.append("signature_img")
        
        for col in cols:
            if col not in self.global_config:
                self.global_config[col] = {"x": 50, "y": 50, "size": 30, "enable": False, "font": "Arial", "color": "Black", "type": "text"}
            
            row_fr = tk.Frame(self.scrollable_frame, bg="white")
            row_fr.pack(fill="x", pady=2)
            
            var = tk.BooleanVar(value=self.global_config[col].get("enable", False))
            self.chk_field_vars[col] = var
            chk = tk.Checkbutton(row_fr, variable=var, bg="white", command=lambda c=col: self.on_field_toggle(c))
            chk.pack(side="left")
            
            lbl_text = "📷 ẢNH CHỮ KÝ" if col == "signature_img" else col
            lbl = tk.Label(row_fr, text=lbl_text, bg="white", anchor="w", font=("Segoe UI", 10), cursor="hand2")
            lbl.pack(side="left", fill="x", expand=True)
            lbl.bind("<Button-1>", lambda e, c=col: self.load_props(c))
            self.field_labels[col] = lbl

    def on_field_toggle(self, col):
        self.global_config[col]["enable"] = self.chk_field_vars[col].get()
        self.save_config_file(); self.load_props(col); self.render_canvas()

    def find_column_insensitive(self, keywords):
        if self.df is None: return None
        for col in self.df.columns:
            for kw in keywords:
                if kw.lower() in col.lower(): return col
        return None

    def populate_treeview(self):
        for i in self.tree.get_children(): self.tree.delete(i)
        col_name = self.find_column_insensitive(["Họ tên", "Họ và tên", "Name"])
        col_gender = self.find_column_insensitive(["Giới tính", "Gender"])
        col_cccd = self.find_column_insensitive(["CCCD", "CMND"])
        col_area = self.find_column_insensitive(["Khu vực", "Thôn"])
        
        if not col_name and len(self.df.columns) > 1: col_name = self.df.columns[1]
        
        for i, row in self.df.iterrows():
            tag = ('custom',) if i in self.custom_configs else ()
            vals = (i+1, row.get(col_name,""), row.get(col_gender,""), row.get(col_cccd,""), row.get(col_area,""))
            self.tree.insert("", "end", iid=i, values=vals, tags=tag)
        self.select_all()

    def select_all(self): self.tree.selection_set(self.tree.get_children()); self.update_count_label()
    def deselect_all(self): self.tree.selection_set([]); self.update_count_label()
    
    def on_tree_select_change(self, event):
        sel = self.tree.selection()
        if sel: 
            self.current_idx = int(sel[0]); self.render_canvas(); 
            if self.selected_field_name: self.load_props(self.selected_field_name)
        self.update_count_label()
        
    def update_count_label(self): self.lbl_count.config(text=f"Sẽ in: {len(self.tree.selection())} người")

    # ========================================================================================
    # KHU VỰC 6: XỬ LÝ ẢNH & FONT
    # ========================================================================================
    def pick_manual_signature(self):
        path = filedialog.askopenfilename(filetypes=[("Image", "*.png;*.jpg;*.jpeg")])
        if not path: return
        idx = self.current_idx
        if idx not in self.custom_configs: self.custom_configs[idx] = {}
        if "signature_img" not in self.custom_configs[idx]:
            self.custom_configs[idx]["signature_img"] = deepcopy(self.global_config.get("signature_img", {}))
        self.custom_configs[idx]["signature_img"]["path"] = path
        self.custom_configs[idx]["signature_img"]["enable"] = True
        if "signature_img" in self.chk_field_vars: self.chk_field_vars["signature_img"].set(True)
        self.save_config_file(); self.tree.item(idx, tags=('custom',)); self.render_canvas()

    def get_signature_image(self, idx):
        # 1. Tìm cấu hình thủ công
        if idx in self.custom_configs and "signature_img" in self.custom_configs[idx]:
            p = self.custom_configs[idx]["signature_img"].get("path")
            if p and os.path.exists(p): return Image.open(p).convert("RGBA")
        
        # 2. Tìm tự động trong folder
        if self.signature_folder:
            row = self.df.iloc[idx]
            cccd = str(row.get(self.find_column_insensitive(["CCCD", "CMND"]) or "", "")).strip()
            names = [cccd, str(idx+1)] if cccd else [str(idx+1)]
            for n in names:
                for ext in [".png", ".jpg", ".jpeg"]:
                    p = os.path.join(self.signature_folder, n + ext)
                    if os.path.exists(p): return Image.open(p).convert("RGBA")
        return None

    def get_pil_font(self, font_name, size, is_bold):
        if platform.system() == "Windows":
            style = "bold" if is_bold else "normal"
            f_map = FONT_MAP.get(font_name, FONT_MAP["Arial"]).get(style, "arial.ttf")
            path = os.path.join(os.environ["WINDIR"], "Fonts", f_map)
            if os.path.exists(path): return ImageFont.truetype(path, size)
        return ImageFont.load_default()

    # ========================================================================================
    # KHU VỰC 7: RENDER ENGINE (VẼ ẢNH)
    # ========================================================================================
    def render_one_image(self, idx):
        if not self.template_path: return None
        row = self.df.iloc[idx]
        img = Image.open(self.template_path).convert("RGB")
        draw = ImageDraw.Draw(img)
        final_config = self.get_current_config(idx)
        
        for col, cfg in final_config.items():
            if not cfg.get("enable", False): continue
            
            if col == "signature_img":
                sig = self.get_signature_image(idx)
                if sig:
                    w, h = cfg.get("w", 150), cfg.get("h", 80)
                    sig = sig.resize((w, h), Image.Resampling.LANCZOS)
                    img.paste(sig, (int(cfg["x"] - w/2), int(cfg["y"] - h/2)), sig)
                continue

            val = str(row.get(col, "")).replace("nan", "")
            if "00:00:00" in val: val = val.split(" ")[0]
            if cfg.get("upper", False): val = val.upper()
            font = self.get_pil_font(cfg.get("font", "Arial"), cfg.get("size", 30), cfg.get("bold", False))
            draw.text((cfg["x"], cfg["y"]), val, font=font, fill=cfg.get("color", "black"), anchor="mm")
            
        return img

    def render_canvas(self):
        if not self.template_path: return
        self.canvas.delete("all")
        
        cw, ch = self.canvas.winfo_width(), self.canvas.winfo_height()
        if cw < 50: cw, ch = 800, 600
        iw, ih = self.pil_image.size
        self.scale_factor = min(cw/iw, ch/ih) * 0.95
        nw, nh = int(iw*self.scale_factor), int(ih*self.scale_factor)
        
        self.tk_image = ImageTk.PhotoImage(self.pil_image.resize((nw, nh), Image.Resampling.LANCZOS))
        cx, cy = cw//2, ch//2
        self.img_origin_x = cx - nw//2; self.img_origin_y = cy - nh//2
        self.canvas.create_image(cx, cy, image=self.tk_image, anchor=tk.CENTER)
        
        if self.df is not None:
            row = self.df.iloc[self.current_idx]
            final_config = self.get_current_config(self.current_idx)
            
            text_items = []; signature_item = None
            for col, cfg in final_config.items():
                if not cfg.get("enable", False): continue
                if col == "signature_img": signature_item = (col, cfg)
                else: text_items.append((col, cfg))
            
            for col, cfg in text_items:
                sx = self.img_origin_x + cfg["x"]*self.scale_factor
                sy = self.img_origin_y + cfg["y"]*self.scale_factor
                val = str(row.get(col, "")).replace("nan", "")
                if "00:00:00" in val: val = val.split(" ")[0]
                if cfg.get("upper", False): val = val.upper()
                f_sz = int(cfg.get("size", 30)*self.scale_factor)
                tk_font = (cfg.get("font", "Arial"), -f_sz, "bold" if cfg.get("bold", False) else "normal")
                
                clr = "red" if (self.edit_mode.get()=="individual" and self.current_idx in self.custom_configs and col in self.custom_configs[self.current_idx]) else cfg.get("color", "black")
                self.canvas.create_text(sx, sy, text=val, font=tk_font, fill=clr, anchor="center", tags=("draggable", f"col:{col}"))

            if signature_item:
                col, cfg = signature_item
                sx = self.img_origin_x + cfg["x"]*self.scale_factor
                sy = self.img_origin_y + cfg["y"]*self.scale_factor
                w = int(cfg.get("w", 150) * self.scale_factor); h = int(cfg.get("h", 80) * self.scale_factor)
                sig = self.get_signature_image(self.current_idx)
                
                if sig:
                    self.tk_sig_ref = ImageTk.PhotoImage(sig.resize((w, h), Image.Resampling.LANCZOS))
                    self.canvas.create_image(sx, sy, image=self.tk_sig_ref, anchor="center", tags=("draggable", f"col:{col}"))
                    self.canvas.create_rectangle(sx - w/2, sy - h/2, sx + w/2, sy + h/2, outline="blue", dash=(2, 4), tags=("draggable", f"col:{col}"))
                else:
                    self.canvas.create_rectangle(sx - w/2, sy - h/2, sx + w/2, sy + h/2, outline="red", width=2, dash=(5, 2), tags=("draggable", f"col:{col}"))
                    self.canvas.create_text(sx, sy, text="CHỖ ĐỂ ẢNH", fill="red", font=("Segoe UI", 8, "bold"), justify="center", tags=("draggable", f"col:{col}"))

    # ========================================================================================
    # KHU VỰC 8: KÉO THẢ (DRAG & DROP)
    # ========================================================================================
    def on_drag_start(self, e):
        items = self.canvas.find_closest(e.x, e.y)
        if items:
            tags = self.canvas.gettags(items[0])
            if "draggable" in tags:
                self.drag_data = {"x": e.x, "y": e.y, "item": items[0]}
                for t in tags:
                    if t.startswith("col:"): self.load_props(t.split(":")[1]); break
    
    def on_drag_motion(self, e):
        if self.drag_data["item"]:
            self.canvas.move(self.drag_data["item"], e.x - self.drag_data["x"], e.y - self.drag_data["y"])
            self.drag_data.update({"x": e.x, "y": e.y})
            
    def on_drag_end(self, e):
        if self.drag_data["item"]:
            coords = self.canvas.coords(self.drag_data["item"])
            if len(coords) == 2: cx, cy = coords
            elif len(coords) == 4: cx, cy = (coords[0]+coords[2])/2, (coords[1]+coords[3])/2
            else: cx, cy = 0, 0
            
            if self.selected_field_name:
                self.update_config_value(self.selected_field_name, "x", int((cx-self.img_origin_x)/self.scale_factor))
                self.update_config_value(self.selected_field_name, "y", int((cy-self.img_origin_y)/self.scale_factor))
                self.render_canvas()
        self.drag_data["item"] = None

    # ========================================================================================
    # KHU VỰC 9: XỬ LÝ SỰ KIỆN GIAO DIỆN
    # ========================================================================================
    def load_props(self, col):
        self.selected_field_name = col
        self.lbl_current_field.config(text=col)
        
        for f, lbl in self.field_labels.items():
            lbl.config(bg="#dff9fb" if f==col else "white", fg="blue" if f==col else "black", font=("Segoe UI", 10, "bold" if f==col else "normal"))
        
        cfg = self.get_current_config(self.current_idx).get(col, {})
        
        if col == "signature_img":
            self.btn_manual_sig.pack(side=tk.TOP, fill=tk.X, pady=5)
            self.row_img_size.pack(fill=tk.X, pady=2)
            self.row_text_font.pack_forget(); self.row_text_style.pack_forget()
            
            self.spin_img_w.delete(0, tk.END); self.spin_img_w.insert(0, cfg.get("w", 150))
            self.spin_img_h.delete(0, tk.END); self.spin_img_h.insert(0, cfg.get("h", 80))
        else:
            self.btn_manual_sig.pack_forget(); self.row_img_size.pack_forget()
            self.row_text_font.pack(fill=tk.X, pady=2); self.row_text_style.pack(fill=tk.X, pady=2)
            
            self.combo_font.set(cfg.get("font", "Arial"))
            self.spin_size.delete(0, tk.END); self.spin_size.insert(0, cfg.get("size", 30))
            self.chk_bold_var.set(cfg.get("bold", False)); self.chk_upper_var.set(cfg.get("upper", False))
            self.combo_color.set(cfg.get("color", "Black"))

    def apply_text_properties(self, e=None):
        if self.selected_field_name and self.selected_field_name != "signature_img":
            self.update_config_value(self.selected_field_name, "font", self.combo_font.get())
            try: self.update_config_value(self.selected_field_name, "size", int(self.spin_size.get()))
            except: pass
            self.update_config_value(self.selected_field_name, "bold", self.chk_bold_var.get())
            self.update_config_value(self.selected_field_name, "upper", self.chk_upper_var.get())
            self.update_config_value(self.selected_field_name, "color", self.combo_color.get())
            self.render_canvas()

    def apply_image_size(self, e=None):
        if self.selected_field_name == "signature_img":
            try:
                w, h = int(self.spin_img_w.get()), int(self.spin_img_h.get())
                if w < 1: w = 1
                if h < 1: h = 1
                self.update_config_value("signature_img", "w", w)
                self.update_config_value("signature_img", "h", h)
                self.render_canvas()
            except ValueError: pass

    def start_batch_print(self):
        sel = self.tree.selection()
        if not sel: return messagebox.showwarning("Chú ý", "Chưa chọn người!")
        if not messagebox.askyesno("In", f"In {len(sel)} thẻ?"): return
        
        td = "temp_batch_final"; 
        if not os.path.exists(td): os.makedirs(td)
        
        for iid in sel:
            idx = int(iid); img = self.render_one_image(idx)
            fn = os.path.join(td, f"job_{idx}.pdf"); img.save(fn)
            try: win32api.ShellExecute(0, "print", fn, None, ".", 0); time.sleep(1.5)
            except Exception as e: print(e)
        messagebox.showinfo("Xong", "Đã gửi lệnh in.")

if __name__ == "__main__":
    root = tk.Tk()
    app = VoterAppV12Final(root)
    root.mainloop()