import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
import sys
import json
import threading
from io import BytesIO
import math

# --- 依赖库检查 ---
def check_imports():
    try:
        global fitz, Image, ImageTk, ImageEnhance
        import fitz  # PyMuPDF
        from PIL import Image, ImageTk, ImageEnhance
        return True
    except ImportError as e:
        try:
            root = tk.Tk()
            root.withdraw()
            messagebox.showerror("缺少必要库", f"无法启动:\n{e}\n\n请运行: pip install pymupdf pillow")
            root.destroy()
        except: pass
        return False

if not check_imports():
    sys.exit(1)

CONFIG_FILE = "watermark_settings.json"

class AdvancedWatermarkApp:
    def __init__(self, root):
        self.root = root
        # 【修改点1】更新窗口标题中的版本号
        self.root.title("可视化 PDF 水印工具 v 1.0.6")
        self.root.geometry("1200x900")
        
        # --- 核心数据 ---
        self.pdf_files = []
        self.current_doc = None
        self.current_page_idx = 0
        self.total_pages = 0
        
        self.current_pdf_img = None 
        self.current_wm_img = None
        self.pt_to_canvas_scale = 1.0    
        
        # 视觉坐标 (Points): 基于 PDF 视觉左下角
        self.wm_x = 0 
        self.wm_y = 0 
        self.vis_pdf_w = 0
        self.vis_pdf_h = 0
        
        # --- 变量 ---
        self.watermark_path = tk.StringVar()
        self.scale_var = tk.DoubleVar(value=1.0)
        self.opacity_var = tk.DoubleVar(value=0.5)
        self.angle_var = tk.DoubleVar(value=0)
        self.preview_zoom_var = tk.DoubleVar(value=1.0)
        self.range_mode_var = tk.StringVar(value="全部页面")
        self.custom_range_var = tk.StringVar(value="")
        
        self.status_var = tk.StringVar(value="准备就绪")
        self.page_info_var = tk.StringVar(value="0 / 0")
        self.is_processing = False

        self.load_config()
        self.setup_ui()
        
        if self.watermark_path.get() and os.path.exists(self.watermark_path.get()):
            try:
                self.current_wm_img = Image.open(self.watermark_path.get()).convert("RGBA")
            except: pass

        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

    def load_config(self):
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                    data = json.load(f)
                    self.watermark_path.set(data.get("watermark_path", ""))
                    self.scale_var.set(data.get("scale", 1.0))
                    self.opacity_var.set(data.get("opacity", 0.5))
                    self.angle_var.set(data.get("angle", 0))
                    self.range_mode_var.set(data.get("range_mode", "全部页面"))
                    self.custom_range_var.set(data.get("custom_range", ""))
            except: pass

    def save_config(self):
        data = {
            "watermark_path": self.watermark_path.get(),
            "scale": self.scale_var.get(),
            "opacity": self.opacity_var.get(),
            "angle": self.angle_var.get(),
            "range_mode": self.range_mode_var.get(),
            "custom_range": self.custom_range_var.get()
        }
        try:
            with open(CONFIG_FILE, "w", encoding="utf-8") as f:
                json.dump(data, f)
        except: pass

    def on_closing(self):
        if self.current_doc: self.current_doc.close()
        self.save_config()
        self.root.destroy()

    def setup_ui(self):
        # 左侧控制面板
        left_panel = tk.Frame(self.root, width=380, padx=15, pady=10)
        left_panel.pack(side="left", fill="y")
        
        # 右侧预览区域
        right_panel = tk.Frame(self.root, bg="#dcdcdc", padx=2, pady=2)
        right_panel.pack(side="right", fill="both", expand=True)

        # 1. 文件选择
        lf_files = tk.LabelFrame(left_panel, text="1. 文件选择", padx=5, pady=5)
        lf_files.pack(fill="x", pady=5)
        tk.Button(lf_files, text="选择 PDF (支持多选)", command=self.select_pdfs).pack(fill="x")
        self.lbl_pdf_info = tk.Label(lf_files, text="未加载", fg="gray")
        self.lbl_pdf_info.pack()

        # 2. 水印图片
        lf_img = tk.LabelFrame(left_panel, text="2. 水印图片", padx=5, pady=5)
        lf_img.pack(fill="x", pady=5)
        tk.Button(lf_img, text="选择图片", command=self.select_watermark).pack(fill="x")
        tk.Label(lf_img, textvariable=self.watermark_path, wraplength=200, fg="gray").pack()

        # 3. 水印样式
        lf_style = tk.LabelFrame(left_panel, text="3. 水印样式", padx=5, pady=5)
        lf_style.pack(fill="x", pady=5)
        tk.Label(lf_style, text="水印大小:").pack(anchor="w")
        tk.Scale(lf_style, from_=0.01, to=3.0, resolution=0.01, orient="horizontal", 
                 variable=self.scale_var, command=self.update_preview).pack(fill="x")
        tk.Label(lf_style, text="透明度:").pack(anchor="w")
        tk.Scale(lf_style, from_=0.1, to=1.0, resolution=0.1, orient="horizontal", 
                 variable=self.opacity_var, command=self.update_preview).pack(fill="x")
        tk.Label(lf_style, text="旋转角度:").pack(anchor="w")
        tk.Scale(lf_style, from_=0, to=360, resolution=5, orient="horizontal", 
                 variable=self.angle_var, command=self.update_preview).pack(fill="x")

        # 4. 位置控制
        lf_pos = tk.LabelFrame(left_panel, text="4. 位置控制 (可拖拽预览图)", padx=5, pady=5)
        lf_pos.pack(fill="x", pady=5)
        btn_frame = tk.Frame(lf_pos)
        btn_frame.pack(fill="x")
        tk.Button(btn_frame, text="↖ 左上角", command=self.set_pos_top_left).pack(side="left", expand=True)
        tk.Button(btn_frame, text="✛ 居中", command=self.set_pos_center).pack(side="left", expand=True)
        self.lbl_coords = tk.Label(lf_pos, text="X: 0, Y: 0")
        self.lbl_coords.pack(pady=5)

        # 5. 应用范围
        lf_range = tk.LabelFrame(left_panel, text="5. 应用范围", padx=5, pady=5)
        lf_range.pack(fill="x", pady=5)
        cb_range = ttk.Combobox(lf_range, values=["全部页面", "奇数页", "偶数页", "指定页面"], state="readonly", textvariable=self.range_mode_var)
        cb_range.pack(fill="x", pady=2)
        cb_range.bind("<<ComboboxSelected>>", self.toggle_range_entry)
        self.entry_range = tk.Entry(lf_range, textvariable=self.custom_range_var, state="disabled")
        self.entry_range.pack(fill="x", pady=2)

        self.progress = ttk.Progressbar(left_panel, orient="horizontal", length=100, mode="determinate")
        self.progress.pack(fill="x", pady=(20, 5))
        
        self.btn_run = tk.Button(left_panel, text="开始批量处理", bg="#d4edda", height=2, font=("微软雅黑", 10, "bold"), command=self.start_processing_thread)
        self.btn_run.pack(fill="x", pady=5)
        
        self.lbl_status = tk.Label(left_panel, textvariable=self.status_var, wraplength=280, fg="blue")
        self.lbl_status.pack()

        # --- [页脚区域] ---
        frame_footer = tk.Frame(left_panel)
        frame_footer.pack(side="bottom", fill="x", pady=10)
        tk.Label(frame_footer, text="design by 比目鱼", font=("Microsoft YaHei", 9, "bold"), fg="#666666").pack()
        tk.Label(frame_footer, text="微信：inkstar97", font=("Microsoft YaHei", 8), fg="#999999").pack()
        # 【修改点2】更新页脚的版本号和日期
        tk.Label(frame_footer, text="v 1.0.6  2026.01.03", font=("Arial", 7), fg="#cccccc").pack()

        # === 右侧预览区域 UI ===
        top_bar = tk.Frame(right_panel, height=40, bg="#eeeeee", pady=5)
        top_bar.pack(side="top", fill="x")
        tk.Scale(top_bar, from_=0.5, to=3.0, resolution=0.1, orient="horizontal", length=120, 
                 variable=self.preview_zoom_var, command=self.update_preview, bg="#eeeeee", label="预览缩放").pack(side="left", padx=10)
        
        frame_page = tk.Frame(top_bar, bg="#eeeeee")
        frame_page.pack(side="left", expand=True) 
        tk.Button(frame_page, text="<<", command=lambda: self.change_page(-1)).pack(side="left")
        self.entry_page = tk.Entry(frame_page, width=5, justify="center")
        self.entry_page.pack(side="left", padx=5)
        self.entry_page.bind("<Return>", self.jump_to_page)
        tk.Label(frame_page, textvariable=self.page_info_var, bg="#eeeeee").pack(side="left")
        tk.Button(frame_page, text=">>", command=lambda: self.change_page(1)).pack(side="left")

        self.canvas = tk.Canvas(right_panel, bg="#808080", cursor="cross")
        self.canvas.pack(fill="both", expand=True)
        self.canvas.tag_bind("watermark", "<Button-1>", self.on_drag_start)
        self.canvas.tag_bind("watermark", "<B1-Motion>", self.on_drag_motion)
        self.canvas.tag_bind("watermark", "<ButtonRelease-1>", self.on_drag_stop)
        self._drag_data = {"x": 0, "y": 0}

    # --- 交互与业务逻辑 ---
    def toggle_range_entry(self, e=None):
        self.entry_range.config(state="normal" if self.range_mode_var.get() == "指定页面" else "disabled")

    def select_pdfs(self):
        files = filedialog.askopenfilenames(filetypes=[("PDF Files", "*.pdf")])
        if files:
            self.pdf_files = files
            self.lbl_pdf_info.config(text=f"已选 {len(files)} 个文件")
            self.load_pdf_doc(files[0])

    def load_pdf_doc(self, path):
        if self.current_doc: self.current_doc.close()
        try:
            self.current_doc = fitz.open(path)
            self.total_pages = self.current_doc.page_count
            self.current_page_idx = 0
            self.update_page_info_label()
            self.render_current_page_preview()
        except Exception as e:
            messagebox.showerror("错误", f"无法打开PDF: {e}")

    def update_page_info_label(self):
        self.page_info_var.set(f" / {self.total_pages}")
        self.entry_page.delete(0, tk.END); self.entry_page.insert(0, str(self.current_page_idx + 1))

    def change_page(self, delta):
        idx = self.current_page_idx + delta
        if 0 <= idx < self.total_pages:
            self.current_page_idx = idx; self.update_page_info_label(); self.render_current_page_preview()

    def jump_to_page(self, e=None):
        try:
            val = int(self.entry_page.get()) - 1
            if 0 <= val < self.total_pages:
                self.current_page_idx = val; self.render_current_page_preview()
            else: self.update_page_info_label()
        except: self.update_page_info_label()

    def render_current_page_preview(self):
        if not self.current_doc: return
        page = self.current_doc.load_page(self.current_page_idx)
        pix = page.get_pixmap(dpi=144)
        self.current_pdf_img = Image.open(BytesIO(pix.tobytes("ppm")))
        # 获取 Points 尺寸
        rot = page.rotation
        rect = page.rect
        self.vis_pdf_w, self.vis_pdf_h = (rect.height, rect.width) if rot % 180 == 90 else (rect.width, rect.height)
        self.update_preview()

    def select_watermark(self):
        f = filedialog.askopenfilename(filetypes=[("Images", "*.png *.jpg *.jpeg")])
        if f:
            self.watermark_path.set(f); self.current_wm_img = Image.open(f).convert("RGBA")
            self.update_preview()

    def update_preview(self, _=None):
        if not self.current_pdf_img: return
        c_w, c_h = max(self.canvas.winfo_width(), 600), max(self.canvas.winfo_height(), 600)
        
        # 144 DPI 图像显示的比例
        img_display_scale = min((c_w-40)/self.current_pdf_img.width, (c_h-40)/self.current_pdf_img.height) * self.preview_zoom_var.get()
        display_w, display_h = int(self.current_pdf_img.width * img_display_scale), int(self.current_pdf_img.height * img_display_scale)
        
        # 计算 Points 到像素的缩放率
        self.pt_to_canvas_scale = (display_w / self.vis_pdf_w)
        
        self.tk_bg_img = ImageTk.PhotoImage(self.current_pdf_img.resize((display_w, display_h), Image.Resampling.LANCZOS))
        
        self.canvas.delete("all")
        cx, cy = c_w/2, c_h/2
        self.canvas.create_image(cx, cy, image=self.tk_bg_img, tags="background")
        self.bg_offset_x, self.bg_offset_y = cx - display_w/2, cy - display_h/2
        
        if self.current_wm_img:
            wm_scale = self.scale_var.get()
            # 预览中水印的大小
            wm_w = int(self.current_wm_img.width * wm_scale * self.pt_to_canvas_scale)
            wm_h = int(self.current_wm_img.height * wm_scale * self.pt_to_canvas_scale)
            
            if wm_w > 0 and wm_h > 0:
                wm_edit = self.current_wm_img.resize((wm_w, wm_h), Image.Resampling.LANCZOS).rotate(self.angle_var.get(), expand=True)
                alpha = self.opacity_var.get()
                r, g, b, a = wm_edit.split()
                wm_edit.putalpha(ImageEnhance.Brightness(a).enhance(alpha))
                self.tk_wm_img = ImageTk.PhotoImage(wm_edit)
                
                # 初始化位置到中心
                if self.wm_x == 0: self.wm_x, self.wm_y = self.vis_pdf_w/2, self.vis_pdf_h/2
                
                vx = self.bg_offset_x + (self.wm_x * self.pt_to_canvas_scale)
                vy = self.bg_offset_y + ((self.vis_pdf_h - self.wm_y) * self.pt_to_canvas_scale)
                
                self.canvas.create_image(vx, vy, image=self.tk_wm_img, tags="watermark")
                self.lbl_coords.config(text=f"视觉坐标(Points): ({int(self.wm_x)}, {int(self.wm_y)})")

    def on_drag_start(self, e):
        self._drag_data["x"], self._drag_data["y"] = e.x, e.y
    def on_drag_motion(self, e):
        dx, dy = e.x - self._drag_data["x"], e.y - self._drag_data["y"]
        self.canvas.move("watermark", dx, dy)
        self._drag_data["x"], self._drag_data["y"] = e.x, e.y
    def on_drag_stop(self, e):
        c = self.canvas.coords("watermark")
        if c:
            self.wm_x = (c[0] - self.bg_offset_x) / self.pt_to_canvas_scale
            self.wm_y = self.vis_pdf_h - ((c[1] - self.bg_offset_y) / self.pt_to_canvas_scale)
            self.lbl_coords.config(text=f"视觉坐标(Points): ({int(self.wm_x)}, {int(self.wm_y)})")

    def set_pos_center(self):
        self.wm_x, self.wm_y = self.vis_pdf_w/2, self.vis_pdf_h/2; self.update_preview()
    def set_pos_top_left(self):
        margin = 50
        self.wm_x = margin + (self.current_wm_img.width*self.scale_var.get())/2
        self.wm_y = self.vis_pdf_h - margin - (self.current_wm_img.height*self.scale_var.get())/2
        self.update_preview()

    def start_processing_thread(self):
        if not self.pdf_files or not self.watermark_path.get():
            messagebox.showwarning("提示", "请先选择PDF文件和水印图片")
            return
        self.is_processing = True
        self.btn_run.config(state="disabled")
        threading.Thread(target=self.process_files, daemon=True).start()

    def process_files(self):
        # 预制水印图片数据
        wm_pil = self.current_wm_img.copy()
        ws, wa, wo = self.scale_var.get(), self.angle_var.get(), self.opacity_var.get()
        w, h = int(wm_pil.width * ws), int(wm_pil.height * ws)
        wm_pil = wm_pil.resize((w, h), Image.Resampling.LANCZOS).rotate(wa, expand=True)
        r, g, b, a = wm_pil.split()
        wm_pil.putalpha(ImageEnhance.Brightness(a).enhance(wo))
        img_byte_arr = BytesIO()
        wm_pil.save(img_byte_arr, format='PNG')
        wm_data = img_byte_arr.getvalue()

        mode = self.range_mode_var.get()
        custom = set()
        if mode == "指定页面":
            try:
                for p in self.custom_range_var.get().replace("，", ",").split(","):
                    if "-" in p:
                        a, b = p.split("-")
                        custom.update(range(int(a), int(b)+1))
                    elif p.strip(): custom.add(int(p))
            except: pass
        
        count = 0
        for i, path in enumerate(self.pdf_files):
            try:
                self.status_var.set(f"正在处理: {os.path.basename(path)}")
                doc = fitz.open(path)
                for page_idx in range(len(doc)):
                    if mode == "全部页面" or \
                       (mode == "奇数页" and (page_idx+1)%2!=0) or \
                       (mode == "偶数页" and (page_idx+1)%2==0) or \
                       (mode == "指定页面" and (page_idx+1) in custom):
                        
                        page = doc.load_page(page_idx)
                        # 物理坐标插入
                        pw, ph = wm_pil.width, wm_pil.height
                        rect_x0 = self.wm_x - pw/2
                        rect_y0 = (self.vis_pdf_h - self.wm_y) - ph/2
                        page.insert_image(fitz.Rect(rect_x0, rect_y0, rect_x0 + pw, rect_y0 + ph), stream=wm_data)
                
                doc.save(os.path.splitext(path)[0] + "_marked.pdf")
                doc.close()
                count += 1
            except Exception as e: print(f"处理失败 {path}: {e}")
            self.progress["value"] = (i+1)/len(self.pdf_files)*100
        
        self.status_var.set("处理完成")
        self.btn_run.config(state="normal")
        messagebox.showinfo("完成", f"成功处理 {count} 个文件")

if __name__ == "__main__":
    root = tk.Tk(); app = AdvancedWatermarkApp(root)
    root.bind("<Configure>", lambda e: app.update_preview() if e.widget == root else None)
    root.mainloop()