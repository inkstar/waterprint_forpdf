import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
import sys
import json
import threading
from io import BytesIO
from PIL import Image, ImageTk, ImageEnhance

# --- 依赖库检查 ---
def check_imports():
    try:
        global fitz
        import fitz  # PyMuPDF
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

# --- 通用滚动框架组件 ---
class ScrollableFrame(tk.Frame):
    def __init__(self, container, *args, **kwargs):
        super().__init__(container, *args, **kwargs)
        canvas = tk.Canvas(self, highlightthickness=0)
        scrollbar = ttk.Scrollbar(self, orient="vertical", command=canvas.yview)
        self.scrollable_window = tk.Frame(canvas)

        self.scrollable_window.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=self.scrollable_window, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # 绑定鼠标滚轮
        self.bind_mouse_wheel(canvas)

    def bind_mouse_wheel(self, widget):
        widget.bind_all("<MouseWheel>", lambda e: widget.yview_scroll(int(-1*(e.delta/120)), "units"))
        widget.bind_all("<Button-4>", lambda e: widget.yview_scroll(-1, "units"))
        widget.bind_all("<Button-5>", lambda e: widget.yview_scroll(1, "units"))

class AdvancedWatermarkApp:
    def __init__(self, root):
        self.root = root
        self.root.title("可视化 PDF 水印工具 v 1.0.7")
        self.root.geometry("1200x900")
        self.root.minsize(800, 600)
        
        # --- 核心数据 ---
        self.pdf_files = []
        self.current_doc = None
        self.current_page_idx = 0
        self.total_pages = 0
        self.current_pdf_img = None 
        self.current_wm_img = None
        self.pt_to_canvas_scale = 1.0    
        self.wm_x, self.wm_y = 0, 0
        
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

        self.load_config()
        self.setup_ui()
        
        if self.watermark_path.get() and os.path.exists(self.watermark_path.get()):
            try: self.current_wm_img = Image.open(self.watermark_path.get()).convert("RGBA")
            except: pass

        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

    def setup_ui(self):
        # 1. 主布局：左右分割
        self.main_paned = tk.PanedWindow(self.root, orient="horizontal", sashrelief="raised", sashwidth=4)
        self.main_paned.pack(fill="both", expand=True)

        # 2. 左侧滚动控制面板
        self.left_scroll_frame = ScrollableFrame(self.main_paned, width=350)
        self.main_paned.add(self.left_scroll_frame, stretch="never")
        ctrl_frame = self.left_scroll_frame.scrollable_window

        # --- 以下是控制栏的具体内容 ---
        # 文件选择
        lf_files = tk.LabelFrame(ctrl_frame, text="1. 文件选择", padx=10, pady=5)
        lf_files.pack(fill="x", padx=10, pady=5)
        tk.Button(lf_files, text="选择 PDF (支持多选)", command=self.select_pdfs).pack(fill="x", pady=2)
        self.lbl_pdf_info = tk.Label(lf_files, text="未加载", fg="gray")
        self.lbl_pdf_info.pack()

        # 水印图片
        lf_img = tk.LabelFrame(ctrl_frame, text="2. 水印图片", padx=10, pady=5)
        lf_img.pack(fill="x", padx=10, pady=5)
        tk.Button(lf_img, text="选择图片", command=self.select_watermark).pack(fill="x", pady=2)
        tk.Label(lf_img, textvariable=self.watermark_path, wraplength=250, fg="gray", font=("Arial", 8)).pack()

        # 水印样式
        lf_style = tk.LabelFrame(ctrl_frame, text="3. 水印样式", padx=10, pady=5)
        lf_style.pack(fill="x", padx=10, pady=5)
        tk.Label(lf_style, text="水印大小:").pack(anchor="w")
        tk.Scale(lf_style, from_=0.01, to=3.0, resolution=0.01, orient="horizontal", variable=self.scale_var, command=self.update_preview).pack(fill="x")
        tk.Label(lf_style, text="透明度:").pack(anchor="w")
        tk.Scale(lf_style, from_=0.1, to=1.0, resolution=0.1, orient="horizontal", variable=self.opacity_var, command=self.update_preview).pack(fill="x")
        tk.Label(lf_style, text="旋转角度:").pack(anchor="w")
        tk.Scale(lf_style, from_=0, to=360, resolution=5, orient="horizontal", variable=self.angle_var, command=self.update_preview).pack(fill="x")

        # 位置控制
        lf_pos = tk.LabelFrame(ctrl_frame, text="4. 位置控制", padx=10, pady=5)
        lf_pos.pack(fill="x", padx=10, pady=5)
        btn_frame = tk.Frame(lf_pos)
        btn_frame.pack(fill="x")
        tk.Button(btn_frame, text="↖ 左上角", command=self.set_pos_top_left).pack(side="left", expand=True)
        tk.Button(btn_frame, text="✛ 居中", command=self.set_pos_center).pack(side="left", expand=True)
        self.lbl_coords = tk.Label(lf_pos, text="X: 0, Y: 0", pady=5)
        self.lbl_coords.pack()

        # 应用范围
        lf_range = tk.LabelFrame(ctrl_frame, text="5. 应用范围", padx=10, pady=5)
        lf_range.pack(fill="x", padx=10, pady=5)
        cb_range = ttk.Combobox(lf_range, values=["全部页面", "奇数页", "偶数页", "指定页面"], state="readonly", textvariable=self.range_mode_var)
        cb_range.pack(fill="x", pady=2)
        cb_range.bind("<<ComboboxSelected>>", self.toggle_range_entry)
        self.entry_range = tk.Entry(lf_range, textvariable=self.custom_range_var, state="disabled")
        self.entry_range.pack(fill="x", pady=2)

        # 执行区域
        self.progress = ttk.Progressbar(ctrl_frame, orient="horizontal", mode="determinate")
        self.progress.pack(fill="x", padx=10, pady=20)
        self.btn_run = tk.Button(ctrl_frame, text="开始批量处理", bg="#28a745", fg="white", height=2, font=("微软雅黑", 10, "bold"), command=self.start_processing_thread)
        self.btn_run.pack(fill="x", padx=10, pady=5)
        tk.Label(ctrl_frame, textvariable=self.status_var, wraplength=280, fg="blue").pack(pady=5)

        # 页脚
        tk.Label(ctrl_frame, text="design by 比目鱼\n微信：inkstar97\nv 1.0.7  2026.01.03", font=("Arial", 8), fg="#999999", pady=20).pack()

        # 3. 右侧预览区域 (带双向滚动条)
        preview_container = tk.Frame(self.main_paned)
        self.main_paned.add(preview_container, stretch="always")

        # 预览顶部工具栏
        top_bar = tk.Frame(preview_container, height=40, bg="#f8f9fa", pady=5)
        top_bar.pack(side="top", fill="x")
        tk.Scale(top_bar, from_=0.5, to=3.0, resolution=0.1, orient="horizontal", length=150, 
                 variable=self.preview_zoom_var, command=self.update_preview, label="预览缩放").pack(side="left", padx=10)
        
        frame_page = tk.Frame(top_bar, bg="#f8f9fa")
        frame_page.pack(side="right", padx=20)
        tk.Button(frame_page, text="< 上一页", command=lambda: self.change_page(-1)).pack(side="left")
        self.entry_page = tk.Entry(frame_page, width=5, justify="center")
        self.entry_page.pack(side="left", padx=5)
        self.entry_page.bind("<Return>", self.jump_to_page)
        tk.Label(frame_page, textvariable=self.page_info_var, bg="#f8f9fa").pack(side="left")
        tk.Button(frame_page, text="下一页 >", command=lambda: self.change_page(1)).pack(side="left")

        # 预览主体：带滚动条的 Canvas
        self.canvas_frame = tk.Frame(preview_container, bg="#444444")
        self.canvas_frame.pack(fill="both", expand=True)

        self.v_scroll = ttk.Scrollbar(self.canvas_frame, orient="vertical")
        self.h_scroll = ttk.Scrollbar(self.canvas_frame, orient="horizontal")
        self.canvas = tk.Canvas(self.canvas_frame, bg="#808080", 
                                xscrollcommand=self.h_scroll.set, 
                                yscrollcommand=self.v_scroll.set,
                                highlightthickness=0)
        
        self.v_scroll.config(command=self.canvas.yview)
        self.h_scroll.config(command=self.canvas.xview)

        self.v_scroll.pack(side="right", fill="y")
        self.h_scroll.pack(side="bottom", fill="x")
        self.canvas.pack(side="left", fill="both", expand=True)

        # 绑定事件
        self.canvas.tag_bind("watermark", "<Button-1>", self.on_drag_start)
        self.canvas.tag_bind("watermark", "<B1-Motion>", self.on_drag_motion)
        self.canvas.tag_bind("watermark", "<ButtonRelease-1>", self.on_drag_stop)
        self._drag_data = {"x": 0, "y": 0}

    # --- 逻辑部分 (保持原有逻辑并优化坐标计算) ---
    def update_preview(self, _=None):
        if not self.current_pdf_img: return
        
        zoom = self.preview_zoom_var.get()
        display_w = int(self.current_pdf_img.width * zoom)
        display_h = int(self.current_pdf_img.height * zoom)
        
        self.pt_to_canvas_scale = (display_w / self.vis_pdf_w)
        
        self.tk_bg_img = ImageTk.PhotoImage(self.current_pdf_img.resize((display_w, display_h), Image.Resampling.LANCZOS))
        
        self.canvas.delete("all")
        # 将背景放在 Canvas 的 0,0 位置
        self.canvas.create_image(0, 0, image=self.tk_bg_img, tags="background", anchor="nw")
        # 设置滚动区域为图像大小
        self.canvas.config(scrollregion=(0, 0, display_w, display_h))
        
        if self.current_wm_img:
            wm_scale = self.scale_var.get()
            wm_w = int(self.current_wm_img.width * wm_scale * self.pt_to_canvas_scale)
            wm_h = int(self.current_wm_img.height * wm_scale * self.pt_to_canvas_scale)
            
            if wm_w > 0 and wm_h > 0:
                wm_edit = self.current_wm_img.resize((wm_w, wm_h), Image.Resampling.LANCZOS).rotate(self.angle_var.get(), expand=True)
                alpha = self.opacity_var.get()
                r, g, b, a = wm_edit.split()
                wm_edit.putalpha(ImageEnhance.Brightness(a).enhance(alpha))
                self.tk_wm_img = ImageTk.PhotoImage(wm_edit)
                
                if self.wm_x == 0: self.wm_x, self.wm_y = self.vis_pdf_w/2, self.vis_pdf_h/2
                
                # 计算 Canvas 坐标 (不再依赖居中偏移，直接使用绝对坐标)
                vx = self.wm_x * self.pt_to_canvas_scale
                vy = (self.vis_pdf_h - self.wm_y) * self.pt_to_canvas_scale
                
                self.canvas.create_image(vx, vy, image=self.tk_wm_img, tags="watermark")
                self.lbl_coords.config(text=f"视觉坐标(Points): ({int(self.wm_x)}, {int(self.wm_y)})")

    def on_drag_motion(self, e):
        # 拖拽时需要考虑 Canvas 的滚动偏移
        cx = self.canvas.canvasx(e.x)
        cy = self.canvas.canvasy(e.y)
        dx, dy = cx - self._drag_data["x"], cy - self._drag_data["y"]
        self.canvas.move("watermark", dx, dy)
        self._drag_data["x"], self._drag_data["y"] = cx, cy

    def on_drag_start(self, e):
        self._drag_data["x"] = self.canvas.canvasx(e.x)
        self._drag_data["y"] = self.canvas.canvasy(e.y)

    def on_drag_stop(self, e):
        c = self.canvas.coords("watermark")
        if c:
            self.wm_x = c[0] / self.pt_to_canvas_scale
            self.wm_y = self.vis_pdf_h - (c[1] / self.pt_to_canvas_scale)
            self.lbl_coords.config(text=f"视觉坐标(Points): ({int(self.wm_x)}, {int(self.wm_y)})")

    # --- 后续通用方法 (复用之前的逻辑) ---
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
        except Exception as e: messagebox.showerror("错误", f"无法打开PDF: {e}")

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
        rot = page.rotation
        rect = page.rect
        self.vis_pdf_w, self.vis_pdf_h = (rect.height, rect.width) if rot % 180 == 90 else (rect.width, rect.height)
        self.update_preview()

    def select_watermark(self):
        f = filedialog.askopenfilename(filetypes=[("Images", "*.png *.jpg *.jpeg")])
        if f:
            self.watermark_path.set(f); self.current_wm_img = Image.open(f).convert("RGBA")
            self.update_preview()

    def set_pos_center(self):
        self.wm_x, self.wm_y = self.vis_pdf_w/2, self.vis_pdf_h/2; self.update_preview()

    def set_pos_top_left(self):
        margin = 50
        self.wm_x = margin + (self.current_wm_img.width*self.scale_var.get())/2
        self.wm_y = self.vis_pdf_h - margin - (self.current_wm_img.height*self.scale_var.get())/2
        self.update_preview()

    def toggle_range_entry(self, e=None):
        self.entry_range.config(state="normal" if self.range_mode_var.get() == "指定页面" else "disabled")

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

    def start_processing_thread(self):
        if not self.pdf_files or not self.watermark_path.get():
            messagebox.showwarning("提示", "请先选择PDF文件和水印图片")
            return
        self.btn_run.config(state="disabled")
        threading.Thread(target=self.process_files, daemon=True).start()

    def process_files(self):
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
                        pw, ph = wm_pil.width, wm_pil.height
                        rect_x0 = self.wm_x - pw/2
                        rect_y0 = (self.vis_pdf_h - self.wm_y) - ph/2
                        page.insert_image(fitz.Rect(rect_x0, rect_y0, rect_x0 + pw, rect_y0 + ph), stream=wm_data)
                doc.save(os.path.splitext(path)[0] + "_marked.pdf")
                doc.close()
                count += 1
            except Exception as e: print(f"失败: {e}")
            self.progress["value"] = (i+1)/len(self.pdf_files)*100
        
        self.status_var.set("处理完成")
        self.btn_run.config(state="normal")
        messagebox.showinfo("完成", f"成功处理 {count} 个文件")

if __name__ == "__main__":
    root = tk.Tk(); app = AdvancedWatermarkApp(root)
    root.mainloop()