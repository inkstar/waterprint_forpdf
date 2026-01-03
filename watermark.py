import tkinter as tk
from tkinter import filedialog, messagebox, ttk, colorchooser, font
import os
import sys
import json
import threading
import webbrowser
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

# --- 核心配置 ---
def get_config_path():
    # 将配置文件存放在用户主目录下，避免在程序目录生成
    return os.path.join(os.path.expanduser("~"), ".pdf_watermark_settings.json")

CONFIG_FILE = get_config_path()

# --- 通用滚动框架组件 ---
def unified_mouse_wheel_bind(widget):
    """统一处理 macOS/Windows/Linux 的鼠标滚轮与触控板绑定"""
    def _on_mousewheel(e):
        if sys.platform == "darwin": # macOS
            widget.yview_scroll(int(-1 * e.delta), "units")
        else: # Windows/Linux
            widget.yview_scroll(int(-1 * (e.delta / 120)), "units")

    def _on_shift_mousewheel(e):
        if hasattr(widget, 'xview'):
            if sys.platform == "darwin": # macOS
                widget.xview_scroll(int(-1 * e.delta), "units")
            else: # Windows/Linux
                widget.xview_scroll(int(-1 * (e.delta / 120)), "units")

    # 绑定垂直滚动
    widget.bind("<MouseWheel>", _on_mousewheel)
    # 绑定水平滚动
    widget.bind("<Shift-MouseWheel>", _on_shift_mousewheel)
    
    # Linux 兼容性
    widget.bind("<Button-4>", lambda e: widget.yview_scroll(-1, "units"))
    widget.bind("<Button-5>", lambda e: widget.yview_scroll(1, "units"))

    # 关键：当鼠标进入组件区域时，强制夺取焦点以接收滚动事件
    widget.bind("<Enter>", lambda e: widget.focus_set())

class ScrollableFrame(tk.Frame):
    def __init__(self, container, *args, **kwargs):
        super().__init__(container, *args, **kwargs)
        self.canvas = tk.Canvas(self, highlightthickness=0)
        scrollbar = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.scrollable_window = tk.Frame(self.canvas)

        self.scrollable_window.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )

        self.canvas.create_window((0, 0), window=self.scrollable_window, anchor="nw")
        self.canvas.configure(yscrollcommand=scrollbar.set)

        self.canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # 使用统一绑定函数
        unified_mouse_wheel_bind(self.canvas)
        # 同时为内部框架也绑定，确保鼠标悬停在按钮上也能滚动
        unified_mouse_wheel_bind(self.scrollable_window)

class AdvancedWatermarkApp:
    def __init__(self, root):
        self.root = root
        self.root.title("可视化 PDF 水印工具 v 1.2.4")
        self.root.geometry("1200x900")
        self.root.minsize(800, 600)
        
        # --- 核心数据 ---
        self.pdf_files = []
        self.current_pdf_idx = 0
        self.current_doc = None
        self.current_page_idx = 0
        self.total_pages = 0
        self.current_pdf_img = None 
        self.pt_to_canvas_scale = 1.0    
        
        # 多水印支持
        self.watermarks = [] # 存储字典：{..., grid_mode, grid_gap_x, grid_gap_y}
        self.selected_wm_idx = -1
        
        # --- 变量 (当前选中的水印属性) ---
        self.wm_text_var = tk.StringVar(value="测试水印")
        self.wm_color_var = tk.StringVar(value="#FF0000")
        self.wm_font_var = tk.StringVar(value="Arial")
        self.grid_mode_var = tk.BooleanVar(value=False)
        self.grid_gap_x_var = tk.DoubleVar(value=100)
        self.grid_gap_y_var = tk.DoubleVar(value=100)
        
        self.watermark_path = tk.StringVar()
        self.scale_var = tk.DoubleVar(value=1.0)
        self.opacity_var = tk.DoubleVar(value=0.5)
        self.angle_var = tk.DoubleVar(value=0)
        
        # 全局变量
        self.preview_zoom_var = tk.DoubleVar(value=1.0)
        self.range_mode_var = tk.StringVar(value="全部页面")
        self.custom_range_var = tk.StringVar(value="")
        self.output_dir_var = tk.StringVar(value="原文件目录")
        self.output_suffix_var = tk.StringVar(value="_marked")
        self.status_var = tk.StringVar(value="准备就绪")
        self.page_info_var = tk.StringVar(value="0 / 0")

        self.last_output_path = "" # 记录最后一次生成的文件或目录
        self.load_config()
        self.setup_ui()
        
        if self.watermark_path.get() and os.path.exists(self.watermark_path.get()):
            try: self.current_wm_img = Image.open(self.watermark_path.get()).convert("RGBA")
            except: pass

        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

    def open_output_folder(self):
        # 如果没有导出过，尝试获取当前设置的目录
        path = getattr(self, 'last_output_dir', None)
        if not path:
            od = self.output_dir_var.get()
            if od != "原文件目录" and os.path.exists(od):
                path = od
        
        if not path:
            messagebox.showinfo("提示", "请先执行批量处理，或在输出设置中指定目录")
            return
            
        try:
            if sys.platform == 'win32':
                os.startfile(path)
            elif sys.platform == 'darwin':
                import subprocess
                subprocess.run(['open', path])
            else:
                import subprocess
                subprocess.run(['xdg-open', path])
        except Exception as e:
            messagebox.showerror("错误", f"无法打开文件夹: {e}")

    def open_output_folder(self):
        if not self.last_output_path or not os.path.exists(self.last_output_path):
            messagebox.showwarning("提示", "尚未生成文件或输出目录不存在")
            return
            
        path = self.last_output_path
        if os.path.isfile(path):
            path = os.path.dirname(path)
            
        if sys.platform == "darwin": # macOS
            import subprocess
            subprocess.run(["open", path])
        elif sys.platform == "win32": # Windows
            os.startfile(path)
        else: # Linux
            import subprocess
            subprocess.run(["xdg-open", path])

    def open_feedback(self, e=None):
        webbrowser.open("https://v.wjx.cn/vm/QgqYdV1.aspx")

    def create_modern_scale(self, parent, label_text, var, from_val, to_val, width=200, is_int=False, command=None):
        frame = tk.Frame(parent)
        
        # 顶部标签栏：标题 + 数值显示
        header = tk.Frame(frame)
        header.pack(fill="x")
        tk.Label(header, text=label_text, font=("Arial", 9, "bold")).pack(side="left")
        
        val_lbl = tk.Label(header, text="", font=("Arial", 9), fg="#666666")
        val_lbl.pack(side="right")
        
        # 更新数值显示的闭包函数
        def update_val_label(*args):
            v = var.get()
            val_lbl.config(text=str(int(v)) if is_int else f"{v:.2f}")
            if command:
                command()
            else:
                self.update_wm_from_ui() # 默认实时更新预览

        # 绑定变量变化
        var.trace_add("write", update_val_label)
        
        # ttk 滑块
        s = ttk.Scale(frame, from_=from_val, to=to_val, variable=var, orient="horizontal", length=width)
        s.pack(fill="x", pady=(2, 0))
        
        # 初始化显示
        update_val_label()
        
        return frame

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
        
        # 文件切换控制
        self.frame_file_switch = tk.Frame(lf_files)
        self.frame_file_switch.pack(fill="x", pady=2)
        tk.Button(self.frame_file_switch, text="< 上一个文件", command=lambda: self.change_file(-1), font=("Arial", 7)).pack(side="left", expand=True)
        tk.Button(self.frame_file_switch, text="下一个文件 >", command=lambda: self.change_file(1), font=("Arial", 7)).pack(side="left", expand=True)
        
        self.lbl_pdf_info = tk.Label(lf_files, text="未加载", fg="gray", wraplength=300)
        self.lbl_pdf_info.pack()

        # 模板管理区域
        lf_templates = tk.LabelFrame(ctrl_frame, text="水印模板 (Presets)", padx=10, pady=5)
        lf_templates.pack(fill="x", padx=10, pady=5)
        
        self.cb_templates = ttk.Combobox(lf_templates, state="readonly")
        self.cb_templates.pack(fill="x", pady=2)
        
        btn_temp_frame = tk.Frame(lf_templates)
        btn_temp_frame.pack(fill="x")
        tk.Button(btn_temp_frame, text="保存为模板", command=self.save_template, font=("Arial", 8)).pack(side="left", expand=True)
        tk.Button(btn_temp_frame, text="加载模板", command=self.load_template, font=("Arial", 8)).pack(side="left", expand=True)
        tk.Button(btn_temp_frame, text="删除模板", command=self.delete_template, font=("Arial", 8), fg="red").pack(side="left", expand=True)

        # 水印管理区域
        lf_wm_list = tk.LabelFrame(ctrl_frame, text="2. 水印列表", padx=10, pady=5)
        lf_wm_list.pack(fill="x", padx=10, pady=5)
        
        btn_wm_actions = tk.Frame(lf_wm_list)
        btn_wm_actions.pack(fill="x")
        tk.Button(btn_wm_actions, text="+ 图片水印", command=self.add_image_watermark, font=("Arial", 8)).pack(side="left", expand=True)
        tk.Button(btn_wm_actions, text="+ 文字水印", command=self.add_text_watermark, font=("Arial", 8)).pack(side="left", expand=True)
        tk.Button(btn_wm_actions, text="删除选中", command=self.delete_selected_wm, font=("Arial", 8), fg="red").pack(side="left", expand=True)

        self.wm_listbox = tk.Listbox(lf_wm_list, height=4)
        self.wm_listbox.pack(fill="x", pady=5)
        self.wm_listbox.bind("<<ListboxSelect>>", self.on_wm_select)

        # 水印属性编辑 (针对选中项)
        self.lf_edit = tk.LabelFrame(ctrl_frame, text="3. 水印属性编辑", padx=10, pady=5)
        self.lf_edit.pack(fill="x", padx=10, pady=5)
        
        # 文字水印特有控件
        self.frame_text_edit = tk.Frame(self.lf_edit)
        tk.Label(self.frame_text_edit, text="文字内容:").pack(anchor="w")
        self.entry_wm_text = tk.Entry(self.frame_text_edit, textvariable=self.wm_text_var)
        self.entry_wm_text.pack(fill="x")
        self.entry_wm_text.bind("<KeyRelease>", lambda e: self.update_wm_from_ui())
        
        # 字体和颜色选择
        fc_frame = tk.Frame(self.frame_text_edit)
        fc_frame.pack(fill="x", pady=5)
        
        tk.Label(fc_frame, text="字体:").pack(side="left")
        self.available_fonts = sorted(font.families())
        self.cb_font = ttk.Combobox(fc_frame, textvariable=self.wm_font_var, values=self.available_fonts, state="readonly", width=15)
        self.cb_font.pack(side="left", padx=5)
        self.cb_font.bind("<<ComboboxSelected>>", lambda e: self.update_wm_from_ui())
        
        self.btn_color = tk.Button(fc_frame, text="颜色", command=self.pick_color, width=5)
        self.btn_color.pack(side="left", padx=5)
        
        # 图片水印特有控件
        self.frame_img_edit = tk.Frame(self.lf_edit)
        tk.Button(self.frame_img_edit, text="更换图片", command=self.select_watermark).pack(fill="x")
        
        # 通用属性
        self.create_modern_scale(self.lf_edit, "水印大小/字号:", self.scale_var, 0.01, 3.0).pack(fill="x", pady=5)
        self.create_modern_scale(self.lf_edit, "透明度:", self.opacity_var, 0.1, 1.0).pack(fill="x", pady=5)
        self.create_modern_scale(self.lf_edit, "旋转角度:", self.angle_var, 0, 360, is_int=True).pack(fill="x", pady=5)

        # 铺满设置 (阵列模式)
        self.frame_grid_edit = tk.LabelFrame(self.lf_edit, text="阵列铺满设置", padx=5, pady=5)
        self.frame_grid_edit.pack(fill="x", pady=5)
        
        tk.Checkbutton(self.frame_grid_edit, text="开启全屏铺满", variable=self.grid_mode_var, command=self.update_wm_from_ui).pack(anchor="w")
        self.create_modern_scale(self.frame_grid_edit, "横向间距:", self.grid_gap_x_var, 50, 500, is_int=True).pack(fill="x")
        self.create_modern_scale(self.frame_grid_edit, "纵向间距:", self.grid_gap_y_var, 50, 500, is_int=True).pack(fill="x")

        # 位置控制
        lf_pos = tk.LabelFrame(ctrl_frame, text="4. 位置控制", padx=10, pady=5)
        lf_pos.pack(fill="x", padx=10, pady=5)
        btn_frame = tk.Frame(lf_pos)
        btn_frame.pack(fill="x")
        tk.Button(btn_frame, text="↖ 左上角", command=self.set_pos_top_left, font=("Arial", 8)).pack(side="left", expand=True)
        tk.Button(btn_frame, text="↗ 右上角", command=self.set_pos_top_right, font=("Arial", 8)).pack(side="left", expand=True)
        tk.Button(btn_frame, text="✛ 居中", command=self.set_pos_center, font=("Arial", 8)).pack(side="left", expand=True)
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

        # 6. 输出设置
        lf_output = tk.LabelFrame(ctrl_frame, text="6. 输出设置", padx=10, pady=5)
        lf_output.pack(fill="x", padx=10, pady=5)
        
        tk.Label(lf_output, text="文件名后缀 (如 _marked):").pack(anchor="w")
        tk.Entry(lf_output, textvariable=self.output_suffix_var).pack(fill="x", pady=2)
        
        tk.Button(lf_output, text="选择输出目录", command=self.select_output_dir).pack(fill="x", pady=2)
        tk.Label(lf_output, textvariable=self.output_dir_var, wraplength=250, fg="gray", font=("Arial", 8)).pack()
        tk.Button(lf_output, text="恢复默认 (原目录)", command=self.reset_output_dir, font=("Arial", 7), fg="blue", bd=0, cursor="hand2").pack(anchor="e")

        # 执行区域
        self.progress = ttk.Progressbar(ctrl_frame, orient="horizontal", mode="determinate")
        self.progress.pack(fill="x", padx=10, pady=20)
        self.btn_run = tk.Button(ctrl_frame, text="开始批量处理", bg="#28a745", fg="black", height=2, font=("微软雅黑", 10, "bold"), command=self.start_processing_thread)
        self.btn_run.pack(fill="x", padx=10, pady=5)
        
        self.btn_open_folder = tk.Button(ctrl_frame, text="📂 打开输出文件夹", command=self.open_output_folder, font=("Arial", 9))
        self.btn_open_folder.pack(fill="x", padx=10, pady=2)
        
        tk.Label(ctrl_frame, textvariable=self.status_var, wraplength=280, fg="blue").pack(pady=5)

        # 页脚
        footer_frame = tk.Frame(ctrl_frame)
        footer_frame.pack(side="bottom", fill="x", pady=20)
        
        tk.Label(footer_frame, text="design by 比目鱼", font=("Arial", 8, "bold"), fg="#999999").pack()
        tk.Label(footer_frame, text="微信：inkstar97", font=("Arial", 8), fg="#999999").pack()
        
        link_lbl = tk.Label(footer_frame, text="点此提交使用反馈", font=("Arial", 8, "underline"), fg="#0066cc", cursor="hand2")
        link_lbl.pack(pady=5)
        link_lbl.bind("<Button-1>", self.open_feedback)
        
        tk.Label(footer_frame, text="v 1.2.4  2026.01.03", font=("Arial", 7), fg="#cccccc").pack()

        # 3. 右侧预览区域 (带双向滚动条)
        preview_container = tk.Frame(self.main_paned)
        self.main_paned.add(preview_container, stretch="always")

        # 预览顶部工具栏
        top_bar = tk.Frame(preview_container, height=40, bg="#f8f9fa", pady=5)
        top_bar.pack(side="top", fill="x")
        
        zoom_frame = self.create_modern_scale(top_bar, "预览缩放:", self.preview_zoom_var, 0.5, 1.5, width=150, command=self.update_preview)
        zoom_frame.pack(side="left", padx=20)
        zoom_frame.config(bg="#f8f9fa")
        for child in zoom_frame.winfo_children():
            try: child.config(bg="#f8f9fa")
            except: pass
        
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
        self.canvas.bind("<B1-Motion>", self.on_drag_motion)
        self.canvas.bind("<ButtonRelease-1>", self.on_drag_stop)
        self._drag_data = {"x": 0, "y": 0}
        
        # 预览区也支持触控板滚动 (支持垂直和水平)
        unified_mouse_wheel_bind(self.canvas)

    # --- 逻辑部分 (保持原有逻辑并优化坐标计算) ---
    def update_preview(self, _=None):
        if not self.current_pdf_img: return
        
        zoom = self.preview_zoom_var.get()
        display_w = int(self.current_pdf_img.width * zoom)
        display_h = int(self.current_pdf_img.height * zoom)
        self.pt_to_canvas_scale = (display_w / self.vis_pdf_w)
        
        self.tk_bg_img = ImageTk.PhotoImage(self.current_pdf_img.resize((display_w, display_h), Image.Resampling.LANCZOS))
        self.canvas.delete("all")
        self.canvas.create_image(0, 0, image=self.tk_bg_img, tags="background", anchor="nw")
        self.canvas.config(scrollregion=(0, 0, display_w, display_h))
        
        self.tk_wm_images = []

        for i, wm in enumerate(self.watermarks):
            tag = f"wm_{i}"
            
            # 计算绘制位置 (阵列模式或单点模式)
            positions = []
            if wm.get('grid_mode'):
                # 阵列模式：铺满全屏
                gap_x = wm.get('grid_gap_x', 150)
                gap_y = wm.get('grid_gap_y', 150)
                # 从 -gap 开始，确保边缘也被覆盖
                for cur_x in range(0, int(self.vis_pdf_w + gap_x), int(gap_x)):
                    for cur_y in range(0, int(self.vis_pdf_h + gap_y), int(gap_y)):
                        positions.append((cur_x, cur_y))
            else:
                # 单点模式
                positions.append((wm['x'], wm['y']))

            if wm['type'] == 'image':
                # 渲染图片水印
                wm_scale = wm['scale']
                wm_w = int(wm['img_obj'].width * wm_scale * self.pt_to_canvas_scale)
                wm_h = int(wm['img_obj'].height * wm_scale * self.pt_to_canvas_scale)
            
            if wm_w > 0 and wm_h > 0:
                    wm_edit = wm['img_obj'].resize((wm_w, wm_h), Image.Resampling.LANCZOS).rotate(wm['angle'], expand=True)
                    alpha = wm['opacity']
                r, g, b, a = wm_edit.split()
                wm_edit.putalpha(ImageEnhance.Brightness(a).enhance(alpha))
                    tk_img = ImageTk.PhotoImage(wm_edit)
                    self.tk_wm_images.append(tk_img) # 保持引用
                    
                    for pos_x, pos_y in positions:
                        vx = pos_x * self.pt_to_canvas_scale
                        vy = (self.vis_pdf_h - pos_y) * self.pt_to_canvas_scale
                        self.canvas.create_image(vx, vy, image=tk_img, tags=("watermark", tag))
            else:
                # 渲染文字水印
                font_size = int(30 * wm['scale'] * self.pt_to_canvas_scale)
                for pos_x, pos_y in positions:
                    vx = pos_x * self.pt_to_canvas_scale
                    vy = (self.vis_pdf_h - pos_y) * self.pt_to_canvas_scale
                    self.canvas.create_text(vx, vy, text=wm['content'], font=(wm.get('font', 'Arial'), font_size), 
                                           fill=wm.get('color', '#FF0000'), angle=wm['angle'], 
                                           stipple="gray50" if wm['opacity'] < 0.8 else "", 
                                           tags=("watermark", tag))
            
            if i == self.selected_wm_idx:
                bbox = self.canvas.bbox(tag)
                if bbox:
                    self.canvas.create_rectangle(bbox, outline="red", dash=(4,4), tags="selection_box")

    def on_drag_start(self, e):
        cx = self.canvas.canvasx(e.x)
        cy = self.canvas.canvasy(e.y)
        
        # 使用 find_withtag("current") 获取当前点击的水印
        items = self.canvas.find_withtag("current")
        if items:
            tags = self.canvas.gettags(items[0])
            for t in tags:
                if t.startswith("wm_"):
                    new_idx = int(t.split("_")[1])
                    if new_idx != self.selected_wm_idx:
                        self.selected_wm_idx = new_idx
                        self.wm_listbox.selection_clear(0, tk.END)
                        self.wm_listbox.selection_set(self.selected_wm_idx)
                        self.on_wm_select() # 这会触发 update_preview
                    break
        
        self._drag_data["x"] = cx
        self._drag_data["y"] = cy

    def on_drag_motion(self, e):
        if self.selected_wm_idx < 0: return
        cx = self.canvas.canvasx(e.x)
        cy = self.canvas.canvasy(e.y)
        dx, dy = cx - self._drag_data["x"], cy - self._drag_data["y"]
        
        # 移动选中的水印和选择框
        tag = f"wm_{self.selected_wm_idx}"
        self.canvas.move(tag, dx, dy)
        self.canvas.move("selection_box", dx, dy)
        
        # 实时更新位置数据，但不触发重绘 (重绘太慢)
        wm = self.watermarks[self.selected_wm_idx]
        c = self.canvas.coords(tag)
        if c:
            wm['x'] = c[0] / self.pt_to_canvas_scale
            wm['y'] = self.vis_pdf_h - (c[1] / self.pt_to_canvas_scale)
            self.lbl_coords.config(text=f"视觉坐标(Points): ({int(wm['x'])}, {int(wm['y'])})")
        
        self._drag_data["x"] = cx
        self._drag_data["y"] = cy

    def on_drag_stop(self, e):
        if self.selected_wm_idx < 0: return
        tag = f"wm_{self.selected_wm_idx}"
        c = self.canvas.coords(tag)
        if c:
            wm = self.watermarks[self.selected_wm_idx]
            wm['x'] = c[0] / self.pt_to_canvas_scale
            wm['y'] = self.vis_pdf_h - (c[1] / self.pt_to_canvas_scale)
            self.lbl_coords.config(text=f"视觉坐标(Points): ({int(wm['x'])}, {int(wm['y'])})")

    # --- 后续通用方法 (复用之前的逻辑) ---
    def select_output_dir(self):
        d = filedialog.askdirectory()
        if d:
            self.output_dir_var.set(d)

    def reset_output_dir(self):
        self.output_dir_var.set("原文件目录")

    def select_pdfs(self):
        files = filedialog.askopenfilenames(filetypes=[("PDF Files", "*.pdf")])
        if files:
            self.pdf_files = list(files)
            self.current_pdf_idx = 0
            self.update_file_info_label()
            self.load_pdf_doc(self.pdf_files[0])

    def update_file_info_label(self):
        if self.pdf_files:
            fname = os.path.basename(self.pdf_files[self.current_pdf_idx])
            self.lbl_pdf_info.config(text=f"文件 ({self.current_pdf_idx + 1}/{len(self.pdf_files)}):\n{fname}", fg="blue")
        else:
            self.lbl_pdf_info.config(text="未加载", fg="gray")

    def change_file(self, delta):
        if not self.pdf_files: return
        new_idx = self.current_pdf_idx + delta
        if 0 <= new_idx < len(self.pdf_files):
            self.current_pdf_idx = new_idx
            self.update_file_info_label()
            self.load_pdf_doc(self.pdf_files[self.current_pdf_idx])

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

    def pick_color(self):
        color = colorchooser.askcolor(initialcolor=self.wm_color_var.get())[1]
        if color:
            self.wm_color_var.set(color)
            self.btn_color.config(bg=color)
            self.update_wm_from_ui()

    def save_template(self):
        from tkinter import simpledialog
        name = simpledialog.askstring("保存模板", "请输入模板名称:")
        if not name: return
        
        # 准备模板数据（移除无法序列化的 img_obj）
        temp_wms = []
        for wm in self.watermarks:
            w = wm.copy()
            if 'img_obj' in w: del w['img_obj']
            temp_wms.append(w)
            
        if not hasattr(self, 'all_templates'): self.all_templates = {}
        self.all_templates[name] = temp_wms
        self.update_template_cb()
        messagebox.showinfo("成功", f"模板 '{name}' 已保存")

    def load_template(self):
        name = self.cb_templates.get()
        if not name or name not in self.all_templates: return
        
        self.watermarks = []
        for w_data in self.all_templates[name]:
            w = w_data.copy()
            if w['type'] == 'image' and os.path.exists(w['path']):
                w['img_obj'] = Image.open(w['path']).convert("RGBA")
            self.watermarks.append(w)
            
        self.refresh_wm_list()
        self.update_preview()
        messagebox.showinfo("成功", f"已加载模板 '{name}'")

    def delete_template(self):
        name = self.cb_templates.get()
        if name and messagebox.askyesno("确认", f"确定要删除模板 '{name}' 吗？"):
            del self.all_templates[name]
            self.update_template_cb()

    def update_template_cb(self):
        if not hasattr(self, 'all_templates'): self.all_templates = {}
        names = list(self.all_templates.keys())
        self.cb_templates.config(values=names)
        if names: self.cb_templates.current(0)

    def add_image_watermark(self):
        f = filedialog.askopenfilename(filetypes=[("Images", "*.png *.jpg *.jpeg")])
        if f:
            wm = {
                "type": "image",
                "path": f,
                "scale": 1.0,
                "opacity": 0.5,
                "angle": 0,
                "x": self.vis_pdf_w / 2 if self.vis_pdf_w > 0 else 100,
                "y": self.vis_pdf_h / 2 if self.vis_pdf_h > 0 else 100,
                "grid_mode": False,
                "grid_gap_x": 150,
                "grid_gap_y": 150,
                "img_obj": Image.open(f).convert("RGBA")
            }
            self.watermarks.append(wm)
            self.refresh_wm_list()
            self.wm_listbox.selection_set(len(self.watermarks)-1)
            self.on_wm_select()

    def add_text_watermark(self):
        wm = {
            "type": "text",
            "content": "测试水印",
            "scale": 1.0, 
            "opacity": 0.5,
            "angle": 0,
            "x": self.vis_pdf_w / 2 if self.vis_pdf_w > 0 else 100,
            "y": self.vis_pdf_h / 2 if self.vis_pdf_h > 0 else 100,
            "grid_mode": False,
            "grid_gap_x": 150,
            "grid_gap_y": 150,
            "color": "#FF0000",
            "font": "Arial"
        }
        self.watermarks.append(wm)
        self.refresh_wm_list()
        self.wm_listbox.selection_set(len(self.watermarks)-1)
        self.on_wm_select()

    def delete_selected_wm(self):
        if self.selected_wm_idx >= 0:
            del self.watermarks[self.selected_wm_idx]
            self.selected_wm_idx = -1
            self.refresh_wm_list()
            self.update_preview()

    def refresh_wm_list(self):
        self.wm_listbox.delete(0, tk.END)
        for i, wm in enumerate(self.watermarks):
            name = f"图: {os.path.basename(wm['path'])}" if wm['type'] == 'image' else f"文: {wm['content']}"
            self.wm_listbox.insert(tk.END, name)

    def on_wm_select(self, e=None):
        selection = self.wm_listbox.curselection()
        if not selection: return
        self.selected_wm_idx = selection[0]
        wm = self.watermarks[self.selected_wm_idx]
        
        # 加载到 UI 变量
        self.scale_var.set(wm['scale'])
        self.opacity_var.set(wm['opacity'])
        self.angle_var.set(wm['angle'])
        self.grid_mode_var.set(wm.get('grid_mode', False))
        self.grid_gap_x_var.set(wm.get('grid_gap_x', 150))
        self.grid_gap_y_var.set(wm.get('grid_gap_y', 150))
        
        if wm['type'] == 'image':
            self.watermark_path.set(wm['path'])
            self.frame_text_edit.pack_forget()
            self.frame_img_edit.pack(fill="x", before=self.lf_edit.winfo_children()[2])
        else:
            self.wm_text_var.set(wm['content'])
            self.wm_color_var.set(wm.get('color', '#FF0000'))
            self.wm_font_var.set(wm.get('font', 'Arial'))
            self.btn_color.config(bg=self.wm_color_var.get())
            self.frame_img_edit.pack_forget()
            self.frame_text_edit.pack(fill="x", before=self.lf_edit.winfo_children()[2])
        
        self.update_preview()

    def update_wm_from_ui(self):
        if self.selected_wm_idx < 0: return
        wm = self.watermarks[self.selected_wm_idx]
        wm['scale'] = self.scale_var.get()
        wm['opacity'] = self.opacity_var.get()
        wm['angle'] = self.angle_var.get()
        wm['grid_mode'] = self.grid_mode_var.get()
        wm['grid_gap_x'] = self.grid_gap_x_var.get()
        wm['grid_gap_y'] = self.grid_gap_y_var.get()
        if wm['type'] == 'text':
            wm['content'] = self.wm_text_var.get()
            wm['color'] = self.wm_color_var.get()
            wm['font'] = self.wm_font_var.get()
        
        # 更新列表显示名
        name = f"图: {os.path.basename(wm['path'])}" if wm['type'] == 'image' else f"文: {wm['content']}"
        self.wm_listbox.delete(self.selected_wm_idx)
        self.wm_listbox.insert(self.selected_wm_idx, name)
        self.wm_listbox.selection_set(self.selected_wm_idx)
        
        self.update_preview()

    def select_watermark(self):
        f = filedialog.askopenfilename(filetypes=[("Images", "*.png *.jpg *.jpeg")])
        if f:
            if self.selected_wm_idx >= 0 and self.watermarks[self.selected_wm_idx]['type'] == 'image':
                wm = self.watermarks[self.selected_wm_idx]
                wm['path'] = f
                wm['img_obj'] = Image.open(f).convert("RGBA")
                self.watermark_path.set(f)
                self.update_wm_from_ui()

    def set_pos_center(self):
        if self.selected_wm_idx >= 0:
            wm = self.watermarks[self.selected_wm_idx]
            wm['x'], wm['y'] = self.vis_pdf_w/2, self.vis_pdf_h/2
            self.update_preview()

    def set_pos_top_right(self):
        if self.selected_wm_idx >= 0:
            wm = self.watermarks[self.selected_wm_idx]
            margin = 50
            # 计算大致宽度（如果是图片）
            w = wm['img_obj'].width * wm['scale'] if wm['type'] == 'image' else 100
            h = wm['img_obj'].height * wm['scale'] if wm['type'] == 'image' else 30
            wm['x'] = self.vis_pdf_w - margin - w/2
            wm['y'] = self.vis_pdf_h - margin - h/2
            self.update_preview()

    def set_pos_top_left(self):
        if self.selected_wm_idx >= 0:
            wm = self.watermarks[self.selected_wm_idx]
        margin = 50
            # 计算大致宽度（如果是图片）
            w = wm['img_obj'].width * wm['scale'] if wm['type'] == 'image' else 100
            h = wm['img_obj'].height * wm['scale'] if wm['type'] == 'image' else 30
            wm['x'] = margin + w/2
            wm['y'] = self.vis_pdf_h - margin - h/2
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
                    self.output_dir_var.set(data.get("output_dir", "原文件目录"))
                    self.output_suffix_var.set(data.get("output_suffix", "_marked"))
                    self.all_templates = data.get("templates", {})
                    self.update_template_cb()
            except: pass

    def save_config(self):
        data = {
            "watermark_path": self.watermark_path.get(),
            "scale": self.scale_var.get(),
            "opacity": self.opacity_var.get(),
            "angle": self.angle_var.get(),
            "range_mode": self.range_mode_var.get(),
            "custom_range": self.custom_range_var.get(),
            "output_dir": self.output_dir_var.get(),
            "output_suffix": self.output_suffix_var.get(),
            "templates": getattr(self, 'all_templates', {})
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

    def hex_to_rgb(self, hex_color):
        hex_color = hex_color.lstrip('#')
        return tuple(int(hex_color[i:i+2], 16)/255.0 for i in (0, 2, 4))

    def get_pdf_font_name(self, font_family, text):
        # 检查是否包含中文字符，若包含则强制使用内置中文字库防止模糊或乱码
        has_chinese = any('\u4e00' <= char <= '\u9fff' for char in text)
        if has_chinese:
            return "china-s"
            
        mapping = {
            "Arial": "helv",
            "Helvetica": "helv",
            "Times New Roman": "tirom",
            "Courier New": "cour",
            "Verdana": "helv",
            "Georgia": "tirom"
        }
        return mapping.get(font_family, "helv")

    def process_files(self):
        # 预编译所有水印数据
        processed_wms = []
        for wm in self.watermarks:
            if wm['type'] == 'image':
                # 图片水印预处理：不再预先 resize，保留原始分辨率以防模糊
                wm_pil = wm['img_obj'].copy()
                ws, wa, wo = wm['scale'], wm['angle'], wm['opacity']
                
                # 使用高质量的双三次插值进行旋转
                wm_pil = wm_pil.rotate(wa, expand=True, resample=Image.Resampling.BICUBIC)
                
        r, g, b, a = wm_pil.split()
        wm_pil.putalpha(ImageEnhance.Brightness(a).enhance(wo))
                
        img_byte_arr = BytesIO()
                wm_pil.save(img_byte_arr, format='PNG', optimize=True)
                
                processed_wms.append({
                    "type": "image",
                    "data": img_byte_arr.getvalue(),
                    "display_w": wm_pil.width * ws,
                    "display_h": wm_pil.height * ws,
                    "x": wm['x'],
                    "y": wm['y'],
                    "grid_mode": wm.get('grid_mode', False),
                    "grid_gap_x": wm.get('grid_gap_x', 150),
                    "grid_gap_y": wm.get('grid_gap_y', 150)
                })
            else:
                # 文字水印 (使用 PyMuPDF 的 insert_text)
                processed_wms.append({
                    "type": "text",
                    "content": wm['content'],
                    "size": 30 * wm['scale'],
                    "opacity": wm['opacity'],
                    "angle": wm['angle'],
                    "color": self.hex_to_rgb(wm.get('color', '#FF0000')),
                    "font": self.get_pdf_font_name(wm.get('font', 'Arial'), wm['content']),
                    "x": wm['x'],
                    "y": wm['y'],
                    "grid_mode": wm.get('grid_mode', False),
                    "grid_gap_x": wm.get('grid_gap_x', 150),
                    "grid_gap_y": wm.get('grid_gap_y', 150)
                })

        mode = self.range_mode_var.get()
        output_dir = self.output_dir_var.get()
        suffix = self.output_suffix_var.get()
        custom = set()
        # ... (custom 范围逻辑保持不变)
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
                        page_w, page_h = page.rect.width, page.rect.height
                        
                        for pwm in processed_wms:
                            # 计算所有要绘制的位置
                            pos_list = []
                            if pwm.get('grid_mode'):
                                gx, gy = pwm['grid_gap_x'], pwm['grid_gap_y']
                                for ix in range(0, int(page_w + gx), int(gx)):
                                    for iy in range(0, int(page_h + gy), int(gy)):
                                        pos_list.append((ix, iy))
                            else:
                                pos_list.append((pwm['x'], pwm['y']))

                            for px, py in pos_list:
                                if pwm['type'] == 'image':
                                    # 保持原始分辨率的高质量插入
                                    rect_x0 = px - pwm['display_w']/2
                                    rect_y0 = (page_h - py) - pwm['display_h']/2
                                    page.insert_image(fitz.Rect(rect_x0, rect_y0, rect_x0 + pwm['display_w'], rect_y0 + pwm['display_h']), 
                                                   stream=pwm['data'])
                                else:
                                    # 插入矢量文字水印
                                    page.insert_text((px, page_h - py), 
                                                   pwm['content'], 
                                                   fontsize=pwm['size'], 
                                                   color=pwm['color'], 
                                                   fontname=pwm['font'],
                                                   rotate=pwm['angle'],
                                                   fill_opacity=pwm['opacity'])
                
                # 计算保存路径
                base_name = os.path.basename(os.path.splitext(path)[0])
                final_name = base_name + suffix + ".pdf"
                
                if output_dir == "原文件目录":
                    out_dir = os.path.dirname(path)
                    save_path = os.path.join(out_dir, final_name)
                else:
                    out_dir = output_dir
                    save_path = os.path.join(out_dir, final_name)
                
                # 记录最后一次导出的目录
                self.last_output_dir = out_dir
                
                doc.save(save_path)
                doc.close()
                self.last_output_path = save_path # 记录最后生成的文件路径
                count += 1
            except Exception as e: print(f"失败: {e}")
            self.progress["value"] = (i+1)/len(self.pdf_files)*100
        
        self.status_var.set("处理完成")
        self.btn_run.config(state="normal")
        messagebox.showinfo("完成", f"成功处理 {count} 个文件")

if __name__ == "__main__":
    # --- Windows 高分屏 (DPI) 适配 ---
    try:
        from ctypes import windll
        windll.shcore.SetProcessDpiAwareness(1)
    except:
        pass
        
    root = tk.Tk(); app = AdvancedWatermarkApp(root)
    root.mainloop()