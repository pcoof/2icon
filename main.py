import customtkinter as ctk
import os
import threading
import tempfile
from tkinter import filedialog, messagebox
from tkinterdnd2 import DND_FILES, TkinterDnD
from PIL import Image
import struct
import ctypes

# 导入SVG处理相关库（在需要使用处再局部导入，避免顶层硬依赖）
try:
    import win32api
    import win32con
    import win32com.client
    import pythoncom
    HAS_WIN32 = True
except ImportError:
    HAS_WIN32 = False

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    HAS_TKINTERDND2 = True
except ImportError:
    HAS_TKINTERDND2 = False


# 常用常量（消除散落各处的重复字面量）
ICON_PRESET_SIZES = [16, 32, 48, 64, 96, 128, 256, 512]
BIT_DEPTH_INFO = {
    "RGBA": "32位RGBA",
    "RGB": "24位RGB",
    "P": "8位索引",
    "L": "8位灰度",
    "1": "1位黑白",
}


def check_svg_support():
    """检查系统上可用的SVG处理方法（均不依赖Cairo）"""
    support_info = {
        "svglib": False,   # 用于SVG解析/校验
        "resvg": False,    # 用于SVG光栅化（Rust实现，非Cairo）
    }
    
    # 检查svglib（SVG解析/校验）
    try:
        from svglib.svglib import svg2rlg
        support_info["svglib"] = True
        print("svglib 可用（用于SVG解析）")
    except ImportError:
        print("警告: svglib 未安装")
    
    # 检查resvg（SVG光栅化，纯Rust实现，不依赖Cairo）
    try:
        import resvg_py
        support_info["resvg"] = True
        print("resvg 可用（非Cairo SVG光栅化）")
    except ImportError:
        print("警告: resvg-py 未安装，SVG 将无法处理")
    
    return support_info

class ImageToIcoConverter(TkinterDnD.Tk):
    def __init__(self):
        super().__init__()
        self.title("图像/文件图标提取工具")
        self.geometry("800x600")
        self.minsize(1000, 700)
        
        # 设置窗口与任务栏图标
        self._set_window_icon()
        
        # 检查SVG支持
        self.svg_support_info = check_svg_support()
        svg_methods = [method for method, supported in self.svg_support_info.items() if supported]
        if svg_methods:
            print(f"SVG支持: 可用方法 - {', '.join(svg_methods)}")
        else:
            print("警告: 系统上没有可用的SVG处理方法。SVG文件将无法处理。")
        
        # 尝试启用拖放功能
        self.drag_drop_supported = False
        if HAS_TKINTERDND2 and hasattr(self, 'drop_target_register'):
            try:
                # 注册拖放目标
                self.drop_target_register(DND_FILES)
                self.dnd_bind("<<Drop>>", self.on_drop)
                self.drag_drop_supported = True
            except Exception as e:
                print(f"警告: 无法启用拖放功能: {e}")
        
        # 创建临时目录
        self.temp_dir = os.path.join(tempfile.gettempdir(), "icon_extractor_temp")
        os.makedirs(self.temp_dir, exist_ok=True)
        # 主题设置
        self.is_dark_theme = True
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")
        # 变量初始化
        self.input_path = ctk.StringVar(value="")
        self.output_path = ctk.StringVar(value="")
        self.status = ctk.StringVar(value="就绪")
        self.icon_list = []
        self.icon_buttons = []
        self.selected_icon_index = ctk.IntVar(value=0)
        self.selected_size = ctk.StringVar(value="64")
        # 独立的输出宽高，支持非方形（ICO会自动转为方形）
        self.selected_width = ctk.IntVar(value=64)
        self.selected_height = ctk.IntVar(value=64)
        self.quality_filter = ctk.BooleanVar(value=True)
        
        # 创建UI
        self.create_widgets()
        # 绑定关闭事件
        self.protocol("WM_DELETE_WINDOW", self.on_closing)
        # 初始化标题栏颜色（需在窗口创建后）
        self.update_idletasks()
        self._apply_titlebar_theme(self.is_dark_theme)

    def on_drop(self, event):
        if not self.drag_drop_supported:
            return
        # 获取拖放的文件路径
        file_path = event.data.strip()
        # 处理Windows路径中的引号
        if file_path.startswith('{') and file_path.endswith('}'):
            file_path = file_path[1:-1]
        # 更新输入路径
        self.input_path.set(file_path)
        # 自动生成输出路径
        output_file = os.path.splitext(os.path.basename(file_path))[0] + "." + self.selected_format.get().lower()
        output_dir = os.path.dirname(file_path)
        self.output_path.set(os.path.join(output_dir, output_file))
        # 清除之前的图标列表
        self.icon_list = []
        self.clear_icon_frame()
        self.update_preview()
        # 自动提取图标
        self.extract_icons()

    def browse_input(self, file_path=None):
        if not file_path:
            file_path = filedialog.askopenfilename(
                filetypes=[
                    ("所有支持的文件", "*.png *.jpg *.jpeg *.webp *.exe *.dll *.lnk *.svg"),
                    ("图像文件", "*.png *.jpg *.jpeg *.webp"),
                    ("Windows可执行文件", "*.exe *.dll"),
                    ("快捷方式文件", "*.lnk"),
                    ("所有文件", "*.*"),
                ]
            )
        if file_path:
            self.input_path.set(file_path)
            # 自动生成输出路径
            output_file = os.path.splitext(os.path.basename(file_path))[0] + "." + self.selected_format.get().lower()
            output_dir = os.path.dirname(file_path)
            self.output_path.set(os.path.join(output_dir, output_file))
            # 清除之前的图标列表
            self.icon_list = []
            self.clear_icon_frame()
            self.update_preview()
            # 自动提取图标
            self.extract_icons()

    def create_widgets(self):
        # 创建主框架
        main_frame = ctk.CTkFrame(self)
        main_frame.pack(fill="both", expand=True)
        # 创建文件路径选择区域
        path_frame = ctk.CTkFrame(main_frame)
        path_frame.pack(fill="x")
        # 输入文件选择
        input_frame = ctk.CTkFrame(path_frame)
        input_frame.pack(fill="x", padx=10, pady=(10, 5))
        ctk.CTkLabel(input_frame, text="输入文件:", font=ctk.CTkFont(size=12)).pack(
            side="left", padx=(10, 5)
        )
        self.input_entry = ctk.CTkEntry(
            input_frame,
            textvariable=self.input_path,
            placeholder_text="选择图像文件或可执行文件...",
        )
        self.input_entry.pack(side="left", fill="x", expand=True, padx=5)
        browse_input_btn = ctk.CTkButton(
            input_frame, text="浏览", command=self.browse_input, width=80
        )
        browse_input_btn.pack(side="right", padx=(5, 10))
        # 输出文件选择
        output_frame = ctk.CTkFrame(path_frame)
        output_frame.pack(fill="x", padx=10, pady=(5, 15))
        ctk.CTkLabel(output_frame, text="输出文件:", font=ctk.CTkFont(size=12)).pack(
            side="left", padx=(10, 5)
        )
        self.output_entry = ctk.CTkEntry(
            output_frame,
            textvariable=self.output_path,
            placeholder_text="保存ICO文件路径...",
        )
        self.output_entry.pack(side="left", fill="x", expand=True, padx=5)
        browse_output_btn = ctk.CTkButton(
            output_frame, text="浏览", command=self.browse_output, width=80
        )
        browse_output_btn.pack(side="right", padx=(5, 10))
        # 创建内容区域
        content_frame = ctk.CTkFrame(main_frame)
        content_frame.pack(fill="both", expand=True)
        # 左侧 - 图标列表
        left_frame = ctk.CTkFrame(content_frame)
        left_frame.pack(side="left", fill="both", expand=True, padx=(10, 10), pady=0)
        ctk.CTkLabel(
            left_frame, text="图标列表", font=ctk.CTkFont(size=14, weight="bold")
        ).pack(pady=(5, 10))
        # 创建滚动框架用于图标显示
        self.icon_scrollable_frame = ctk.CTkScrollableFrame(
            left_frame, width=400, height=300
        )
        self.icon_scrollable_frame.pack(fill="both", expand=True, padx=15, pady=(0, 15))
        # 右侧 - 控制面板
        right_frame = ctk.CTkFrame(content_frame, width=450)  # 设置固定宽度
        right_frame.pack(side="right", fill="y", padx=(0, 10), pady=0)
        right_frame.pack_propagate(False)  # 防止框架根据内容调整大小
        # 图标质量过滤（已在 __init__ 中定义 self.quality_filter）
        quality_checkbox = ctk.CTkCheckBox(
            right_frame,
            text="仅显示高质量图标",
            variable=self.quality_filter,
            command=self.on_quality_filter_change,
        )
        quality_checkbox.pack(pady=(5))
        # 图标尺寸选择
        size_frame = ctk.CTkFrame(right_frame)
        size_frame.pack(fill="x", padx=5, pady=0)
        # 创建尺寸选择按钮
        sizes = ICON_PRESET_SIZES
        self.size_buttons = []
        for i, size in enumerate(sizes):
            row, col = i // 4, i % 4
            btn = ctk.CTkRadioButton(
                size_frame,
                text=f"{size}x{size}",
                variable=self.selected_size,
                value=str(size),
                command=lambda s=size: self.on_preset_size_selected(s),
            )
            btn.grid(row=row, column=col, sticky="w", padx=5, pady=5)
            self.size_buttons.append(btn)
        
        # 添加自定义尺寸输入
        custom_size_frame = ctk.CTkFrame(right_frame)
        custom_size_frame.pack(fill="x", padx=5, pady=(10, 0))
        ctk.CTkLabel(custom_size_frame, text="自定义尺寸: ").pack(side="left", padx=5, pady=5)
        self.custom_size_var = ctk.StringVar(value="")
        custom_size_entry = ctk.CTkEntry(custom_size_frame, textvariable=self.custom_size_var, width=60)
        custom_size_entry.pack(side="left", padx=5)
        ctk.CTkLabel(custom_size_frame, text="x").pack(side="left")
        self.custom_size_var2 = ctk.StringVar(value="")
        custom_size_entry2 = ctk.CTkEntry(custom_size_frame, textvariable=self.custom_size_var2, width=60)
        custom_size_entry2.pack(side="left", padx=5)
        
        # 添加应用自定义尺寸按钮
        apply_custom_size_btn = ctk.CTkButton(
            custom_size_frame,
            text="应用",
            command=self.apply_custom_size,
            width=60
        )
        apply_custom_size_btn.pack(side="left", padx=5)
        
        # 添加默认尺寸按钮
        default_size_btn = ctk.CTkButton(
            custom_size_frame,
            text="默认",
            command=self.set_default_size,
            width=60
        )
        default_size_btn.pack(side="right", padx=5)
        
        # 保存按钮（先pack，固定在底部，然后再pack内容区域填充剩余空间）
        # 按钮行（保存 + 复制），固定在底部
        button_row = ctk.CTkFrame(right_frame, fg_color="transparent")
        button_row.pack(side="bottom", pady=10, padx=15, fill="x")
        self.save_button = ctk.CTkButton(
            button_row,
            text="保存图标",
            command=self.save_icon,
            height=45,
            font=ctk.CTkFont(size=15, weight="bold"),
            fg_color=("#2b7a2b", "#2b9a2b"),  # 绿色，更显眼
            hover_color=("#1e5e1e", "#1e7e1e"),
            text_color="white",
            corner_radius=8,
        )
        self.save_button.pack(side="left", fill="x", expand=True, padx=(0, 5))
        self.copy_button = ctk.CTkButton(
            button_row,
            text="复制",
            command=self.copy_icon,
            height=45,
            font=ctk.CTkFont(size=15, weight="bold"),
            fg_color=("#2563eb", "#1d4ed8"),
            hover_color=("#1e40af", "#1e3a8a"),
            text_color="white",
            corner_radius=8,
        )
        self.copy_button.pack(side="right", fill="x", expand=True, padx=(5, 0))
        
        # 创建一个容器框架来容纳预览区域和输出格式选择（填充保存按钮上方的剩余空间）
        right_content_frame = ctk.CTkFrame(right_frame, fg_color="transparent")
        right_content_frame.pack(fill="both", expand=True, padx=0, pady=(5, 0))
        
        # 预览区域（上方）
        self.preview_frame = ctk.CTkFrame(right_content_frame, width=200, height=200)
        self.preview_frame.pack(pady=(5, 8))
        self.preview_frame.pack_propagate(False)
        
        # 输出格式选择（预览区域下方）
        format_frame_container = ctk.CTkFrame(right_content_frame)
        format_frame_container.pack(padx=10, pady=(0, 3), fill="x")
        
        format_label = ctk.CTkLabel( format_frame_container, text="输出格式", font=ctk.CTkFont(size=13, weight="bold") )
        format_label.pack(pady=(3, 3))
        self.selected_format = ctk.StringVar(value="ICO")
        format_frame = ctk.CTkFrame(format_frame_container, fg_color="transparent")
        format_frame.pack(fill="x", padx=5, pady=(0, 3))
        
        formats = ["ICO", "PNG", "WEBP"]
        self.format_buttons = []
        for i, fmt in enumerate(formats):
            btn = ctk.CTkRadioButton(
                format_frame,
                text=fmt,
                variable=self.selected_format,
                value=fmt,
                command=self.on_format_change,
            )
            btn.pack(side="left", anchor="center", expand=True, padx=5, pady=2)
            self.format_buttons.append(btn)

        # 压缩比例滑块（位于输出格式区域底部）
        quality_header = ctk.CTkFrame(format_frame_container, fg_color="transparent")
        quality_header.pack(fill="x", padx=5, pady=(4, 0))
        ctk.CTkLabel(
            quality_header, text="压缩质量", font=ctk.CTkFont(size=12)
        ).pack(side="left", padx=5)
        self.quality_label = ctk.CTkLabel(
            quality_header, text="95%", font=ctk.CTkFont(size=12), width=45
        )
        self.quality_label.pack(side="right", padx=5)
        self.quality_var = ctk.DoubleVar(value=95)
        self.compress_slider = ctk.CTkSlider(
            format_frame_container,
            from_=10,
            to=100,
            number_of_steps=90,
            variable=self.quality_var,
            command=self._update_quality_label,
        )
        self.compress_slider.pack(fill="x", padx=15, pady=(2, 8))

        # 图标信息显示
        self.info_label = ctk.CTkLabel( right_content_frame, text="", font=ctk.CTkFont(size=10), text_color="gray" )
        self.info_label.pack(pady=(2, 0))
        # 状态栏
        status_frame = ctk.CTkFrame(main_frame)
        status_frame.pack(fill="x", padx=10, pady=(10))
        self.status_label = ctk.CTkLabel(
            status_frame, textvariable=self.status, font=ctk.CTkFont(size=12)
        )
        self.status_label.pack(side="left", padx=15, pady=10)
        # 进度条（初始隐藏）
        self.progress_bar = ctk.CTkProgressBar(status_frame, width=200)
        self.progress_bar.pack(side="left", padx=10, pady=10)
        self.progress_bar.pack_forget()  # 初始隐藏
        # 关于按钮
        about_button = ctk.CTkButton(
            status_frame,
            text="关于",
            command=self.show_about,
            width=40,
            height=30,
            font=ctk.CTkFont(size=14),
        )
        about_button.pack(side="right", padx=5, pady=10)
        # 主题切换按钮
        # 黑/白 主题开关
        self.theme_switch = ctk.CTkSwitch(
            status_frame,
            text="深色" if self.is_dark_theme else "浅色",
            command=self.toggle_theme,
            width=80,
            font=ctk.CTkFont(size=13),
            switch_width=46,
            switch_height=24,
        )
        self.theme_switch.pack(side="right", padx=5, pady=10)
        if self.is_dark_theme:
            self.theme_switch.select()
        else:
            self.theme_switch.deselect()
        # 版权信息
        copyright_label = ctk.CTkLabel(
            status_frame,
            text="© 2025 图标提取工具",
            font=ctk.CTkFont(size=10),
            text_color="gray",
        )
        copyright_label.pack(side="right", padx=5, pady=10)

    def toggle_theme(self):
        """切换黑/白主题（由开关状态决定）"""
        if self.theme_switch.get():
            ctk.set_appearance_mode("dark")
            self.is_dark_theme = True
            self.theme_switch.configure(text="深色")
        else:
            ctk.set_appearance_mode("light")
            self.is_dark_theme = False
            self.theme_switch.configure(text="浅色")
        # 让原生标题栏颜色跟随主题
        self._apply_titlebar_theme(self.is_dark_theme)

    def show_about(self):
        """显示关于对话框"""
        about_text = """图像/文件图标提取工具 v1.0
功能特性：
• 支持多种图像格式 (PNG, JPG, WEBP, SVG)
• 从可执行文件提取图标 (EXE, DLL)
• 从快捷方式提取图标 (LNK)
• 智能图标质量排序
• 高质量图标优先显示
• 多种输出尺寸选择
• 现代化界面设计
• 明暗主题切换
• 自动清理临时文件
界面按钮说明：
主题 - 切换明暗主题
清理 - 手动清理临时文件
关于 - 显示此帮助信息

© 2025 图标提取工具
使用 CustomTkinter 构建"""
        messagebox.showinfo("关于", about_text)

    def on_quality_filter_change(self):
        """质量过滤选项改变时重新提取图标"""
        input_file = self.input_path.get()
        if input_file and os.path.exists(input_file):
            # 重新提取图标以应用新的过滤设置
            self.extract_icons()

    def cleanup_temp_dir(self):
        """清理临时目录"""
        try:
            if os.path.exists(self.temp_dir):
                import shutil
                import time
                # 尝试多次清理，处理文件被占用的情况
                max_attempts = 3
                for attempt in range(max_attempts):
                    try:
                        # 删除临时目录及其所有内容
                        shutil.rmtree(self.temp_dir)
                        print(f"已清理临时目录: {self.temp_dir}")
                        break
                    except Exception as e:
                        if attempt < max_attempts - 1:
                            print(f"清理尝试 {attempt + 1} 失败，重试中...")
                            time.sleep(0.5)  # 等待0.5秒后重试
                        else:
                            # 最后一次尝试失败，尝试清理单个文件
                            print(f"无法删除整个目录，尝试清理单个文件...")
                            self._cleanup_individual_files()
        except Exception as e:
            print(f"清理临时目录失败: {e}")

    def _cleanup_individual_files(self):
        """逐个清理临时文件（兜底：处理被占用/只读的文件与子目录）"""
        try:
            for root, dirs, files in os.walk(self.temp_dir, topdown=False):
                for file in files:
                    file_path = os.path.join(root, file)
                    try:
                        # 去除只读属性后再删除（Windows 常见）
                        try:
                            os.chmod(file_path, 0o777)
                        except Exception:
                            pass
                        os.remove(file_path)
                    except:
                        pass  # 忽略无法删除的文件
                # 删除已清空的子目录
                for d in dirs:
                    try:
                        os.rmdir(os.path.join(root, d))
                    except:
                        pass
            print("已清理可清理的临时文件")
        except Exception as e:
            print(f"单个文件清理失败: {e}")

    def on_closing(self):
        """程序关闭时的清理工作"""
        print("程序正在关闭，清理临时文件...")
        # 先关闭仍被 PIL 打开的图标文件句柄：PIL 懒加载会保留文件句柄，
        # 直到 img.close() 或 GC；Windows 下被占用会导致 rmtree 删除失败
        for img in list(getattr(self, "icon_list", [])):
            try:
                img.close()
            except Exception:
                pass
        self.icon_list = []
        self.icon_buttons = []
        self.cleanup_temp_dir()
        self.destroy()

    def _set_window_icon(self):
        """设置窗口标题栏与任务栏图标为根目录的 logo.ico（两者统一）"""
        logo_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "logo.ico")
        if not os.path.exists(logo_path):
            print(f"提示: 未找到图标文件 {logo_path}，跳过设置窗口图标")
            return
        # 1) 用 .ico 设置任务栏/标题栏大图标（含多尺寸，Windows 直接用）
        try:
            self.iconbitmap(logo_path)
        except Exception as e:
            print(f"警告: 设置窗口图标(ico)失败: {e}")
        # 2) 额外用 iconphoto 设置标题栏小图标：Tk 对标题栏小图标
        #    有时只认 PhotoImage，补上可确保标题栏与任务栏图标统一、清晰
        try:
            from PIL import Image, ImageTk
            img = Image.open(logo_path).convert("RGBA")
            photo = ImageTk.PhotoImage(img.resize((48, 48), Image.LANCZOS))
            self.iconphoto(True, photo)
            self._icon_photo = photo  # 保留引用，防止被垃圾回收
        except Exception as e:
            print(f"提示: 设置标题栏小图标(iconphoto)失败: {e}")

    def _apply_titlebar_theme(self, dark):
        """让原生标题栏颜色跟随主题（Windows 沉浸式深色模式）"""
        if not hasattr(ctypes, "windll"):
            return
        try:
            user32 = ctypes.windll.user32
            dwmapi = ctypes.windll.dwmapi
            # 取得真正的最顶层窗口句柄（部分框架外层还有容器窗口）
            hwnd = int(self.winfo_id())
            user32.GetAncestor.argtypes = [ctypes.wintypes.HWND, ctypes.wintypes.UINT]
            user32.GetAncestor.restype = ctypes.wintypes.HWND
            root_hwnd = user32.GetAncestor(hwnd, 2)  # GA_ROOT = 2
            if root_hwnd:
                hwnd = root_hwnd

            dwmapi.DwmSetWindowAttribute.argtypes = [
                ctypes.wintypes.HWND,
                ctypes.wintypes.DWORD,
                ctypes.POINTER(ctypes.wintypes.BOOL),
                ctypes.wintypes.DWORD,
            ]
            dwmapi.DwmSetWindowAttribute.restype = ctypes.c_int32  # HRESULT
            value = ctypes.wintypes.BOOL(1 if dark else 0)
            # 20 = DWMWA_USE_IMMERSIVE_DARK_MODE（Win10 17763+），
            # 19 用于更早的预览版本，依次尝试
            ok = False
            for attr in (20, 19):
                hr = dwmapi.DwmSetWindowAttribute(
                    hwnd, attr, ctypes.byref(value), ctypes.sizeof(value)
                )
                if hr == 0:
                    ok = True
                    break
                else:
                    print(f"提示: DwmSetWindowAttribute(attr={attr}) 返回 HRESULT={hr}")
            if not ok:
                print("提示: 当前系统不支持沉浸式深色标题栏（不影响功能）")
                return

            # 强制非客户区（标题栏）重绘，确保颜色立即生效
            user32.SetWindowPos.argtypes = [
                ctypes.wintypes.HWND, ctypes.wintypes.HWND,
                ctypes.wintypes.INT, ctypes.wintypes.INT,
                ctypes.wintypes.INT, ctypes.wintypes.INT, ctypes.wintypes.UINT,
            ]
            user32.SetWindowPos.restype = ctypes.wintypes.BOOL
            SWP_NOMOVE = 0x0002
            SWP_NOSIZE = 0x0001
            SWP_FRAMECHANGED = 0x0020
            user32.SetWindowPos(
                hwnd, 0, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_FRAMECHANGED
            )
        except Exception as e:
            print(f"提示: 设置标题栏颜色失败（不影响功能）: {e}")

    def browse_output(self):
        file_path = filedialog.asksaveasfilename(
            defaultextension=".ico",
            filetypes=[("ICO文件", "*.ico"), ("所有文件", "*.*")],
        )
        if file_path:
            self.output_path.set(file_path)

    def extract_icons(self):
        input_file = self.input_path.get()
        if not input_file:
            messagebox.showerror("错误", "请选择输入文件")
            return
        # 在新线程中执行提取，避免界面卡顿
        threading.Thread(target=self._extract_thread, daemon=True).start()
    def _extract_thread(self):
        try:
            self.status.set("提取中...")
            # 显示进度条
            self.after(
                0, lambda: self.progress_bar.pack(side="left", padx=10, pady=10)
            )
            self.after(0, lambda: self.progress_bar.set(0.1))
            input_file = self.input_path.get()
            ext = os.path.splitext(input_file)[1].lower()
            if ext in [".png", ".jpg", ".jpeg", ".webp"]:
                # 处理图像文件
                img = Image.open(input_file)

                # 转换为RGBA
                img = img.convert("RGBA")
                self.icon_list = [img]
            elif ext == ".svg":
                # 处理SVG文件
                svg_methods = [method for method, supported in self.svg_support_info.items() if supported]
                if not svg_methods:
                    messagebox.showerror("错误", "系统上没有可用的SVG处理方法。请安装 resvg-py（非Cairo的SVG光栅化）")
                    self.status.set("就绪")
                    return
                
                try:
                    # 创建临时PNG文件路径
                    temp_png = os.path.join(self.temp_dir, "temp_svg_converted.png")
                    
                    # 检查文件是否为有效的SVG
                    with open(input_file, 'r', encoding='utf-8') as f:
                        svg_content = f.read(1000)  # 只读取前1000个字符进行检查
                        if not ('<svg' in svg_content and 'xmlns' in svg_content):
                            raise ValueError("文件不是有效的SVG格式")
                    
                    # 根据可用的SVG处理方法选择最佳路径
                    svg_converted = False
                    errors = []
                    
                    # 方法1: svglib解析 + resvg光栅化（均不依赖Cairo）
                    if self.svg_support_info.get("resvg") and not svg_converted:
                        try:
                            import resvg_py
                            # 可选：用svglib解析以校验SVG有效性
                            try:
                                from svglib.svglib import svg2rlg
                                drawing = svg2rlg(input_file)
                                if drawing is None:
                                    raise ValueError("SVG解析结果无效 (svglib)")
                            except Exception as parse_err:
                                print(f"svglib解析警告（将由resvg重试）: {parse_err}")
                            # 使用resvg将SVG光栅化为PNG（Rust实现，非Cairo）
                            png_bytes = resvg_py.svg_to_bytes(
                                svg_path=input_file,
                                background="white",
                            )
                            with open(temp_png, "wb") as f:
                                f.write(png_bytes)
                            # 检查生成的PNG是否存在
                            if not os.path.exists(temp_png) or os.path.getsize(temp_png) == 0:
                                raise ValueError("SVG渲染失败，未生成有效的PNG文件")
                            
                            # 使用PIL打开生成的PNG
                            img = Image.open(temp_png)
                            img = img.convert("RGBA")
                            self.icon_list = [img]
                            svg_converted = True
                            print("使用resvg成功处理SVG文件")
                        except Exception as e:
                            errors.append(f"方法1 (resvg) 失败: {str(e)}")
                    
                    # 如果所有方法都失败，抛出错误
                    if not svg_converted:
                        error_msg = "\n".join(errors)
                        raise Exception(f"所有SVG处理方法都失败:\n{error_msg}")
                    
                    # 删除临时文件
                    try:
                        if os.path.exists(temp_png):
                            os.remove(temp_png)
                    except:
                        pass
                except Exception as e:
                    error_msg = str(e)
                    solution_msg = ""
                    
                    if "文件不是有效的SVG格式" in error_msg:
                        solution_msg = "\n\n请确保文件是有效的SVG格式，包含正确的XML命名空间。"
                    elif "SVG解析结果无效" in error_msg:
                        solution_msg = "\n\n该SVG文件可能包含不支持的特性或格式错误。"
                    
                    messagebox.showerror("错误", f"SVG处理失败: {error_msg}{solution_msg}")
                    self.status.set("就绪")
                    return
            elif ext == ".lnk":
                # 处理快捷方式文件
                try:
                    self.icon_list = self._extract_from_lnk(input_file)
                except Exception as e:
                    messagebox.showerror("错误", f"从快捷方式提取图标失败:\n{str(e)}")
                    self.status.set("就绪")
                    return
            elif ext in [".exe", ".dll"]:
                # 处理可执行文件和DLL
                try:
                    self.icon_list = self._extract_from_exe_dll(input_file)
                except Exception as e:
                    messagebox.showerror("错误", f"从文件提取图标失败:\n{str(e)}")
                    self.status.set("就绪")
                    return

            else:
                messagebox.showerror("错误", "不支持的文件类型")
                self.status.set("就绪")
                return
            # 更新图标列表
            if self.icon_list:
                self.after(0, lambda: self.progress_bar.set(0.8))
                self.after(0, self._update_icon_grid)
                self.after(0, lambda: self.progress_bar.set(1.0))
                self.status.set(f"{len(self.icon_list)}个图标 | 完成")
            else:
                self.status.set("未找到图标")
        except Exception as e:
            self.status.set(f"提取失败: {str(e)}")
            messagebox.showerror("错误", f"提取过程中出现错误:\n{str(e)}")
        finally:
            # 隐藏进度条
            self.after(1000, lambda: self.progress_bar.pack_forget())



    def clear_icon_frame(self):
        # 清除图标框架中的所有按钮
        for widget in self.icon_scrollable_frame.winfo_children():
            widget.destroy()
        self.icon_buttons = []

    def _update_icon_grid(self):
        # 清空图标框架
        self.clear_icon_frame()
        # 显示图标网格
        icon_size = 72  # 图标显示大小
        padding = 10    # 图标间距
        # 根据可用宽度自适应列数（限制 3~8 列，避免过疏或过密）
        cell = icon_size + 46  # 单格目标宽度（图标 + 文字 + 边距）
        avail = self.icon_scrollable_frame.winfo_width()
        columns = max(3, min(8, avail // cell)) if avail > cell else 4
        for i, img in enumerate(self.icon_list):
            # 计算行列位置
            row = i // columns
            col = i % columns
            # 调整图像大小
            img_copy = img.copy()
            img_copy.thumbnail((icon_size, icon_size), Image.LANCZOS)
            # 创建CTkImage
            ctk_image = ctk.CTkImage(
                light_image=img_copy, dark_image=img_copy, size=(icon_size, icon_size)
            )
            # 获取图标详细信息
            width, height = img.size
            # 获取图像模式信息
            bit_info = BIT_DEPTH_INFO.get(img.mode, img.mode)

            icon_size_text = f"{width}x{height}\n{bit_info}"
            # 创建图标容器框架
            icon_container = ctk.CTkFrame(self.icon_scrollable_frame)
            icon_container.grid(
                row=row, column=col, padx=padding, pady=padding, sticky="nsew"
            )

            # 创建按钮
            btn = ctk.CTkButton(
                icon_container,
                image=ctk_image,
                text=icon_size_text,
                compound="top",
                command=lambda idx=i: self.select_icon(idx),
                width=icon_size + 24,
                height=icon_size + 54,  # 增加高度以适应两行文本
                font=ctk.CTkFont(size=10),
                text_color=("gray10", "white"),  # 日间深色/夜间白色，确保可见
            )
            btn.ctk_image = ctk_image  # 保持引用，防止被垃圾回收
            btn.pack(padx=5, pady=5)

            # 存储按钮引用
            self.icon_buttons.append(btn)

        # 配置网格权重
        for col in range(columns):
            self.icon_scrollable_frame.grid_columnconfigure(col, weight=1)

        # 选中第一个图标
        if self.icon_list:
            self.select_icon(0)
            # 设置默认尺寸为选中图标的原始尺寸
            self.after(100, self.set_default_size)

    def _extract_from_lnk(self, lnk_path):
        # 检查是否支持快捷方式处理
        if not HAS_WIN32:
            raise Exception("系统缺少win32模块，无法处理快捷方式文件。请确保已安装pywin32。")
        
        # 从快捷方式提取图标
        pythoncom.CoInitialize()
        try:
            # 获取快捷方式指向的目标
            shell = win32com.client.Dispatch("WScript.Shell")
            shortcut = shell.CreateShortCut(lnk_path)
            target_path = shortcut.TargetPath

            # 从目标文件提取图标
            return self._extract_from_exe_dll(target_path)
        except Exception as e:
            raise Exception(f"{str(e)}")
        finally:
            pythoncom.CoUninitialize()

    def _extract_from_exe_dll(self, file_path):
        # 检查是否支持EXE/DLL处理
        if not HAS_WIN32:
            raise Exception("系统缺少win32模块，无法处理可执行文件。请确保已安装pywin32。")
        
        # 从EXE或DLL提取图标
        icons = []
        icon_info_list = []  # 存储图标信息用于排序
        try:
            # 加载文件
            hModule = win32api.LoadLibraryEx(
                file_path, 0, win32con.LOAD_LIBRARY_AS_DATAFILE
            )
            # 枚举资源
            resource_types = win32api.EnumResourceTypes(hModule)
            # print(f"找到资源类型: {len(resource_types)} 个")
            # print(f"资源类型列表: {resource_types}")
            # 检查是否有单独的图标资源（不在组中）
            individual_icon_data = []
            if win32con.RT_ICON in resource_types:
                # print("找到单独图标资源")
                try:
                    individual_icons = win32api.EnumResourceNames(
                        hModule, win32con.RT_ICON
                    )
                    # print(f"单独图标数量: {len(individual_icons)}")
                    # 尝试直接加载这些单独的图标
                    for icon_id in individual_icons[:15]:  # 增加处理数量
                        try:
                            icon_data = win32api.LoadResource(
                                hModule, win32con.RT_ICON, icon_id
                            )
                            if len(icon_data) > 40:  # 确保有足够的数据
                                individual_icon_data.append((icon_id, icon_data))
                                # print(f"  加载图标 {icon_id}: {len(icon_data)} 字节")
                        except Exception as e:
                            # print(f"  加载图标 {icon_id} 失败: {e}")
                            continue
                except:
                    pass
                    # print("无法枚举单独图标")
            if win32con.RT_GROUP_ICON in resource_types:
                # print("找到图标组资源")
                # 枚举图标组
                icon_groups = win32api.EnumResourceNames(
                    hModule, win32con.RT_GROUP_ICON
                )
                # print(f"📁 找到 {len(icon_groups)} 个图标组: {icon_groups}")
                for icon_group in icon_groups:
                    # 获取图标组数据
                    icon_group_data = win32api.LoadResource(
                        hModule, win32con.RT_GROUP_ICON, icon_group
                    )
                    # 解析图标组头
                    if len(icon_group_data) < 6:
                        continue  # 数据不足，跳过
                    header = struct.unpack("<HHH", icon_group_data[:6])
                    num_icons = header[1]

                    # 收集所有图标信息
                    group_icons = []
                    # print(f"  📁 图标组 {icon_group}: {num_icons} 个图标")
                    for i in range(num_icons):
                        # 图标目录项结构
                        entry_offset = 6 + i * 14
                        if entry_offset + 14 > len(icon_group_data):
                            continue  # 数据不足，跳过
                        entry_data = icon_group_data[entry_offset : entry_offset + 14]
                        # 检查数据长度是否足够
                        if len(entry_data) < 14:
                            continue  # 数据不足，跳过
                        try:
                            (
                                width,
                                height,
                                colors,
                                _,
                                planes,
                                bit_count,
                                bytes_in_res,
                                icon_id,
                            ) = struct.unpack(
                                "<BBBBHHLL", entry_data + b"\x00\x00"
                            )  # 补足16字节

                            # 处理0值（表示256像素）
                            actual_width = 256 if width == 0 else width
                            actual_height = 256 if height == 0 else height

                            # 计算图标质量分数（优先高分辨率、高色深）
                            quality_score = (actual_width * actual_height) + (
                                bit_count * 10
                            )
                            icon_info = {
                                "width": actual_width,
                                "height": actual_height,
                                "colors": colors,
                                "planes": planes,
                                "bit_count": bit_count,
                                "bytes_in_res": bytes_in_res,
                                "icon_id": icon_id,
                                "quality_score": quality_score,
                                "group": icon_group,
                                "index": i,
                            }
                            # print(f"    图标 {i+1}: {actual_width}x{actual_height} {bit_count}位 大小:{bytes_in_res}B (质量分数: {quality_score})")
                            group_icons.append(icon_info)
                        except struct.error:
                            continue  # 解包失败，跳过
                    # 按质量分数排序，优先选择高质量图标
                    group_icons.sort(key=lambda x: x["quality_score"], reverse=True)
                    # 添加所有图标到列表，稍后进行全局去重
                    icon_info_list.extend(group_icons)
            # 全局去重：保留每个尺寸的最佳图标
            seen_sizes = {}
            unique_icons = []
            # 按质量分数排序所有图标
            icon_info_list.sort(key=lambda x: x["quality_score"], reverse=True)
            for icon_info in icon_info_list:
                size_key = (icon_info["width"], icon_info["height"])
                # 如果这个尺寸还没有见过，或者当前图标质量更好，则保留
                if size_key not in seen_sizes:
                    seen_sizes[size_key] = icon_info["quality_score"]
                    unique_icons.append(icon_info)
                elif icon_info["quality_score"] > seen_sizes[size_key]:
                    # 找到更好的图标，替换之前的
                    seen_sizes[size_key] = icon_info["quality_score"]
                    # 移除之前的低质量图标
                    unique_icons = [
                        ico
                        for ico in unique_icons
                        if not (
                            ico["width"] == icon_info["width"]
                            and ico["height"] == icon_info["height"]
                        )
                    ]
                    unique_icons.append(icon_info)
            # 重新按质量排序
            unique_icons.sort(key=lambda x: x["quality_score"], reverse=True)
            icon_info_list = unique_icons

            # 如果图标组中的图标太少，尝试处理单独的图标资源
            if len(icon_info_list) <= 3 and individual_icon_data:
                # print("处理单独的图标资源...")
                for icon_id, icon_data in individual_icon_data:
                    try:
                        # 尝试解析图标数据头部来获取尺寸信息
                        if len(icon_data) >= 16:
                            # PNG格式图标
                            if icon_data[:8] == b"\x89PNG\r\n\x1a\n":
                                # print(f"   图标 {icon_id}: PNG格式")
                                # 直接从内存中解析PNG
                                try:
                                    from io import BytesIO
                                    img = Image.open(BytesIO(icon_data))
                                    width, height = img.size
                                    quality_score = (
                                        width * height
                                    ) + 320  # PNG通常是32位
                                    icon_info_list.append({
                                        "width": width,
                                        "height": height,
                                        "colors": 0,
                                        "planes": 1,
                                        "bit_count": 32,
                                        "bytes_in_res": len(icon_data),
                                        "icon_id": icon_id,
                                        "quality_score": quality_score,
                                        "group": f"INDIVIDUAL_{icon_id}",
                                        "index": 0,
                                        "data": icon_data,
                                        "format": "PNG",
                                    })
                                    # print(f"    {width}x{height} PNG图标")
                                except Exception as e:
                                    # print(f"    PNG解析失败: {e}")
                                    pass
                            # 标准ICO格式
                            else:
                                # 尝试作为ICO数据处理
                                try:
                                    from io import BytesIO

                                    # 解析图标实际尺寸：ICO 中 BITMAPINFOHEADER 的
                                    # biWidth 为宽，biHeight 为 高*2（含 XOR+AND 掩码）
                                    img_w = img_h = 0
                                    if len(icon_data) >= 16:
                                        bi_w, bi_h = struct.unpack(
                                            "<ii", icon_data[4:12]
                                        )
                                        img_w = bi_w
                                        img_h = bi_h // 2
                                    # ICO 目录项的宽高单字节：0 表示 256，其余用真实值
                                    width_byte = img_w if 0 < img_w < 256 else 0
                                    height_byte = img_h if 0 < img_h < 256 else 0
                                    # 创建完整的ICO文件数据
                                    ico_data = BytesIO()
                                    ico_data.write(
                                        struct.pack("<HHH", 0, 1, 1)
                                    )  # ICO文件头
                                    ico_data.write(
                                        struct.pack(
                                            "<BBBBHHLL",
                                            width_byte,
                                            height_byte,
                                            0,
                                            0,
                                            1,
                                            32,
                                            len(icon_data),
                                            22,
                                        )
                                    )  # 图标目录项
                                    ico_data.write(icon_data)  # 图标数据
                                    ico_data.seek(0)

                                    img = Image.open(ico_data)
                                    width, height = img.size
                                    quality_score = (width * height) + 320
                                    icon_info_list.append({
                                        "width": width,
                                        "height": height,
                                        "colors": 0,
                                        "planes": 1,
                                        "bit_count": 32,
                                        "bytes_in_res": len(icon_data),
                                        "icon_id": icon_id,
                                        "quality_score": quality_score,
                                        "group": f"INDIVIDUAL_{icon_id}",
                                        "index": 0,
                                        "data": icon_data,
                                        "format": "ICO",
                                    })
                                    # print(f"    {width}x{height} ICO图标")
                                except Exception as e:
                                    # print(f"    ICO解析失败: {e}")
                                    pass
                    except Exception as e:
                        # print(f"  处理图标 {icon_id} 失败: {e}")
                        continue

                # 重新排序包含单独图标的列表
                icon_info_list.sort(key=lambda x: x["quality_score"], reverse=True)

            # 提取排序后的图标
            for icon_info in icon_info_list:
                try:
                    # 加载图标资源
                    icon_data = win32api.LoadResource(
                        hModule, win32con.RT_ICON, icon_info["icon_id"]
                    )
                    # 创建临时ICO文件到.temp目录
                    temp_ico_name = (
                        f"temp_icon_{icon_info['group']}_{icon_info['index']}.ico"
                    )
                    temp_ico = os.path.join(self.temp_dir, temp_ico_name)
                    with open(temp_ico, "wb") as f:
                        # 写入ICO文件头
                        f.write(
                            struct.pack("<HHH", 0, 1, 1)
                        )  # 保留, 图标类型, 图标数量
                        # 写入图标目录项
                        f.write(
                            struct.pack(
                                "<BBBBHHLL",
                                icon_info["width"] if icon_info["width"] < 256 else 0,
                                icon_info["height"] if icon_info["height"] < 256 else 0,
                                icon_info["colors"],
                                0,
                                icon_info["planes"],
                                icon_info["bit_count"],
                                icon_info["bytes_in_res"],
                                22,
                            )
                        )  # 22是图标数据的偏移量
                        # 写入图标数据
                        f.write(icon_data)
                    # 打开临时ICO文件
                    img = Image.open(temp_ico)

                    # 根据质量过滤选项决定是否添加图标
                    should_add = True
                    if (
                        hasattr(self, "quality_filter")
                        and self.quality_filter
                        and self.quality_filter.get()
                    ):
                        # 启用质量过滤：只添加质量较好的图标（至少24x24，色深至少4位）
                        if not (
                            icon_info["width"] >= 24
                            and icon_info["height"] >= 24
                            and icon_info["bit_count"] >= 4
                        ):
                            should_add = False

                    if should_add:
                        icons.append(img)
                    elif len(icons) == 0:  # 如果没有找到任何图标，至少保留一个
                        icons.append(img)

                    # 删除临时文件
                    os.remove(temp_ico)
                except Exception as e:
                    continue  # 处理单个图标失败，继续处理下一个

            # 释放模块
            win32api.FreeLibrary(hModule)

            # 如果没有找到足够的图标，尝试使用ExtractIcon API进行验证
            if len(icons) <= 1:
                try:
                    import win32gui

                    icon_count = win32gui.ExtractIcon(0, file_path, -1)
                    # 只在调试模式下显示信息
                    # print(f"ExtractIcon 检测到 {icon_count} 个图标，实际提取 {len(icons)} 个")
                except Exception as e:
                    pass

            # 根据质量过滤选项决定返回的图标数量
            if (
                hasattr(self, "quality_filter")
                and self.quality_filter
                and self.quality_filter.get()
            ):
                return icons[:12]  # 质量过滤时限制数量
            else:
                return icons[:20]  # 不过滤时显示更多图标

        except Exception as e:
            raise Exception(f"{str(e)}")

    def select_icon(self, index):
        # 更新选中的图标索引
        self.selected_icon_index.set(index)
        # 更新按钮样式
        for i, btn in enumerate(self.icon_buttons):
            if i == index:
                btn.configure(
                    fg_color=("#3B82F6", "#1E40AF"),
                    border_width=2,
                    border_color=("#60A5FA", "#3B82F6"),
                )  # 选中状态的颜色
            else:
                btn.configure(fg_color=("gray90", "gray13"), border_width=0)  # 默认颜色
        # 更新预览
        self.update_preview()

    def apply_custom_size(self):
        try:
            # 获取自定义尺寸输入
            width_str = self.custom_size_var.get()
            height_str = self.custom_size_var2.get()
            
            # 检查输入是否为空
            if not width_str or not height_str:
                messagebox.showerror("错误", "请输入宽度和高度")
                return
            
            # 转换为整数
            width = int(width_str)
            height = int(height_str)
            
            if width < 16 or width > 1024 or height < 16 or height > 1024:
                messagebox.showerror("错误", "尺寸必须在16-1024之间")
                return
            # 取消所有预设尺寸的选择
            for btn in self.size_buttons:
                btn.deselect()
            # 更新选中的宽高
            self.selected_width.set(width)
            self.selected_height.set(height)
            # 兼容旧逻辑
            self.selected_size.set(str(width))
            # 更新预览
            self.update_preview()
            # 更新信息标签
            if self.icon_list and 0 <= self.selected_icon_index.get() < len(self.icon_list):
                selected_img = self.icon_list[self.selected_icon_index.get()]
                bit_info = BIT_DEPTH_INFO.get(selected_img.mode, selected_img.mode)
                info_text = f"原始: {selected_img.width}x{selected_img.height} | {bit_info}\n输出: {width}x{height}"
                self.info_label.configure(text=info_text)
        except ValueError:
            messagebox.showerror("错误", "请输入有效的数字")

    def on_preset_size_selected(self, size):
        """选择预设尺寸时，宽高都设为相同值"""
        try:
            self.custom_size_var.set("")
            self.custom_size_var2.set("")
            self.selected_width.set(int(size))
            self.selected_height.set(int(size))
            self.selected_size.set(str(size))
            self.update_preview()
        except Exception:
            pass
    
    def set_default_size(self):
        """设置默认尺寸为选中图标的原始尺寸"""
        # 清空自定义尺寸输入
        self.custom_size_var.set("")
        self.custom_size_var2.set("")
        
        # 取消所有预设尺寸的选择
        for btn in self.size_buttons:
            btn.deselect()
        
        # 获取选中图标的原始尺寸并设置为默认尺寸
        if self.icon_list and 0 <= self.selected_icon_index.get() < len(self.icon_list):
            selected_img = self.icon_list[self.selected_icon_index.get()]
            original_size = max(selected_img.width, selected_img.height)  # 取较大的尺寸作为参考
            
            # 直接使用原始宽高作为默认尺寸
            self.selected_width.set(selected_img.width)
            self.selected_height.set(selected_img.height)
            self.selected_size.set(str(original_size))
            
            # 更新信息标签
            bit_info = BIT_DEPTH_INFO.get(selected_img.mode, selected_img.mode)
            info_text = f"原始: {selected_img.width}x{selected_img.height} | {bit_info}\n输出: {self.selected_width.get()}x{self.selected_height.get()}"
            self.info_label.configure(text=info_text)
        else:
            # 如果没有图标，则默认选中64x64
            self.selected_size.set("64")
            self.selected_width.set(64)
            self.selected_height.set(64)
            for btn in self.size_buttons:
                if btn.cget("value") == "64":
                    btn.invoke()
                    break
        
        # 更新预览
        self.update_preview()

    def update_preview(self):
        # 清除当前预览
        for widget in self.preview_frame.winfo_children(): 
            widget.destroy()
        # 获取选中的图标和尺寸
        if self.icon_list and 0 <= self.selected_icon_index.get() < len(self.icon_list):
            selected_img = self.icon_list[self.selected_icon_index.get()]
            target_w = max(1, int(self.selected_width.get() or 1))
            target_h = max(1, int(self.selected_height.get() or 1))
            # 预览最大150x150内等比缩放
            scale = min(150 / target_w, 150 / target_h)
            disp_w = max(1, int(target_w * scale))
            disp_h = max(1, int(target_h * scale))
            img = selected_img.copy().resize((disp_w, disp_h), Image.LANCZOS)
            # 创建CTkImage
            ctk_image = ctk.CTkImage(
                light_image=img, dark_image=img, size=(disp_w, disp_h)
            )
            # 创建标签并居中显示
            preview_label = ctk.CTkLabel(self.preview_frame, image=ctk_image, text="")
            preview_label.ctk_image = ctk_image  # 保持引用
            # 使用place方法居中显示
            preview_label.place(relx=0.5, rely=0.5, anchor="center")

            # 更新图标信息
            width, height = selected_img.size
            bit_info = BIT_DEPTH_INFO.get(selected_img.mode, selected_img.mode)
            info_text = f"原始: {width}x{height} | {bit_info}\n输出: {target_w}x{target_h}"
            self.info_label.configure(text=info_text)
            
            # 更新输出文件扩展名
            self.update_output_extension()
        else:
            self.info_label.configure(text="")
    
    def on_format_change(self):
        """当输出格式改变时更新预览和输出文件扩展名"""
        self.update_preview()

    def _update_quality_label(self, value):
        """压缩滑块拖动时更新百分比显示"""
        self.quality_label.configure(text=f"{int(float(value))}%")

    def update_output_extension(self):
        """根据选择的格式更新输出文件扩展名"""
        # 获取当前输出路径
        current_path = self.output_path.get()
        if current_path:
            # 获取选择的格式
            fmt = self.selected_format.get().lower()
            # 更新扩展名
            name, ext = os.path.splitext(current_path)
            new_path = name + "." + fmt
            self.output_path.set(new_path)

    def save_icon(self):
        output_file = self.output_path.get()
        if not output_file:
            messagebox.showerror("错误", "请选择输出文件")
            return
        if not self.icon_list or not self.icon_buttons:
            messagebox.showerror("错误", "没有可保存的图标")
            return
        # 根据选择的格式更新文件扩展名
        fmt = self.selected_format.get().lower()
        if not output_file.lower().endswith(f".{fmt}"):
            output_file = os.path.splitext(output_file)[0] + f".{fmt}"
            self.output_path.set(output_file)
        # 在新线程中执行保存，避免界面卡顿
        threading.Thread(target=self._save_thread, daemon=True).start()

    def _save_thread(self):
        try:
            self.status.set("保存中...")
            # 显示进度条
            self.after(
                0, lambda: self.progress_bar.pack(side="left", padx=10, pady=10)
            )
            self.after(0, lambda: self.progress_bar.set(0.2))
            output_file = self.output_path.get()
            target_w = int(self.selected_width.get())
            target_h = int(self.selected_height.get())
            icon_index = self.selected_icon_index.get()
            fmt = self.selected_format.get().upper()
            # 确保尺寸在16-1024之间
            target_w = min(max(16, target_w), 1024)
            target_h = min(max(16, target_h), 1024)
            # 获取选中的图标
            img = self.icon_list[icon_index]
            self.after(0, lambda: self.progress_bar.set(0.5))
            # 调整图像大小（非方形也允许；仅ICO会强制方形）
            if fmt == "ICO" and target_w != target_h:
                square = max(target_w, target_h)
                messagebox.showwarning("提示", f"ICO格式需方形，已将输出调整为 {square}x{square}")
                target_w = target_h = square
            img = img.resize((target_w, target_h), Image.LANCZOS)
            self.after(0, lambda: self.progress_bar.set(0.8))
            # 根据选择的格式保存文件
            if fmt == "ICO":
                # 获取原始图标（不缩放），构建多尺寸ICO
                original_img = self.icon_list[icon_index]
                if original_img.mode != 'RGBA':
                    original_img = original_img.convert('RGBA')
                
                # ICO始终包含从16到目标尺寸的所有标准尺寸
                base_size = target_w
                standard_sizes = [256, 128, 96, 64, 48, 32, 24, 16]
                ico_sizes = [s for s in standard_sizes if s <= max(base_size, 256)]
                if not ico_sizes:
                    ico_sizes = [base_size]
                ico_sizes.sort(reverse=True)
                
                # 为每个尺寸生成缩放版本
                ico_frames = []
                for w, h in [(s, s) for s in ico_sizes]:
                    frame = original_img.resize((w, h), Image.LANCZOS)
                    ico_frames.append(frame)
                
                # 保存ICO（第一帧为主帧，其余通过append_images）
                ico_frames[0].save(
                    output_file,
                    format="ICO",
                    append_images=ico_frames[1:],
                    sizes=[(f.width, f.height) for f in ico_frames],
                )
                    
            elif fmt == "PNG":
                img.save(output_file, format="PNG", optimize=True)
            elif fmt == "WEBP":
                quality = int(self.quality_var.get())
                img.save(output_file, format="WEBP", quality=quality)
            self.after(0, lambda: self.progress_bar.set(1.0))
            self.status.set(f"保存完成: {output_file}")
            messagebox.showinfo("成功", f"图标已成功保存到:\n{output_file}")
        except Exception as e:
            self.status.set(f"保存失败: {str(e)}")
            messagebox.showerror("错误", f"保存过程中出现错误:\n{str(e)}")
        finally:
            # 隐藏进度条
            self.after(1000, lambda: self.progress_bar.pack_forget())

    def _render_selected_image(self):
        """按当前选中的图标与尺寸渲染出待保存/复制的图像"""
        icon_index = self.selected_icon_index.get()
        target_w = int(self.selected_width.get())
        target_h = int(self.selected_height.get())
        target_w = min(max(16, target_w), 1024)
        target_h = min(max(16, target_h), 1024)
        fmt = self.selected_format.get().upper()
        if fmt == "ICO" and target_w != target_h:
            square = max(target_w, target_h)
            target_w = target_h = square
        img = self.icon_list[icon_index]
        img = img.resize((target_w, target_h), Image.LANCZOS).convert("RGBA")
        return img, fmt

    def copy_icon(self):
        """将当前选中的图标复制到剪贴板"""
        if not self.icon_list or not self.icon_buttons:
            messagebox.showerror("错误", "没有可复制的图标")
            return
        try:
            img, fmt = self._render_selected_image()
            if HAS_WIN32:
                import win32clipboard
                from io import BytesIO
                # 转换为DIB（去掉BMP文件头）后写入剪贴板，可直接粘贴到画图等程序
                buf = BytesIO()
                img.convert("RGB").save(buf, "BMP")
                dib = buf.getvalue()[14:]
                win32clipboard.OpenClipboard()
                try:
                    win32clipboard.EmptyClipboard()
                    win32clipboard.SetClipboardData(win32clipboard.CF_DIB, dib)
                finally:
                    win32clipboard.CloseClipboard()
                self.status.set("已复制图像到剪贴板")
                messagebox.showinfo("成功", "图标图像已复制到剪贴板，可粘贴到画图等程序")
            else:
                # 非Windows环境回退：复制PNG的base64数据
                import base64
                from io import BytesIO
                buf = BytesIO()
                img.save(buf, "PNG")
                b64 = base64.b64encode(buf.getvalue()).decode()
                self.clipboard_clear()
                self.clipboard_append(b64)
                self.status.set("已复制图像数据(base64)到剪贴板")
                messagebox.showinfo("成功", "已将图标PNG的base64数据复制到剪贴板")
        except Exception as e:
            messagebox.showerror("错误", f"复制失败:\n{str(e)}")

    # 去除ICO手写备选实现所需的BMP工具函数

if __name__ == "__main__":
    # 创建应用实例
    app = ImageToIcoConverter()
    app.mainloop()
