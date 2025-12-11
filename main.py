import os
import sys
import hashlib
import json
import shutil
import webbrowser
from datetime import datetime
from urllib.parse import quote

import tkinter as tk
from tkinter import filedialog, messagebox, ttk, TclError

try:
    from PIL import Image, ExifTags
except Exception:
    Image = None
    ExifTags = None

try:
    import requests
except Exception:
    requests = None

APP_TITLE = "图片时间与证据生成器（测试版）"
APP_DIR = os.path.abspath(os.path.dirname(__file__))
OUTPUT_DIR = os.path.join(APP_DIR, "evidence_packages")

if not os.path.exists(OUTPUT_DIR):
    os.makedirs(OUTPUT_DIR, exist_ok=True)


def safe_filename(name: str) -> str:
    return "".join(c for c in name if c.isalnum() or c in ("-", "_", ".")).strip()


def sha256_file(path: str) -> str:
    h = hashlib.sha256()
    with open(path, 'rb') as f:
        for chunk in iter(lambda: f.read(8192), b''):
            h.update(chunk)
    return h.hexdigest()


def read_exif(path: str) -> dict:
    if Image is None:
        return {"error": "Pillow 未安装，无法读取EXIF"}
    try:
        img = Image.open(path)
        exif_data = img._getexif() if hasattr(img, "_getexif") else None
        result = {}
        if exif_data:
            for k, v in exif_data.items():
                tag = ExifTags.TAGS.get(k, str(k)) if ExifTags else str(k)
                result[tag] = str(v)
        return result
    except Exception as e:
        return {"error": f"读取EXIF失败: {e}"}


def fetch_world_time() -> dict:
    """
    获取 UTC 时间，具备多源回退与重试：
    1) worldtimeapi.org JSON
    2) timeapi.io JSON
    3) 多站点 HTTP Date 响应头（Google / Microsoft / Baidu）
    返回字段：{"source", "utc_datetime", "unixtime"}
    """
    if requests is None:
        return {"error": "requests 未安装，无法获取网络时间"}

    # 构建带重试的会话
    try:
        from requests.adapters import HTTPAdapter
        from urllib3.util import Retry
        session = requests.Session()
        retries = Retry(total=3, backoff_factor=0.3, status_forcelist=[429, 500, 502, 503, 504])
        session.mount("https://", HTTPAdapter(max_retries=retries))
        session.mount("http://", HTTPAdapter(max_retries=retries))
    except Exception:
        session = requests  # 回退为直接使用 requests

    sources = [
        ("https://worldtimeapi.org/api/timezone/Etc/UTC", "json_wta", "json_wta"),
        ("https://timeapi.io/api/Time/current/zone?timeZone=UTC", "json_timeapi", "json_timeapi"),
        ("https://www.google.com", "http_date_google", "http_date"),
        ("https://www.microsoft.com", "http_date_ms", "http_date"),
        ("https://www.baidu.com", "http_date_baidu", "http_date"),
    ]

    errors = []

    for url, tag, kind in sources:
        try:
            if kind == "json_wta":
                r = session.get(url, timeout=8)
                r.raise_for_status()
                data = r.json()
                utc_dt = data.get("utc_datetime")
                unixtime = data.get("unixtime")
                if utc_dt:
                    return {"source": url, "utc_datetime": utc_dt, "unixtime": unixtime}

            elif kind == "json_timeapi":
                r = session.get(url, timeout=8)
                r.raise_for_status()
                data = r.json()
                # timeapi.io 返回的字段可能为 dateTime（ISO8601）
                iso = data.get("dateTime") or data.get("time")
                if iso:
                    try:
                        from datetime import datetime, timezone
                        # 兼容末尾 Z 的情况
                        d = datetime.fromisoformat(iso.replace("Z", "+00:00"))
                        ts = int(d.timestamp())
                    except Exception:
                        ts = None
                    return {"source": url, "utc_datetime": iso, "unixtime": ts}

            elif kind == "http_date":
                # 优先 HEAD，若失败或无 Date 则回退 GET（部分站点对 HEAD 支持不佳）
                r = session.head(url, timeout=6, allow_redirects=True)
                if r.status_code >= 400 or ("Date" not in r.headers and "date" not in r.headers):
                    r = session.get(url, timeout=8, stream=True)
                date_hdr = r.headers.get("Date") or r.headers.get("date")
                if date_hdr:
                    try:
                        from email.utils import parsedate_to_datetime
                        from datetime import timezone
                        d = parsedate_to_datetime(date_hdr)
                        if d.tzinfo is None:
                            d = d.replace(tzinfo=timezone.utc)
                        iso = d.astimezone(timezone.utc).isoformat()
                        ts = int(d.timestamp())
                        return {"source": f"{url} (HTTP Date)", "utc_datetime": iso, "unixtime": ts}
                    except Exception:
                        # 无法解析则直接返回原始 Date 文本
                        return {"source": f"{url} (HTTP Date)", "utc_datetime": date_hdr}
        except Exception as e:
            errors.append(f"{tag}: {e}")

    return {"error": "获取网络时间失败", "details": errors}


def fetch_stock_quote(symbol: str) -> dict:
    """
    使用 stooq 免费接口，无需API Key。
    例如: AAPL.US, TSLA.US, 000001.SS
    文档: https://stooq.com/db/
    """
    if requests is None:
        return {"error": "requests 未安装，无法获取股价"}
    sym = symbol.strip().lower()
    if not sym:
        return {"error": "股票代码不能为空"}
    # CSV 简易接口
    url = f"https://stooq.com/q/l/?s={quote(sym)}&f=sd2t2ohlcv&h&e=csv"
    try:
        r = requests.get(url, timeout=8)
        r.raise_for_status()
        lines = r.text.strip().splitlines()
        if len(lines) < 2:
            return {"error": "返回数据异常", "source": url}
        header = [h.strip().lower() for h in lines[0].split(',')]
        values = [v.strip() for v in lines[1].split(',')]
        data = dict(zip(header, values))
        data["source"] = url
        return data
    except Exception as e:
        return {"error": f"获取股价失败: {e}", "source": url}


class EvidenceApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title(APP_TITLE)
        self.root.geometry("980x700")
        self.photo_path = None
        self.state = {
            "hash": None,
            "exif": None,
            "world_time": None,
            "stocks": [],  # 改为列表存储多个股票数据
            "publish_urls": [],
            "created_at": None,
        }

        self._build_ui()

    def _build_ui(self):
        # 顶部工具条（固定）
        top = ttk.Frame(self.root)
        top.pack(fill=tk.X, padx=12, pady=8)

        ttk.Button(top, text="选择照片", command=self.choose_photo).pack(side=tk.LEFT)
        self.photo_label = ttk.Label(top, text="未选择")
        self.photo_label.pack(side=tk.LEFT, padx=10)

        ttk.Button(top, text="计算SHA-256", command=self.do_hash).pack(side=tk.LEFT, padx=10)
        ttk.Button(top, text="读取EXIF", command=self.do_exif).pack(side=tk.LEFT)

        # 中部滚动容器
        container = ttk.Frame(self.root)
        container.pack(fill=tk.BOTH, expand=True)
        self.canvas = tk.Canvas(container, highlightthickness=0)
        vbar = ttk.Scrollbar(container, orient=tk.VERTICAL, command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=vbar.set)
        vbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # 内容区放入 Canvas
        self.content = ttk.Frame(self.canvas)
        self.content_id = self.canvas.create_window((0, 0), window=self.content, anchor="nw")

        def _on_frame_configure(event):
            # 根据内容高度更新滚动区域
            self.canvas.configure(scrollregion=self.canvas.bbox("all"))
            # 自适应内容宽度
            try:
                self.canvas.itemconfigure(self.content_id, width=self.canvas.winfo_width())
            except Exception:
                pass
        self.content.bind("<Configure>", _on_frame_configure)

        # 绑定鼠标滚轮
        self._bind_mousewheel(self.canvas)

        # 预览区域（可滚动）
        preview = ttk.LabelFrame(self.content, text="图片预览")
        preview.pack(fill=tk.BOTH, expand=False, padx=12, pady=4)
        self.preview_label = ttk.Label(preview, text="未选择图片", anchor=tk.CENTER)
        self.preview_label.pack(fill=tk.BOTH, expand=True)
        self._preview_image = None  # 保存缩略图引用防止被GC

        # 股票/时间区块（可滚动）
        stock_frame = ttk.LabelFrame(self.content, text="股票行情数据")
        stock_frame.pack(fill=tk.X, padx=12, pady=8)
        
        # 股票数量设置
        stock_count_frame = ttk.Frame(stock_frame)
        stock_count_frame.pack(fill=tk.X, padx=8, pady=4)
        ttk.Label(stock_count_frame, text="股票数量:").pack(side=tk.LEFT)
        self.stock_count_var = tk.StringVar(value="3")
        stock_count_spinbox = ttk.Spinbox(stock_count_frame, from_=1, to=10, width=5, 
                                         textvariable=self.stock_count_var, 
                                         command=self.update_stock_inputs)
        stock_count_spinbox.pack(side=tk.LEFT, padx=4)
        ttk.Button(stock_count_frame, text="更新输入框", command=self.update_stock_inputs).pack(side=tk.LEFT, padx=4)
        self.btn_time = ttk.Button(stock_count_frame, text="获取网络时间(UTC)", command=self.do_time)
        self.btn_time.pack(side=tk.RIGHT)
        
        # 股票代码输入区域
        self.stock_inputs_frame = ttk.Frame(stock_frame)
        self.stock_inputs_frame.pack(fill=tk.X, padx=8, pady=4)
        self.stock_entries = []  # 存储股票输入框
        
        # 股票操作按钮
        stock_btn_frame = ttk.Frame(stock_frame)
        stock_btn_frame.pack(fill=tk.X, padx=8, pady=4)
        self.btn_stock = ttk.Button(stock_btn_frame, text="获取所有股价", command=self.do_stocks)
        self.btn_stock.pack(side=tk.LEFT)
        ttk.Button(stock_btn_frame, text="清空股票数据", command=self.clear_stocks).pack(side=tk.LEFT, padx=4)
        
        # 初始化股票输入框
        self.update_stock_inputs()

        # 发布URL（可滚动）
        pub = ttk.LabelFrame(self.content, text="发布URL（手动在平台发布后将网址粘贴于此，一行一个）")
        pub.pack(fill=tk.BOTH, expand=False, padx=12, pady=8)
        self.publish_text = tk.Text(pub, height=6)
        self.publish_text.pack(fill=tk.BOTH, expand=True)

        # 操作按钮（可滚动）
        act = ttk.Frame(self.content)
        act.pack(fill=tk.X, padx=12, pady=8)
        self.btn_report = ttk.Button(act, text="生成证明报告(HTML)", command=self.generate_report)
        self.btn_report.pack(side=tk.LEFT)
        ttk.Button(act, text="打开输出目录", command=self.open_output_dir).pack(side=tk.LEFT, padx=8)

        # Markdown 预览区（可滚动）
        md = ttk.LabelFrame(self.content, text="Markdown预览（只读）")
        md.pack(fill=tk.BOTH, expand=False, padx=12, pady=8)
        toolbar = ttk.Frame(md)
        toolbar.pack(fill=tk.X, pady=(6,4))
        ttk.Button(toolbar, text="打开 .md 文件", command=self.choose_md).pack(side=tk.LEFT)
        self.md_text = tk.Text(md, height=12, wrap=tk.WORD)
        self.md_text.pack(fill=tk.BOTH, expand=True)
        # Markdown 标签样式
        self.md_text.tag_configure("h1", font=(None, 16, "bold"))
        self.md_text.tag_configure("h2", font=(None, 14, "bold"))
        self.md_text.tag_configure("h3", font=(None, 12, "bold"))
        self.md_text.tag_configure("bold", font=(None, 11, "bold"))
        self.md_text.tag_configure("italic", font=(None, 11, "italic"))
        self.md_text.tag_configure("code", font=("Consolas", 11), background="#f5f5f5")
        self.md_text.tag_configure("link", foreground="#1a73e8")
        self.md_text.config(state=tk.DISABLED)

        # 日志（可滚动区域的一部分）
        self.log = tk.Text(self.content, height=16)
        self.log.pack(fill=tk.BOTH, expand=True, padx=12, pady=8)
        self.log.config(state=tk.DISABLED)

        # 底部状态栏（固定）
        status = ttk.Frame(self.root)
        status.pack(fill=tk.X, padx=12, pady=(0, 8))
        self.status_var = tk.StringVar(value="就绪")
        ttk.Label(status, textvariable=self.status_var).pack(side=tk.LEFT)
        self.progress = ttk.Progressbar(status, mode="indeterminate", length=160)
        self.progress.pack(side=tk.RIGHT)

    def _bind_mousewheel(self, widget):
        # Windows 使用 <MouseWheel>
        widget.bind_all("<MouseWheel>", self._on_mousewheel)

    def _on_mousewheel(self, event):
        try:
            # event.delta 在 Windows 上通常为 120 的倍数
            delta = int(-1 * (event.delta / 120))
            self.canvas.yview_scroll(delta, "units")
        except Exception:
            pass

    def _set_status(self, msg: str):
        try:
            self.status_var.set(msg)
            self.root.update_idletasks()
        except Exception:
            pass

    def _busy(self, on: bool):
        try:
            if on:
                if hasattr(self, 'progress'):
                    self.progress.start(80)
                for btn in (getattr(self, 'btn_time', None), getattr(self, 'btn_stock', None), getattr(self, 'btn_report', None)):
                    if btn is not None:
                        btn.config(state=tk.DISABLED)
            else:
                if hasattr(self, 'progress'):
                    self.progress.stop()
                for btn in (getattr(self, 'btn_time', None), getattr(self, 'btn_stock', None), getattr(self, 'btn_report', None)):
                    if btn is not None:
                        btn.config(state=tk.NORMAL)
        except Exception:
            pass

    def log_append(self, text: str):
        self.log.config(state=tk.NORMAL)
        self.log.insert(tk.END, f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} - {text}\n")
        self.log.see(tk.END)
        self.log.config(state=tk.DISABLED)

    def choose_photo(self):
        path = filedialog.askopenfilename(title="选择照片", filetypes=[
            ("图片文件", ".jpg .jpeg .png .bmp .tif .tiff"),
            ("所有文件", ".*")
        ])
        if not path:
            return
        self.photo_path = path
        self.photo_label.config(text=os.path.basename(path))
        self.log_append(f"选择照片: {path}")
        # 加载缩略图预览
        try:
            if Image is not None:
                img = Image.open(path)
                img.thumbnail((640, 360))
                from PIL import ImageTk
                self._preview_image = ImageTk.PhotoImage(img)
                self.preview_label.configure(image=self._preview_image, text="")
            else:
                self.preview_label.configure(text="未安装Pillow，无法预览")
        except Exception as e:
            self.preview_label.configure(text=f"预览失败: {e}")

    def do_hash(self):
        if not self.photo_path:
            messagebox.showwarning("提示", "请先选择照片")
            return
        h = sha256_file(self.photo_path)
        self.state["hash"] = h
        self.log_append(f"SHA-256: {h}")

    def do_exif(self):
        if not self.photo_path:
            messagebox.showwarning("提示", "请先选择照片")
            return
        exif = read_exif(self.photo_path)
        self.state["exif"] = exif
        if "error" in exif:
            self.log_append(exif["error"])
        else:
            dt = exif.get("DateTimeOriginal") or exif.get("DateTime") or exif.get("CreateDate")
            self.log_append(f"读取EXIF成功，拍摄时间: {dt if dt else '未知'}")

    def do_time(self):
        self._busy(True)
        self._set_status("正在获取网络UTC时间…")
        try:
            data = fetch_world_time()
            self.state["world_time"] = data
            if "error" in data:
                self.log_append(data["error"])
            else:
                self.log_append(f"网络UTC时间: {data.get('utc_datetime')} (来源: {data.get('source')})")
        finally:
            self._set_status("就绪")
            self._busy(False)

    def update_stock_inputs(self):
        """更新股票输入框数量"""
        try:
            count = int(self.stock_count_var.get())
        except ValueError:
            count = 3
        
        # 清除现有输入框
        for widget in self.stock_inputs_frame.winfo_children():
            widget.destroy()
        self.stock_entries.clear()
        
        # 创建新的输入框
        default_symbols = ["AAPL.US", "TSLA.US", "000001.SS", "MSFT.US", "GOOGL.US", 
                          "AMZN.US", "META.US", "NVDA.US", "BABA.US", "JD.US"]
        
        for i in range(count):
            row = ttk.Frame(self.stock_inputs_frame)
            row.pack(fill=tk.X, pady=2)
            
            ttk.Label(row, text=f"股票{i+1}:").pack(side=tk.LEFT)
            entry = ttk.Entry(row)
            entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=4)
            
            # 设置默认值
            default_symbol = default_symbols[i] if i < len(default_symbols) else f"STOCK{i+1}.US"
            entry.insert(0, default_symbol)
            
            self.stock_entries.append(entry)
    
    def do_stocks(self):
        """获取多个股票的行情数据"""
        if not self.stock_entries:
            messagebox.showwarning("提示", "请先设置股票数量")
            return
        
        self._busy(True)
        self._set_status("正在获取股票行情数据…")
        
        stocks_data = []
        errors = []
        
        for i, entry in enumerate(self.stock_entries):
            symbol = entry.get().strip()
            if not symbol:
                continue
                
            try:
                self._set_status(f"正在获取股价: {symbol}…")
                data = fetch_stock_quote(symbol)
                
                # 添加股票代码信息
                data["symbol"] = symbol.upper()
                
                if "error" in data:
                    errors.append(f"{symbol}: {data['error']}")
                else:
                    stocks_data.append(data)
                    self.log_append(f"股价: {symbol} {data.get('close')} 日期: {data.get('date')} 时间: {data.get('time')}")
                    
            except Exception as e:
                errors.append(f"{symbol}: 获取失败 - {e}")
        
        # 更新状态
        self.state["stocks"] = stocks_data
        
        if errors:
            self.log_append(f"部分股票获取失败: {'; '.join(errors)}")
        
        if stocks_data:
            self.log_append(f"成功获取 {len(stocks_data)} 个股票的行情数据")
        else:
            self.log_append("未能获取任何股票数据")
            
        self._set_status("就绪")
        self._busy(False)
    
    def clear_stocks(self):
        """清空股票数据"""
        self.state["stocks"] = []
        self.log_append("已清空股票数据")
    
    def do_stock(self):
        """保留原有的单个股票获取方法以兼容性"""
        symbol = self.stock_entries[0].get().strip() if self.stock_entries else ""
        if not symbol:
            messagebox.showwarning("提示", "请先输入股票代码")
            return
        self.do_stocks()

    def choose_md(self):
        path = filedialog.askopenfilename(title="选择Markdown文件", filetypes=[
            ("Markdown", ".md"),
            ("所有文件", ".*")
        ])
        if not path:
            return
        try:
            self.preview_md(path)
            self.log_append(f"预览Markdown: {path}")
        except Exception as e:
            messagebox.showerror("错误", f"加载Markdown失败: {e}")

    def preview_md(self, path: str):
        try:
            with open(path, "r", encoding="utf-8") as f:
                content = f.read()
            # 简易Markdown解析（标题、加粗、斜体、行内代码、代码块、链接）
            self.md_text.config(state=tk.NORMAL)
            self.md_text.delete("1.0", tk.END)
            import re
            lines = content.splitlines()
            in_code_block = False
            for i, line in enumerate(lines):
                # 代码块 ```
                if line.strip().startswith("```"):
                    in_code_block = not in_code_block
                    if in_code_block:
                        self.md_text.insert(tk.END, "\n")
                    else:
                        self.md_text.insert(tk.END, "\n")
                    continue
                if in_code_block:
                    start = self.md_text.index(tk.END)
                    self.md_text.insert(tk.END, line + "\n")
                    end = self.md_text.index(tk.END)
                    self.md_text.tag_add("code", start, end)
                    continue
                # 标题
                if line.startswith("### "):
                    start = self.md_text.index(tk.END)
                    self.md_text.insert(tk.END, line[4:] + "\n")
                    end = self.md_text.index(tk.END)
                    self.md_text.tag_add("h3", start, end)
                    continue
                if line.startswith("## "):
                    start = self.md_text.index(tk.END)
                    self.md_text.insert(tk.END, line[3:] + "\n")
                    end = self.md_text.index(tk.END)
                    self.md_text.tag_add("h2", start, end)
                    continue
                if line.startswith("# "):
                    start = self.md_text.index(tk.END)
                    self.md_text.insert(tk.END, line[2:] + "\n")
                    end = self.md_text.index(tk.END)
                    self.md_text.tag_add("h1", start, end)
                    continue
                # 普通段落，处理加粗、斜体、行内代码、链接
                text = line + "\n"
                start_para = self.md_text.index(tk.END)
                self.md_text.insert(tk.END, text)
                end_para = self.md_text.index(tk.END)
                # 加粗 **text**
                for m in re.finditer(r"\*\*(.+?)\*\*", text):
                    s = f"{start_para.split('.')[0]}.{int(start_para.split('.')[1]) + m.start()}"
                    e = f"{start_para.split('.')[0]}.{int(start_para.split('.')[1]) + m.end()}"
                    self.md_text.tag_add("bold", s, e)
                # 斜体 *text* 或 _text_
                for m in re.finditer(r"(?<!\*)\*(.+?)\*(?!\*)|_(.+?)_", text):
                    s = f"{start_para.split('.')[0]}.{int(start_para.split('.')[1]) + m.start()}"
                    e = f"{start_para.split('.')[0]}.{int(start_para.split('.')[1]) + m.end()}"
                    self.md_text.tag_add("italic", s, e)
                # 行内代码 `code`
                for m in re.finditer(r"`([^`]+)`", text):
                    s = f"{start_para.split('.')[0]}.{int(start_para.split('.')[1]) + m.start()}"
                    e = f"{start_para.split('.')[0]}.{int(start_para.split('.')[1]) + m.end()}"
                    self.md_text.tag_add("code", s, e)
                # 链接 [text](url)
                for m in re.finditer(r"\[([^\]]+)\]\(([^\)]+)\)", text):
                    s = f"{start_para.split('.')[0]}.{int(start_para.split('.')[1]) + m.start(1)}"
                    e = f"{start_para.split('.')[0]}.{int(start_para.split('.')[1]) + m.end(1)}"
                    self.md_text.tag_add("link", s, e)
            self.md_text.config(state=tk.DISABLED)
        except Exception as e:
            self.md_text.config(state=tk.NORMAL)
            self.md_text.delete("1.0", tk.END)
            self.md_text.insert(tk.END, f"预览失败: {e}")
            self.md_text.config(state=tk.DISABLED)

    def open_output_dir(self):
        os.makedirs(OUTPUT_DIR, exist_ok=True)
        webbrowser.open(OUTPUT_DIR)

    def _collect_publish_urls(self):
        text = self.publish_text.get("1.0", tk.END).strip()
        urls = [line.strip() for line in text.splitlines() if line.strip()]
        self.state["publish_urls"] = urls
        return urls

    def _make_package_dir(self) -> str:
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        base = os.path.join(OUTPUT_DIR, f"evidence_{ts}")
        os.makedirs(base, exist_ok=True)
        return base

    def _copy_photo(self, pkg_dir: str):
        if not self.photo_path:
            return None
        name = safe_filename(os.path.basename(self.photo_path))
        dst = os.path.join(pkg_dir, name)
        try:
            shutil.copy2(self.photo_path, dst)
            return dst
        except Exception as e:
            self.log_append(f"复制图片失败: {e}")
            return None

    def generate_report(self):
        if not self.photo_path:
            messagebox.showwarning("提示", "请先选择照片并计算哈希/读取EXIF")
            return
        if not self.state.get("hash"):
            self.do_hash()
        if not self.state.get("world_time"):
            self.do_time()
        self._collect_publish_urls()
        self.state["created_at"] = datetime.now().isoformat()

        pkg = self._make_package_dir()
        copied = self._copy_photo(pkg)

        # 保存元数据
        meta = {
            "photo_original_path": self.photo_path,
            "photo_copied_path": copied,
            "sha256": self.state.get("hash"),
            "exif": self.state.get("exif"),
            "world_time": self.state.get("world_time"),
            "stocks": self.state.get("stocks", []),  # 改为多个股票数据
            "stock": self.state.get("stocks", [{}])[0] if self.state.get("stocks") else None,  # 保留兼容性
            "publish_urls": self.state.get("publish_urls", []),
            "created_at": self.state.get("created_at"),
            "app": APP_TITLE,
            "version": "1.0.0"
        }

        # 第三方验证：本地哈希比对 + 发布平台验证
        verification = {
            "hash_check": None,
            "publish_checks": []
        }
        # 本地哈希复算
        try:
            if copied and os.path.isfile(copied):
                new_hash = sha256_file(copied)
                verification["hash_check"] = {
                    "recomputed_sha256": new_hash,
                    "expected_sha256": meta.get("sha256"),
                    "match": (new_hash == meta.get("sha256"))
                }
            else:
                verification["hash_check"] = {
                    "error": "未找到已复制文件，无法复算哈希"
                }
        except Exception as e:
            verification["hash_check"] = {"error": f"哈希复算失败: {e}"}

        # 发布平台验证：抓取HTTP Date/状态码/重定向后URL
        urls = meta.get("publish_urls") or []
        if urls:
            if requests is None:
                verification["publish_checks"].append({"error": "requests 未安装，无法验证发布平台"})
            else:
                try:
                    from requests.adapters import HTTPAdapter
                    from urllib3.util import Retry
                    sess = requests.Session()
                    retries = Retry(total=2, backoff_factor=0.2, status_forcelist=[429, 500, 502, 503, 504])
                    sess.mount("https://", HTTPAdapter(max_retries=retries))
                    sess.mount("http://", HTTPAdapter(max_retries=retries))
                except Exception:
                    sess = requests
                for u in urls:
                    info = {"url": u}
                    try:
                        r = None
                        # 先尝试 HEAD
                        try:
                            r = sess.head(u, timeout=8, allow_redirects=True)
                            # 某些站点对 HEAD 支持不佳，若无 Date 继续 GET
                            if r.status_code >= 400 or ("Date" not in r.headers and "date" not in r.headers):
                                raise Exception("HEAD 无 Date 或状态异常")
                        except Exception:
                            r = sess.get(u, timeout=10, stream=True, allow_redirects=True)
                        info.update({
                            "status_code": getattr(r, 'status_code', None),
                            "final_url": getattr(r, 'url', u),
                            "http_date": (r.headers.get("Date") or r.headers.get("date")) if r is not None else None,
                            "content_type": r.headers.get("Content-Type") if r is not None else None,
                            "content_length": r.headers.get("Content-Length") if r is not None else None,
                        })
                    except Exception as e:
                        info["error"] = str(e)
                    verification["publish_checks"].append(info)

        # 将验证信息写入元数据
        meta["verification"] = verification

        # 写 evidence.json（包含验证结果）
        meta_path = os.path.join(pkg, "evidence.json")
        with open(meta_path, "w", encoding="utf-8") as f:
            json.dump(meta, f, ensure_ascii=False, indent=2)

        # 生成HTML报告
        report_path = os.path.join(pkg, "report.html")
        html = self._build_html_report(meta)
        with open(report_path, "w", encoding="utf-8") as f:
            f.write(html)

        # 生成更详细的 Markdown 报告
        md_path = os.path.join(pkg, "report.md")
        try:
            lines = []
            lines.append(f"# {APP_TITLE} - 证明报告\n")
            lines.append(f"- 生成时间：{meta.get('created_at')}")
            lines.append(f"- 应用版本：{meta.get('version')}\n")

            lines.append("## 一、照片基础信息")
            lines.append(f"- 原始路径：`{meta.get('photo_original_path')}`")
            lines.append(f"- 复制路径：`{meta.get('photo_copied_path')}`")
            lines.append(f"- SHA-256 摘要：`{meta.get('sha256')}`\n")

            lines.append("## 二、EXIF 元数据")
            exif = meta.get('exif') or {}
            if isinstance(exif, dict) and exif:
                lines.append("")
                lines.append("| 字段 | 值 |")
                lines.append("|---|---|")
                for k, v in exif.items():
                    lines.append(f"| {str(k)} | {str(v).replace('|','\\|')} |")
                lines.append("")
            else:
                lines.append("- 无或读取失败\n")

            lines.append("## 三、外部不可预测信息")
            world = meta.get('world_time') or {}
            if world:
                lines.append(f"- 网络UTC时间：{world.get('utc_datetime')} (来源：{world.get('source')})")
            
            stocks = meta.get('stocks', [])
            if stocks:
                lines.append(f"- 股票行情数据（共{len(stocks)}只）：")
                lines.append("")
                lines.append("| 股票代码 | 收盘价 | 日期 | 时间 | 成交量 | 数据源 |")
                lines.append("|---|---|---|---|---|---|")
                for stock_item in stocks:
                    if isinstance(stock_item, dict):
                        lines.append(f"| {stock_item.get('symbol','')} | {stock_item.get('close','')} | {stock_item.get('date','')} | {stock_item.get('time','')} | {stock_item.get('volume','')} | [数据源]({stock_item.get('source','')}) |")
                lines.append("")
            else:
                # 兼容性：单个股票数据
                stock = meta.get('stock') or {}
                if stock:
                    lines.append(f"- 股票：收盘 {stock.get('close')}，日期 {stock.get('date')}，时间 {stock.get('time')}，来源 {stock.get('source')}")
            
            lines.append("")
            lines.append(
                "> 注：通过引用公开且不可预知的第三方数据（如网络标准时间与股票行情）来佐证生成时间点。\n"
            )

            lines.append("## 四、发布记录与第三方验证")
            urls = meta.get('publish_urls') or []
            if urls:
                for u in urls:
                    lines.append(f"- 发布URL：{u}")
            else:
                lines.append("- 暂无（请在平台发布后粘贴URL）")
            lines.append("")
            v = meta.get('verification') or {}
            hc = v.get('hash_check') or {}
            if hc:
                ok = hc.get('match')
                lines.append("### 本地哈希比对结果")
                if isinstance(ok, bool):
                    lines.append(f"- 复算SHA-256：`{hc.get('recomputed_sha256')}`")
                    lines.append(f"- 期望SHA-256：`{hc.get('expected_sha256')}`")
                    lines.append(f"- 是否一致：{'是' if ok else '否'}\n")
                elif 'error' in hc:
                    lines.append(f"- 错误：{hc.get('error')}\n")
            pcs = v.get('publish_checks') or []
            if pcs:
                lines.append("### 发布平台验证（HTTP 头部）")
                lines.append("")
                lines.append("| URL | 状态码 | 最终URL | HTTP Date | Content-Type | Content-Length | 备注 |")
                lines.append("|---|---:|---|---|---|---:|---|")
                for item in pcs:
                    lines.append(
                        f"| {item.get('url','')} | {item.get('status_code','')} | {item.get('final_url','')} | {item.get('http_date','')} | {item.get('content_type','')} | {item.get('content_length','')} | {item.get('error','')} |"
                    )
                lines.append("")

            lines.append("## 五、验证指引")
            lines.append("1. 第三方可独立下载图片并复算 SHA-256，与报告一致即证明未被篡改。")
            lines.append("2. 打开上述发布URL，核对平台公开时间与 HTTP Date，辅助限定生成时间窗口。")
            lines.append("3. 对网络时间与股票行情来源，可访问其官方接口或站点进行交叉验证。\n")

            with open(md_path, 'w', encoding='utf-8') as f:
                f.write("\n".join(lines))
        except Exception as e:
            self.log_append(f"生成 Markdown 报告失败: {e}")

        self.log_append(f"已生成证据包: {pkg}")
        webbrowser.open(report_path)

    def _build_html_report(self, meta: dict) -> str:
        def esc(s: str) -> str:
            return (s or "").replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")

        photo_rel = os.path.basename(meta.get("photo_copied_path") or "")
        exif = meta.get("exif") or {}
        world = meta.get("world_time") or {}
        stocks = meta.get("stocks", [])  # 获取多个股票数据
        stock = meta.get("stock") or {}  # 兼容性
        publish_urls = meta.get("publish_urls") or []

        exif_rows = "".join(
            f"<tr><td>{esc(str(k))}</td><td>{esc(str(v))}</td></tr>" for k, v in (exif.items() if isinstance(exif, dict) else [])
        ) if exif else "<tr><td colspan=2>无或读取失败</td></tr>"

        publish_list = "".join(
            f"<li><a href='{esc(u)}' target='_blank' rel='noopener noreferrer'>{esc(u)}</a></li>" for u in publish_urls
        ) or "<li>暂无（请在平台发布后粘贴URL）</li>"

        # 生成多个股票数据的HTML表格
        stock_block = ""
        if stocks:
            stock_rows = "".join(
                f"<tr><td>{esc(stock_item.get('symbol',''))}</td>"
                f"<td>{esc(stock_item.get('close',''))}</td>"
                f"<td>{esc(stock_item.get('date',''))}</td>"
                f"<td>{esc(stock_item.get('time',''))}</td>"
                f"<td>{esc(stock_item.get('volume',''))}</td>"
                f"<td><a href='{esc(stock_item.get('source',''))}' target='_blank'>数据源</a></td></tr>"
                for stock_item in stocks if isinstance(stock_item, dict)
            )
            stock_block = f"""
            <table class="table">
              <thead><tr><th>股票代码</th><th>收盘价</th><th>日期</th><th>时间</th><th>成交量</th><th>数据源</th></tr></thead>
              <tbody>{stock_rows}</tbody>
            </table>
            """
        elif stock:  # 兼容性
            stock_block = f"<p><strong>股票数据</strong>：代码: {esc(stock.get('symbol',''))}, 收盘: {esc(stock.get('close',''))}, 日期: {esc(stock.get('date',''))}, 时间: {esc(stock.get('time',''))}，来源: <a href='{esc(stock.get('source',''))}' target='_blank'>数据源</a></p>"

        world_block = "" if not world else (
            f"<p><strong>网络UTC时间</strong>：{esc(world.get('utc_datetime',''))}（来源：<a href='{esc(world.get('source',''))}' target='_blank'>worldtimeapi</a>）</p>"
        )

        return f"""
<!DOCTYPE html>
<html lang=zh>
<head>
<meta charset="utf-8" />
<title>{esc(APP_TITLE)} - 证明报告</title>
<style>
body {{ font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, 'Helvetica Neue', Arial, 'Noto Sans', 'PingFang SC', 'Microsoft YaHei', sans-serif; margin: 24px; }}
code, pre {{ background: #f7f7f7; padding: 2px 4px; }}
.table {{ border-collapse: collapse; width: 100%; }}
.table th, .table td {{ border: 1px solid #ddd; padding: 8px; }}
.section {{ margin: 18px 0; }}
small {{ color: #666; }}
</style>
</head>
<body>
<h1>{esc(APP_TITLE)} - 证明报告</h1>
<p><small>生成时间：{esc(meta.get('created_at') or '')}</small></p>
<div class="section">
  <h2>一、照片基础信息</h2>
  <p>原始路径：<code>{esc(meta.get('photo_original_path') or '')}</code></p>
  <p>复制路径：<code>{esc(meta.get('photo_copied_path') or '')}</code></p>
  <p>SHA-256 摘要：<code>{esc(meta.get('sha256') or '')}</code></p>
  {f"<p>预览：</p><img src='{esc(photo_rel)}' style='max-width:100%;border:1px solid #ddd;' />" if photo_rel else ""}
</div>
<div class="section">
  <h2>二、EXIF 元数据</h2>
  <table class="table">
    <thead><tr><th>字段</th><th>值</th></tr></thead>
    <tbody>{exif_rows}</tbody>
  </table>
</div>
<div class="section">
  <h2>三、不可预测信息</h2>
  {world_block}
  {stock_block}
  <p><small>注：通过引用公开且不可预知的第三方数据（如网络标准时间与股票行情）来佐证生成时间点。</small></p>
</div>
<div class="section">
  <h2>四、发布记录（外部平台）</h2>
  <ul>
    {publish_list}
  </ul>
  <p><small>提示：将同一图片在多平台公开发布可增强举证链。此处记录发布URL以供核验。</small></p>
</div>
<div class="section">
  <h2>五、完整性与可验证性</h2>
  <ol>
    <li>使用 SHA-256 对图片文件取哈希，确保内容未被篡改；第三方可复算核验。</li>
    <li>记录 EXIF 拍摄时间等元数据（若存在），结合外部公开数据佐证。</li>
    <li>通过公开时间源与市场行情等不可预测信息，限定生成时间窗口。</li>
    <li>可访问以上发布URL验证公开发布的时间戳。</li>
  </ol>
</div>
<hr />
<p><small>本报告由 {esc(meta.get('app'))} v{esc(meta.get('version'))} 自动生成。</small></p>
</body>
</html>
"""


def main():
    try:
        root = tk.Tk()
        app = EvidenceApp(root)
        root.mainloop()
    except TclError as e:
        sys.stderr.write("[错误] Tk 初始化失败：" + str(e) + "\n")
        sys.stderr.write("[提示] 你的 Python 可能缺少 Tcl/Tk 运行库。可尝试：\n")
        sys.stderr.write("  - 重新从 python.org 安装官方 Windows 安装包（含 Tcl/Tk）\n")
        sys.stderr.write("  - 或安装 Microsoft Store 版 Python 并确保包含 Tcl/Tk\n")
        sys.stderr.write("  - 若已安装，可检查环境变量 TCL_LIBRARY/TK_LIBRARY 是否指向有效目录\n\n")
        # 命令行降级：支持传入图片路径生成证据包
        if len(sys.argv) > 1:
            photo = sys.argv[1]
            if not os.path.isfile(photo):
                sys.stderr.write(f"[失败] 找不到文件：{photo}\n")
                return
            # 组装数据
            h = sha256_file(photo)
            exif = read_exif(photo)
            world = fetch_world_time()
            ts = datetime.now().isoformat()
            # 打包目录
            ts_dir = datetime.now().strftime("%Y%m%d_%H%M%S")
            pkg = os.path.join(OUTPUT_DIR, f"evidence_{ts_dir}")
            os.makedirs(pkg, exist_ok=True)
            # 复制图片
            name = safe_filename(os.path.basename(photo))
            copied = os.path.join(pkg, name)
            try:
                shutil.copy2(photo, copied)
            except Exception as ce:
                sys.stderr.write(f"[警告] 复制图片失败：{ce}\n")
                copied = None
            # 写元数据
            meta = {
                "photo_original_path": photo,
                "photo_copied_path": copied,
                "sha256": h,
                "exif": exif,
                "world_time": world,
                "stocks": [],  # 新格式：多个股票数据
                "stock": None,  # 兼容性：单个股票数据
                "publish_urls": [],
                "created_at": ts,
                "app": APP_TITLE,
                "version": "1.0.0"
            }
            meta_path = os.path.join(pkg, "evidence.json")
            with open(meta_path, "w", encoding="utf-8") as f:
                json.dump(meta, f, ensure_ascii=False, indent=2)

            # 生成简化 HTML（不依赖 Tk 组件）
            def _build_html(meta: dict) -> str:
                def esc(s: str) -> str:
                    return (s or "").replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
                photo_rel = os.path.basename(meta.get("photo_copied_path") or "")
                exif = meta.get("exif") or {}
                world = meta.get("world_time") or {}
                exif_rows = "".join(
                    f"<tr><td>{esc(str(k))}</td><td>{esc(str(v))}</td></tr>" for k, v in (exif.items() if isinstance(exif, dict) else [])
                ) if exif else "<tr><td colspan=2>无或读取失败</td></tr>"
                world_block = "" if not world else (
                    f"<p><strong>网络UTC时间</strong>：{esc(world.get('utc_datetime',''))}（来源：<a href='{esc(world.get('source',''))}' target='_blank'>worldtimeapi</a>）</p>"
                )
                return f"""
<!DOCTYPE html>
<html lang=zh>
<head>
<meta charset=\"utf-8\" />
<title>{esc(APP_TITLE)} - 证明报告</title>
<style>
body {{ font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, 'Helvetica Neue', Arial, 'Noto Sans', 'PingFang SC', 'Microsoft YaHei', sans-serif; margin: 24px; }}
.table {{ border-collapse: collapse; width: 100%; }}
.table th, .table td {{ border: 1px solid #ddd; padding: 8px; }}
.section {{ margin: 18px 0; }}
small {{ color: #666; }}
</style>
</head>
<body>
<h1>{esc(APP_TITLE)} - 证明报告（命令行模式）</h1>
<p><small>生成时间：{esc(meta.get('created_at') or '')}</small></p>
<div class=\"section\">
  <h2>照片基础信息</h2>
  <p>原始路径：<code>{esc(meta.get('photo_original_path') or '')}</code></p>
  <p>复制路径：<code>{esc(meta.get('photo_copied_path') or '')}</code></p>
  <p>SHA-256 摘要：<code>{esc(meta.get('sha256') or '')}</code></p>
  {f"<p>预览：</p><img src='{esc(photo_rel)}' style='max-width:100%;border:1px solid #ddd;' />" if photo_rel else ""}
</div>
<div class=\"section\">
  <h2>EXIF 元数据</h2>
  <table class=\"table\">
    <thead><tr><th>字段</th><th>值</th></tr></thead>
    <tbody>{exif_rows}</tbody>
  </table>
</div>
<div class=\"section\">
  <h2>外部时间源</h2>
  {world_block}
</div>
<hr />
<p><small>本报告由 {esc(meta.get('app'))} v{esc(meta.get('version'))} 自动生成（命令行模式）。</small></p>
</body>
</html>
"""
            report_path = os.path.join(pkg, "report.html")
            with open(report_path, "w", encoding="utf-8") as f:
                f.write(_build_html(meta))
            sys.stderr.write(f"[完成] 已生成证据包：{pkg}\n")
            sys.stderr.write(f"[提示] 可用浏览器打开：{report_path}\n")
        else:
            sys.stderr.write("[用法] 命令行降级：.\\.venv\\Scripts\\python main.py \\path\\to\\photo.jpg\n")
            sys.stderr.write("[或] 直接：python main.py \\path\\to\\photo.jpg\n")


if __name__ == "__main__":
    main()
