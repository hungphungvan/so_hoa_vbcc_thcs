from __future__ import annotations

import os
import sys
import threading
import queue
import subprocess
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinter.scrolledtext import ScrolledText

from src.lookup_manager import LookupManager, get_data_dir
from src.excel_processor import ExcelProcessor
from src.gui.dialogs import (
    SchoolBlockDialog,
    CohortBlockDialog,
    DateFixDialog,
    GradDateBlockDialog,
    DataInspectorDialog
)
from src.gui.theme import (
    IS_DARK,
    FONT_FAMILY,
    FONT_FAMILY_MONO,
    CLR_PRIMARY,
    CLR_ACCENT,
    CLR_SUCCESS,
    CLR_WARNING,
    CLR_DANGER,
    CLR_MUTED,
    CLR_TEXT,
    CLR_WINDOW_BG,
    CLR_LOG_BG,
    CLR_LOG_FG,
)

class MainWindow:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Số hóa Văn bằng Chứng chỉ THCS")
        
        # Đặt kích thước và vị trí cửa sổ tối ưu đa nền tảng
        sw = self.root.winfo_screenwidth()
        sh = self.root.winfo_screenheight()
        win_w = min(1120, max(960, sw - 60))
        win_h = min(780, max(640, sh - 80))
        x = max(0, (sw - win_w) // 2)
        y = max(0, (sh - win_h) // 2)
        self.root.geometry(f"{win_w}x{win_h}+{x}+{y}")
        self.root.minsize(920, 620)
        self.root.configure(bg=CLR_WINDOW_BG)

        # Cấu hình style
        self._setup_styles()

        # Biến trạng thái
        self.lookup_mgr = None
        self.processor = None
        self.dialog_queue = queue.Queue()
        self.log_queue = queue.Queue()
        self.is_processing = False
        self.cancel_requested = False
        self.all_unit_students = []
        self.last_output_dir = None

        self._build_ui()
        self._init_data()

        # Bắt đầu vòng lặp kiểm tra queue cho thread-safe GUI
        self.root.after(50, self._process_queues)

        # Cập nhật layout và kích hoạt khắc phục lỗi Cocoa redraw lag trên macOS
        self.root.update_idletasks()
        self.root.lift()
        if sys.platform == "darwin":
            self.root.after(80, self._force_macos_redraw)

    def _force_macos_redraw(self):
        """Khắc phục triệt để lỗi macOS Tkinter không vẽ nội dung (trắng/trống cửa sổ) cho đến khi người dùng resize."""
        try:
            w = self.root.winfo_width()
            h = self.root.winfo_height()
            if w > 10 and h > 10:
                self.root.geometry(f"{w+1}x{h+1}")
                self.root.update_idletasks()
                self.root.geometry(f"{w}x{h}")
                self.root.update_idletasks()
                self.root.lift()
        except Exception:
            pass

    def _setup_styles(self):
        style = ttk.Style(self.root)
        # Chỉ áp dụng 'clam' trên Windows/Linux. Trên macOS giữ nguyên 'aqua' nguyên bản để Cocoa vẽ mượt mà.
        if sys.platform != "darwin":
            try:
                style.theme_use('clam')
            except Exception:
                pass

        style.configure("Card.TLabelframe", font=(FONT_FAMILY, 10, "bold"))
        style.configure("Card.TLabelframe.Label", font=(FONT_FAMILY, 10, "bold"), foreground=CLR_PRIMARY)

        style.configure("Action.TButton", font=(FONT_FAMILY, 10, "bold"), padding=6)
        style.configure("Inspect.TButton", font=(FONT_FAMILY, 10, "bold"), foreground=CLR_PRIMARY)
        if sys.platform != "darwin":
            style.configure("Success.TButton", font=(FONT_FAMILY, 10, "bold"), background="#2E7D32", foreground="#FFFFFF")
            style.configure("Stop.TButton", font=(FONT_FAMILY, 10, "bold"), background="#C62828", foreground="#FFFFFF")
        else:
            style.configure("Success.TButton", font=(FONT_FAMILY, 10, "bold"))
            style.configure("Stop.TButton", font=(FONT_FAMILY, 10, "bold"))

    def _build_ui(self):
        # 1. Header Banner
        hdr_frame = tk.Frame(self.root, bg="#1565C0", padx=16, pady=12)
        hdr_frame.pack(fill=tk.X)

        title_lbl = tk.Label(
            hdr_frame,
            text="SỐ HÓA VĂN BẰNG CHỨNG CHỈ THCS",
            bg="#1565C0",
            fg="#FFFFFF",
            font=(FONT_FAMILY, 14, "bold")
        )
        title_lbl.pack(anchor=tk.W)

        sub_lbl = tk.Label(
            hdr_frame,
            text="Chuẩn hóa dữ liệu cấp bằng và lập hồ sơ số hóa",
            bg="#1565C0",
            fg="#E3F2FD",
            font=(FONT_FAMILY, 9)
        )
        sub_lbl.pack(anchor=tk.W, pady=(2, 0))

        # Main Container
        main_container = ttk.Frame(self.root, padding=12)
        main_container.pack(fill=tk.BOTH, expand=True)

        # 2. Card 1: Chọn Đơn vị
        unit_card = ttk.LabelFrame(main_container, text="1. Chọn đơn vị", style="Card.TLabelframe", padding=10)
        unit_card.pack(fill=tk.X, pady=(0, 6))

        row1 = ttk.Frame(unit_card)
        row1.pack(fill=tk.X, pady=2)
        ttk.Label(row1, text="Tìm kiếm đơn vị: ", font=(FONT_FAMILY, 9, "bold")).pack(side=tk.LEFT)

        self.unit_search_var = tk.StringVar()
        self.unit_search_var.trace_add("write", self._on_search_unit)
        self.unit_search_entry = ttk.Entry(row1, textvariable=self.unit_search_var, width=22)
        self.unit_search_entry.pack(side=tk.LEFT, padx=(0, 10))

        ttk.Label(row1, text="Chọn đơn vị: ", font=(FONT_FAMILY, 9, "bold")).pack(side=tk.LEFT)
        self.unit_combo_var = tk.StringVar()
        self.unit_combo = ttk.Combobox(row1, textvariable=self.unit_combo_var, state="readonly", width=50)
        self.unit_combo.pack(side=tk.LEFT, padx=6, fill=tk.X, expand=True)
        self.unit_combo.bind("<<ComboboxSelected>>", self._on_unit_selected)

        row2 = ttk.Frame(unit_card)
        row2.pack(fill=tk.X, pady=(6, 2))

        ttk.Label(row2, text="Mã đơn vị: ", font=(FONT_FAMILY, 9)).pack(side=tk.LEFT)
        self.unit_code_lbl = ttk.Label(row2, text="---", font=(FONT_FAMILY, 9, "bold"), foreground=CLR_PRIMARY)
        self.unit_code_lbl.pack(side=tk.LEFT, padx=(0, 12))

        ttk.Label(row2, text="Tên đơn vị: ", font=(FONT_FAMILY, 9)).pack(side=tk.LEFT)
        self.unit_name_lbl = ttk.Label(row2, text="---", font=(FONT_FAMILY, 9, "bold"), foreground=CLR_PRIMARY)
        self.unit_name_lbl.pack(side=tk.LEFT, padx=(0, 12))

        ttk.Label(row2, text="Xã/phường: ", font=(FONT_FAMILY, 9)).pack(side=tk.LEFT)
        self.unit_commune_lbl = ttk.Label(row2, text="---", font=(FONT_FAMILY, 9, "bold"), foreground=CLR_ACCENT)
        self.unit_commune_lbl.pack(side=tk.LEFT, padx=(0, 12))

        ttk.Label(row2, text="Số trường trực thuộc: ", font=(FONT_FAMILY, 9)).pack(side=tk.LEFT)
        self.school_count_lbl = ttk.Label(row2, text="0 trường", font=(FONT_FAMILY, 9, "bold"), foreground=CLR_SUCCESS)
        self.school_count_lbl.pack(side=tk.LEFT)

        row3 = ttk.Frame(unit_card)
        row3.pack(fill=tk.X, pady=(4, 0))
        ttk.Label(row3, text="Danh sách trường: ", font=(FONT_FAMILY, 8, "italic")).pack(side=tk.LEFT)
        self.schools_display_lbl = ttk.Label(row3, text="(Chưa chọn đơn vị)", font=(FONT_FAMILY, 8), foreground=CLR_MUTED)
        self.schools_display_lbl.pack(side=tk.LEFT, fill=tk.X, expand=True)

        # 3. Card 2: Chọn Tệp hoặc Thư mục nguồn
        file_card = ttk.LabelFrame(main_container, text="2. Tệp hoặc Thư mục nguồn", style="Card.TLabelframe", padding=10)
        file_card.pack(fill=tk.X, pady=(0, 6))

        f_row1 = ttk.Frame(file_card)
        f_row1.pack(fill=tk.X)

        self.path_var = tk.StringVar()
        self.path_entry = ttk.Entry(f_row1, textvariable=self.path_var)
        self.path_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8))

        browse_file_btn = ttk.Button(f_row1, text="📂 Chọn 1 file...", command=self._browse_file)
        browse_file_btn.pack(side=tk.LEFT, padx=3)

        browse_dir_btn = ttk.Button(f_row1, text="📁 Chọn thư mục đơn vị...", command=self._browse_dir)
        browse_dir_btn.pack(side=tk.LEFT, padx=3)

        self.files_count_lbl = ttk.Label(
            file_card,
            text="Hệ thống tự động quét và nhận diện các trường, niên khóa từ thư mục, tên file hoặc sheet.",
            font=(FONT_FAMILY, 8, "italic"),
            foreground=CLR_MUTED
        )
        self.files_count_lbl.pack(anchor=tk.W, pady=(4, 0))

        # 4. Card 3: Tùy chọn Xuất file (Export Options)
        opt_card = ttk.LabelFrame(main_container, text="3. Tùy chọn xuất kết quả", style="Card.TLabelframe", padding=8)
        opt_card.pack(fill=tk.X, pady=(0, 6))

        opt_grid = ttk.Frame(opt_card)
        opt_grid.pack(fill=tk.X)

        # Cột 1: Tùy chọn Gộp
        gop_frame = ttk.LabelFrame(opt_grid, text="Gộp dữ liệu", padding=6)
        gop_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 4))

        self.opt_gop_truong = tk.BooleanVar(value=True)
        ttk.Checkbutton(gop_frame, text="Gộp theo Trường", variable=self.opt_gop_truong).pack(anchor=tk.W, pady=3)

        self.opt_gop_nam = tk.BooleanVar(value=False)
        ttk.Checkbutton(gop_frame, text="Gộp theo Năm", variable=self.opt_gop_nam).pack(anchor=tk.W, pady=3)

        # Cột 2: Tùy chọn Tách
        tach_frame = ttk.LabelFrame(opt_grid, text="Tách dữ liệu", padding=6)
        tach_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=4)

        self.opt_tach_truong = tk.BooleanVar(value=True)
        ttk.Checkbutton(tach_frame, text="Tách theo Trường", variable=self.opt_tach_truong).pack(anchor=tk.W, pady=2)

        self.opt_tach_nam = tk.BooleanVar(value=False)
        ttk.Checkbutton(tach_frame, text="Tách theo Năm", variable=self.opt_tach_nam).pack(anchor=tk.W, pady=2)

        self.opt_tach_truong_nam = tk.BooleanVar(value=False)
        ttk.Checkbutton(tach_frame, text="Tách theo Trường & Năm", variable=self.opt_tach_truong_nam).pack(anchor=tk.W, pady=2)

        # Cột 3: Mô-đun 2: Danh sách tài liệu số hóa
        doc_frame = ttk.LabelFrame(opt_grid, text="Danh mục hồ sơ số hóa", padding=6)
        doc_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(4, 0))

        self.opt_doc_truong = tk.BooleanVar(value=True)
        ttk.Checkbutton(doc_frame, text="Tài liệu số hóa theo Trường", variable=self.opt_doc_truong).pack(anchor=tk.W, pady=2)

        self.opt_doc_nam = tk.BooleanVar(value=False)
        ttk.Checkbutton(doc_frame, text="Tài liệu số hóa theo Năm", variable=self.opt_doc_nam).pack(anchor=tk.W, pady=2)

        self.opt_doc_truong_nam = tk.BooleanVar(value=False)
        ttk.Checkbutton(doc_frame, text="Tài liệu số hóa theo Trường & Năm", variable=self.opt_doc_truong_nam).pack(anchor=tk.W, pady=2)

        # 5. Card 4: Tiến trình & Điều khiển
        prog_card = ttk.LabelFrame(main_container, text="4. Tiến trình xử lý", style="Card.TLabelframe", padding=8)
        prog_card.pack(fill=tk.X, pady=(0, 6))

        self.progress_bar = ttk.Progressbar(prog_card, orient=tk.HORIZONTAL, mode='determinate')
        self.progress_bar.pack(fill=tk.X, pady=(0, 6))

        ctrl_row = ttk.Frame(prog_card)
        ctrl_row.pack(fill=tk.X)

        self.status_lbl = ttk.Label(ctrl_row, text="Sẵn sàng thực hiện.", font=(FONT_FAMILY, 9, "bold"), foreground=CLR_PRIMARY)
        self.status_lbl.pack(side=tk.LEFT)

        self.stop_btn = ttk.Button(ctrl_row, text="⏹ Dừng", state=tk.DISABLED, command=self._stop_processing)
        self.stop_btn.pack(side=tk.RIGHT, padx=4)

        self.start_btn = ttk.Button(ctrl_row, text="▶️ Bắt đầu xử lý", style="Action.TButton", command=self._start_processing)
        self.start_btn.pack(side=tk.RIGHT, padx=4)

        # 6. Card 5: Thống kê & Log Console
        log_card = ttk.LabelFrame(main_container, text="5. Thống kê & Nhật ký hoạt động", style="Card.TLabelframe", padding=8)
        log_card.pack(fill=tk.BOTH, expand=True, pady=(0, 6))

        stat_row = ttk.Frame(log_card)
        stat_row.pack(fill=tk.X, pady=(0, 4))

        self.stat_total_lbl = ttk.Label(stat_row, text="Tổng học sinh: 0", font=(FONT_FAMILY, 9, "bold"))
        self.stat_total_lbl.pack(side=tk.LEFT, padx=(0, 20))

        self.stat_da_so_hoa_lbl = ttk.Label(stat_row, text="✅ Đã số hóa đầy đủ: 0", font=(FONT_FAMILY, 9, "bold"), foreground=CLR_SUCCESS)
        self.stat_da_so_hoa_lbl.pack(side=tk.LEFT, padx=(0, 20))

        self.stat_chua_day_du_lbl = ttk.Label(stat_row, text="⚠️ Số hoá chưa đầy đủ: 0", font=(FONT_FAMILY, 9, "bold"), foreground=CLR_WARNING)
        self.stat_chua_day_du_lbl.pack(side=tk.LEFT, padx=(0, 20))

        self.stat_chua_so_hoa_lbl = ttk.Label(stat_row, text="❌ Chưa số hoá: 0", font=(FONT_FAMILY, 9, "bold"), foreground=CLR_DANGER)
        self.stat_chua_so_hoa_lbl.pack(side=tk.LEFT)

        # Log text console
        self.log_text = ScrolledText(
            log_card,
            height=8,
            font=(FONT_FAMILY_MONO, 9),
            wrap=tk.WORD,
            bg=CLR_LOG_BG,
            fg=CLR_LOG_FG,
            insertbackground=CLR_LOG_FG
        )
        self.log_text.pack(fill=tk.BOTH, expand=True)

        self.log_text.tag_config("info", foreground=CLR_LOG_FG)
        self.log_text.tag_config("success", foreground=CLR_SUCCESS, font=(FONT_FAMILY_MONO, 9, "bold"))
        self.log_text.tag_config("warning", foreground=CLR_WARNING)
        self.log_text.tag_config("error", foreground=CLR_DANGER, font=(FONT_FAMILY_MONO, 9, "bold"))

        # 7. Bottom Bar
        bottom_frame = ttk.Frame(self.root, padding=(12, 6))
        bottom_frame.pack(fill=tk.X)

        self.open_output_btn = ttk.Button(bottom_frame, text="📂 Mở thư mục kết quả", command=self._open_output_dir)
        self.open_output_btn.pack(side=tk.LEFT, padx=(0, 6))

        self.inspect_btn = ttk.Button(
            bottom_frame,
            text="🔍 Tra cứu & Đối soát dữ liệu",
            command=self._open_data_inspector,
            state=tk.DISABLED
        )
        self.inspect_btn.pack(side=tk.LEFT, padx=6)

        exit_btn = ttk.Button(bottom_frame, text="Thoát", command=self.root.destroy)
        exit_btn.pack(side=tk.RIGHT)

    def _init_data(self):
        """Khởi tạo LookupManager và nạp danh mục tra cứu."""
        try:
            self.lookup_mgr = LookupManager()
            self.processor = ExcelProcessor(self.lookup_mgr)
            self._update_unit_combobox(self.lookup_mgr.unit_names_list)
            self._log(f"📖 Nguồn danh mục: {self.lookup_mgr.excel_path}", "info")
            self._log(f"✅ Đã nạp thành công danh mục tra cứu: {len(self.lookup_mgr.unit_names_list)} đơn vị (kèm Xã/phường & Mã trường).", "success")
        except Exception as e:
            self._log(f"❌ Lỗi nạp danh mục Danh_Muc_Tra_Cuu.xlsx: {e}", "error")
            messagebox.showerror("Lỗi nạp danh mục", f"Không thể đọc file danh mục tra cứu:\n{e}", parent=self.root)

    def _update_unit_combobox(self, unit_list: list[str]):
        self.unit_combo["values"] = unit_list
        if unit_list:
            self.unit_combo.current(0)
            self._on_unit_selected()

    def _on_search_unit(self, *args):
        """Lọc danh sách đơn vị khi người dùng gõ vào ô tìm kiếm (hỗ trợ tìm tên, xã, hoặc mã)."""
        if not self.lookup_mgr:
            return
        kw = self.unit_search_var.get().strip().lower()
        if not kw:
            filtered = self.lookup_mgr.unit_names_list
        else:
            filtered = [u for u in self.lookup_mgr.unit_names_list if kw in u.lower()]
        self.unit_combo["values"] = filtered
        if filtered:
            self.unit_combo.current(0)
            self._on_unit_selected()

    def _on_unit_selected(self, event=None):
        """Khi người dùng chọn một đơn vị từ combobox."""
        selected_unit = self.unit_combo_var.get()
        if not selected_unit or not self.lookup_mgr:
            return
        info = self.lookup_mgr.get_unit_info(selected_unit)
        if info:
            self.unit_code_lbl.config(text=info.get("code", "---"))
            self.unit_name_lbl.config(text=info.get("name", "---"))
            self.unit_commune_lbl.config(text=info.get("commune", "(Trống)") or "(Trống)")
            sch_list = info.get("schools", [])
            self.school_count_lbl.config(text=f"{len(sch_list)} trường")
            if sch_list:
                disp = "; ".join(sch_list)
                if len(disp) > 130:
                    disp = disp[:127] + "..."
                self.schools_display_lbl.config(text=disp)
            else:
                self.schools_display_lbl.config(text="(Không có trường trực thuộc trong danh mục)")

    def _browse_file(self):
        file_selected = filedialog.askopenfilename(
            title="Chọn file Excel nguồn cần chuẩn hóa",
            filetypes=[("Excel Files", "*.xlsx"), ("All Files", "*.*")],
            parent=self.root
        )
        if file_selected:
            self.path_var.set(file_selected)
            self.files_count_lbl.config(text=f"Đã chọn 1 tệp: {Path(file_selected).name}")

    def _browse_dir(self):
        dir_selected = filedialog.askdirectory(
            title="Chọn thư mục Đơn vị chứa các file Excel nguồn",
            parent=self.root
        )
        if dir_selected:
            self.path_var.set(dir_selected)
            p = Path(dir_selected)
            files = [
                f for f in p.rglob("*.xlsx")
                if not f.name.startswith("~$")
                and not f.name.startswith("._")
                and not any(tag in f.name for tag in ["_Đa_Gop_", "_DS_Loi", "_Du_Lieu_Goc", "_DS_Lỗi", "_TongHop", "_DS_Tai_Lieu_So_Hoa"])
                and not any(tag in part.lower() for tag in ["da_chuan_hoa", "ket_qua_so_hoa"] for part in f.parts)
            ]
            self.files_count_lbl.config(
                text=f"Tìm thấy {len(files)} tệp Excel trong thư mục {p.name} (bao gồm cả các thư mục trường con)"
            )

    def _log(self, msg: str, level: str = "info"):
        self.log_queue.put((msg, level))

    def _process_queues(self):
        """Xử lý hàng đợi log và hàng đợi dialog tương tác trên Main GUI Thread."""
        while not self.log_queue.empty():
            msg, level = self.log_queue.get_nowait()
            self.log_text.insert(tk.END, msg + "\n", level)
            self.log_text.see(tk.END)

        if not self.dialog_queue.empty():
            req = self.dialog_queue.get_nowait()
            d_type = req["type"]
            args = req["args"]
            res_box = req["res_box"]
            evt = req["event"]

            try:
                if d_type == "school_block":
                    block_info, allowed_schools, sugg_school = args
                    dlg = SchoolBlockDialog(self.root, block_info, allowed_schools, sugg_school)
                    self.root.wait_window(dlg)
                    res_box["result"] = (dlg.result, dlg.apply_to_all_similar)
                elif d_type == "cohort_block":
                    block_info = args[0]
                    default_y_ra = args[1] if len(args) > 1 else ""
                    reason = args[2] if len(args) > 2 else ""
                    dlg = CohortBlockDialog(self.root, block_info, default_y_ra, reason)
                    self.root.wait_window(dlg)
                    res_box["result"] = (dlg.result, dlg.apply_to_all_similar)
                elif d_type == "date_fix":
                    student_info, field_name, sugg_dob = args
                    dlg = DateFixDialog(self.root, student_info, field_name, sugg_dob)
                    self.root.wait_window(dlg)
                    res_box["result"] = (dlg.result, dlg.skip_all)
                elif d_type == "grad_date_block":
                    block_info, sugg_grad = args
                    dlg = GradDateBlockDialog(self.root, block_info, sugg_grad)
                    self.root.wait_window(dlg)
                    res_box["result"] = (dlg.result, dlg.apply_to_group)
            except Exception as e:
                print(f"Lỗi hiển thị dialog: {e}")
                res_box["result"] = (None, False)
            finally:
                evt.set()

        self.root.after(50, self._process_queues)

    def _dialog_callback(self, dialog_type: str, *args):
        """Hàm callback được gọi từ Worker Thread để mở Dialog GUI trên Main Thread."""
        evt = threading.Event()
        res_box = {"result": None}
        self.dialog_queue.put({
            "type": dialog_type,
            "args": args,
            "res_box": res_box,
            "event": evt
        })
        evt.wait()
        return res_box["result"]

    def _start_processing(self):
        if self.is_processing:
            return

        selected_unit = self.unit_combo_var.get()
        if not selected_unit or not self.lookup_mgr:
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn Đơn vị cấp bằng trước khi bắt đầu!", parent=self.root)
            return

        unit_info = self.lookup_mgr.get_unit_info(selected_unit)
        if not unit_info:
            messagebox.showwarning("Cảnh báo", "Đơn vị đã chọn không tồn tại trong danh mục!", parent=self.root)
            return

        raw_path = self.path_var.get().strip().strip("'\"")
        if not raw_path or not Path(raw_path).exists():
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn tệp hoặc thư mục Excel hợp lệ!", parent=self.root)
            return

        target_path = Path(raw_path)
        if target_path.is_file():
            files_to_process = [target_path]
            output_dir = target_path.parent / "Ket_Qua_So_Hoa"
        else:
            files_to_process = sorted(
                [
                    f for f in target_path.rglob("*.xlsx")
                    if not f.name.startswith("~$")
                    and not f.name.startswith("._")
                    and not any(tag in f.name for tag in ["_Đa_Gop_", "_DS_Loi", "_Du_Lieu_Goc", "_DS_Lỗi", "_TongHop", "_DS_Tai_Lieu_So_Hoa"])
                    and not any(tag in part.lower() for tag in ["da_chuan_hoa", "ket_qua_so_hoa"] for part in f.parts)
                ],
                key=lambda x: str(x).lower()
            )
            output_dir = target_path / "Ket_Qua_So_Hoa"

        if not files_to_process:
            messagebox.showwarning("Cảnh báo", "Không tìm thấy file Excel (*.xlsx) nào để xử lý!", parent=self.root)
            return

        export_options = {
            "gop_theo_truong": self.opt_gop_truong.get(),
            "gop_theo_nam": self.opt_gop_nam.get(),
            "tach_theo_truong": self.opt_tach_truong.get(),
            "tach_theo_nam": self.opt_tach_nam.get(),
            "tach_theo_truong_va_nam": self.opt_tach_truong_nam.get(),
            "doc_theo_truong": self.opt_doc_truong.get(),
            "doc_theo_nam": self.opt_doc_nam.get(),
            "doc_theo_truong_va_nam": self.opt_doc_truong_nam.get()
        }

        self.is_processing = True
        self.cancel_requested = False
        self.start_btn.config(state=tk.DISABLED)
        self.stop_btn.config(state=tk.NORMAL)
        self.inspect_btn.config(state=tk.DISABLED)
        self.progress_bar["value"] = 0
        self.progress_bar["maximum"] = len(files_to_process)
        self.last_output_dir = output_dir

        # Reset thống kê
        self.stat_total_lbl.config(text="Tổng học sinh: 0")
        self.stat_da_so_hoa_lbl.config(text="✅ Đã số hóa đầy đủ: 0")
        self.stat_chua_day_du_lbl.config(text="⚠️ Số hoá chưa đầy đủ: 0")
        self.stat_chua_so_hoa_lbl.config(text="❌ Chưa số hoá: 0")

        # Khởi chạy worker thread
        t = threading.Thread(
            target=self._worker_thread,
            args=(files_to_process, unit_info, output_dir, export_options),
            daemon=True
        )
        t.start()

    def _stop_processing(self):
        if self.is_processing:
            ans = messagebox.askyesno("Xác nhận", "Bạn có chắc chắn muốn dừng quá trình xử lý không?", parent=self.root)
            if ans:
                self.cancel_requested = True
                self._log("⚠️ Đã gửi yêu cầu dừng xử lý...", "warning")

    def _worker_thread(
        self,
        files: list[Path],
        unit_info: dict,
        output_dir: Path,
        export_options: dict
    ):
        unit_code = unit_info.get("code", "")
        unit_name = unit_info.get("name", "")
        unit_commune = unit_info.get("commune", "")
        allowed_schools = unit_info.get("schools", [])

        self._log("\n" + "=" * 80, "info")
        self._log(f"🚀 BẮT ĐẦU CHUẨN HÓA DỮ LIỆU ĐƠN VỊ: {unit_name} [Mã: {unit_code}]", "info")
        if unit_commune:
            self._log(f"📍 Địa bàn Xã/phường: {unit_commune}", "info")
        self._log(f"📁 Tổng số file Excel phát hiện: {len(files)} tệp.", "info")
        self._log(f"🏫 Số trường trực thuộc đơn vị: {len(allowed_schools)} trường.", "info")
        self._log("=" * 80, "info")

        def progress_cb(curr_idx, total_cnt, current_stats, stage_msg=None):
            self.root.after(0, lambda: self._update_progress_ui(curr_idx, total_cnt, current_stats, stage_msg))

        def check_cancel_cb():
            return self.cancel_requested

        all_students, overall_stats = self.processor.process_all_files(
            files=files,
            unit_info=unit_info,
            output_dir=output_dir,
            export_options=export_options,
            dialog_callback=self._dialog_callback,
            log_callback=self._log,
            progress_callback=progress_cb,
            check_cancel=check_cancel_cb
        )

        self.all_unit_students = all_students

        total_st = overall_stats["total_students"]
        da_so_hoa = overall_stats["da_so_hoa"]
        chua_day_du = overall_stats["chua_day_du"]
        chua_so_hoa = overall_stats["chua_so_hoa"]

        self._log("\n" + "=" * 80, "info")
        self._log(f"🎉 HOÀN THÀNH XỬ LÝ {len(files)} TỆP EXCEL CHO ĐƠN VỊ {unit_name}!", "success")
        self._log(f"✅ Tổng số học sinh: {total_st}", "info")
        self._log(f"   • Đã số hóa đầy đủ, chính xác: {da_so_hoa}", "success")
        self._log(f"   • Số hoá chưa đầy đủ, chính xác: {chua_day_du}", "warning")
        self._log(f"   • Chưa số hoá: {chua_so_hoa}", "error")
        self._log(f"📁 Thư mục kết quả: {output_dir.absolute()}", "success")
        self._log("=" * 80, "info")

        self.root.after(0, self._on_processing_finished, total_st, da_so_hoa, chua_day_du, chua_so_hoa, output_dir)

    def _update_progress_ui(self, curr_idx, total_cnt, stats, stage_msg=None):
        self.progress_bar["maximum"] = total_cnt
        self.progress_bar["value"] = curr_idx
        if stage_msg:
            self.status_lbl.config(text=stage_msg)
        else:
            self.status_lbl.config(text=f"Đang xử lý [{curr_idx}/{total_cnt}] tệp...")
        self.stat_total_lbl.config(text=f"Tổng học sinh: {stats.get('total_students', 0)}")
        self.stat_da_so_hoa_lbl.config(text=f"✅ Đã số hóa đầy đủ: {stats.get('da_so_hoa', 0)}")
        self.stat_chua_day_du_lbl.config(text=f"⚠️ Số hoá chưa đầy đủ: {stats.get('chua_day_du', 0)}")
        self.stat_chua_so_hoa_lbl.config(text=f"❌ Chưa số hoá: {stats.get('chua_so_hoa', 0)}")

    def _on_processing_finished(self, total_st, da_so_hoa, chua_day_du, chua_so_hoa, output_dir):
        self.status_lbl.config(text="Hoàn thành chuẩn hóa dữ liệu & xuất danh sách tài liệu số hóa!")
        self.is_processing = False
        self.start_btn.config(state=tk.NORMAL)
        self.stop_btn.config(state=tk.DISABLED)
        if self.all_unit_students:
            self.inspect_btn.config(state=tk.NORMAL)

        messagebox.showinfo(
            "Hoàn thành",
            f"Đã hoàn thành chuẩn hóa dữ liệu đơn vị!\n\n"
            f"• Tổng học sinh: {total_st}\n"
            f"• Đã số hóa đầy đủ: {da_so_hoa}\n"
            f"• Số hoá chưa đầy đủ: {chua_day_du}\n"
            f"• Chưa số hoá: {chua_so_hoa}\n\n"
            f"Tất cả file kết quả & Danh sách tài liệu số hóa đã xuất tại:\n{output_dir.name}/\n\n"
            f"Bạn có thể bấm '🔍 Tra cứu & Đối soát dữ liệu' để kiểm tra chi tiết!",
            parent=self.root
        )

    def _open_data_inspector(self):
        """Mở cửa sổ tra cứu và đối soát dữ liệu gốc song song."""
        if not self.all_unit_students:
            messagebox.showinfo("Thông báo", "Chưa có dữ liệu học sinh nào được xử lý.", parent=self.root)
            return
        DataInspectorDialog(self.root, self.all_unit_students)

    def _open_output_dir(self):
        """Mở thư mục kết quả trên hệ điều hành."""
        out_path = getattr(self, "last_output_dir", None)
        if not out_path or not out_path.exists():
            raw_path = self.path_var.get().strip().strip("'\"")
            if raw_path:
                p = Path(raw_path)
                cand = (p.parent if p.is_file() else p) / "Ket_Qua_So_Hoa"
                if not cand.exists():
                    cand = (p.parent if p.is_file() else p) / "Da_Chuan_Hoa"
                if cand.exists():
                    out_path = cand

        if not out_path or not out_path.exists():
            messagebox.showinfo("Thông báo", "Chưa có thư mục kết quả Ket_Qua_So_Hoa nào được tạo.", parent=self.root)
            return

        try:
            if sys.platform == "win32":
                os.startfile(str(out_path))
            elif sys.platform == "darwin":
                subprocess.Popen(["open", str(out_path)])
            else:
                subprocess.Popen(["xdg-open", str(out_path)])
        except Exception as e:
            messagebox.showerror("Lỗi mở thư mục", f"Không thể mở thư mục: {e}", parent=self.root)

def main():
    root = tk.Tk()
    app = MainWindow(root)
    root.mainloop()

if __name__ == "__main__":
    main()
