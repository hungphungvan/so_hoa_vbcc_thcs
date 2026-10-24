from __future__ import annotations

import os
import sys
import subprocess
import threading
import queue
import traceback
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinter.scrolledtext import ScrolledText
import openpyxl

from src.gui.theme import (
    get_theme_palette, IS_DARK,
    FONT_TITLE, FONT_SUBTITLE, FONT_LABEL, FONT_LABEL_BOLD,
    FONT_LOG, FONT_MUTED
)
from src.exporter import sanitize_filename
from src.custom_excel_manager import (
    detect_header_and_columns,
    ExcelSplitter,
    ExcelMerger
)
from src.smart_merger import (
    SmartMerger,
    MappingProfileManager,
    TARGET_COLUMNS
)
from src.gui.column_mapping_dialog import ColumnMappingDialog
from src.file_converter import FileFormatConverter
from src.data_inspector import DataInspector
from src.ocr_processor import (
    OCRProcessor,
    OCRConfigManager,
    DEFAULT_BASE_URL,
    DEFAULT_MODEL_NAME,
)

def open_path_in_os(path: Path):
    """Mở file hoặc thư mục bằng trình quản lý tệp của hệ điều hành (macOS Finder / Windows Explorer)."""
    p_str = str(path.resolve())
    if sys.platform == "darwin":
        subprocess.run(["open", p_str], check=False)
    elif sys.platform == "win32":
        os.startfile(p_str)
    else:
        subprocess.run(["xdg-open", p_str], check=False)


class CustomToolsView(ttk.Frame):
    """
    Giao diện Tab 3: CÔNG CỤ - Tiện ích Tách & Gộp Excel theo Cột tùy chọn.
    Gồm 2 Tab:
    - Tab 1: 📤 Tách tệp theo cột tùy chọn
    - Tab 2: 📥 Nhập / Gộp nhiều tệp có cùng cấu trúc tiêu đề
    """
    def __init__(self, parent, is_dark: bool | None = None):
        super().__init__(parent, style="Window.TFrame")
        self.is_dark = IS_DARK if is_dark is None else is_dark
        self.palette = get_theme_palette(self.is_dark)

        # Queue log và tiến trình
        self.msg_queue = queue.Queue()
        self.is_processing = False
        self.cancel_requested = False

        # Biến Tab 1 (Tách file)
        self.split_file_var = tk.StringVar()
        self.split_sheet_var = tk.StringVar()
        self.split_header_row_var = tk.StringVar(value="2")
        self.split_stt_col_var = tk.StringVar()
        self.split_name_col_var = tk.StringVar()
        self.split_target_col_var = tk.StringVar()
        self.split_renumber_var = tk.BooleanVar(value=True)
        self.split_columns_data: list[dict] = []
        self.last_split_output_dir: Path | None = None

        # Biến Tab 2 (Gộp file)
        self.merge_files_list: list[Path] = []
        self.merge_out_var = tk.StringVar()
        self.merge_renumber_var = tk.BooleanVar(value=True)
        self.merge_source_col_var = tk.BooleanVar(value=True)
        self.merge_all_sheets_var = tk.BooleanVar(value=True)
        self.last_merge_output_file: Path | None = None

        # Biến Tab 1 (Ghép & Chuẩn hóa vào Mẫu chuẩn 28 cột - SmartMerger)
        self.smart_merger = SmartMerger()
        self.smart_files_list: list[Path] = []
        self.smart_out_var = tk.StringVar()
        self.smart_renumber_var = tk.BooleanVar(value=True)
        self.smart_fill_defaults_var = tk.BooleanVar(value=True)
        self.smart_all_sheets_var = tk.BooleanVar(value=True)
        self.last_smart_output_file: Path | None = None
        self.smart_cancel_requested = False

        # Biến Tab 2 (Kiểm tra & Chuẩn hoá dữ liệu 28 cột)
        self.inspector = DataInspector()
        self.inspect_files_list: list[Path] = []
        self.inspect_out_dir_var = tk.StringVar()
        self.inspect_auto_open_var = tk.BooleanVar(value=True)
        self.last_inspect_output_file: Path | None = None
        self.inspect_cancel_requested = False

        # Biến Tab 4-1 (Chuyển đổi định dạng tệp cũ .doc, .xls)
        self.convert_files_list: list[Path] = []
        self.convert_use_custom_dir = tk.BooleanVar(value=False)
        self.convert_custom_dir_var = tk.StringVar()
        self.convert_delete_source_var = tk.BooleanVar(value=False)
        self.convert_skip_existing_var = tk.BooleanVar(value=True)
        self.last_convert_dir: Path | None = None

        # Biến Tab 4-2 (Chuyển VNI, TCVN -> UNICODE)
        self.encoding_files_list: list[Path] = []
        self.encoding_overwrite_var = tk.BooleanVar(value=False)
        self.last_encoding_output_file: Path | None = None
        self.encoding_cancel_requested = False

        # Biến Tab 5 (OCR trích xuất ảnh sổ sang Excel)
        self.ocr_config_mgr = OCRConfigManager()
        self.ocr_processor = OCRProcessor(config_mgr=self.ocr_config_mgr)
        ocr_init_cfg = self.ocr_config_mgr.load_config()
        self.ocr_api_key_var = tk.StringVar(value=ocr_init_cfg.get("api_key", ""))
        self.ocr_base_url_var = tk.StringVar(value=ocr_init_cfg.get("base_url", DEFAULT_BASE_URL))
        self.ocr_model_var = tk.StringVar(value=ocr_init_cfg.get("model_name", DEFAULT_MODEL_NAME))
        self.ocr_show_key = False
        init_prompt, init_prompt_src = self.ocr_config_mgr.resolve_prompt_for_folder(None)
        self.ocr_prompt_text: str = init_prompt
        self.ocr_prompt_source_var = tk.StringVar(value=init_prompt_src)
        self.ocr_images_list: list[Path] = []
        self.ocr_out_var = tk.StringVar()
        self.last_ocr_output_file: Path | None = None
        self.ocr_cancel_requested = False

        # Quản lý tệp vừa được tạo mới nhất toàn hệ thống
        self.last_created_file: Path | None = None

        self._build_ui()
        self._start_queue_poller()

    def _build_ui(self):
        p = self.palette

        # 1. Header Banner
        self.hdr_frame = tk.Frame(self, bg=p.HEADER_BANNER_BG, padx=16, pady=8)
        self.hdr_frame.pack(fill=tk.X)

        self.title_lbl = tk.Label(
            self.hdr_frame,
            text="🛠️ SOHOATHCS - HỆ THỐNG XỬ LÝ & SỐ HÓA VĂN BẰNG THCS",
            font=FONT_TITLE,
            bg=p.HEADER_BANNER_BG,
            fg=p.HEADER_BANNER_FG
        )
        self.title_lbl.pack(anchor=tk.W)

        # 2. Main Container
        main_container = ttk.Frame(self, style="Window.TFrame", padding=6)
        main_container.pack(fill=tk.BOTH, expand=True)

        # 3. Vùng Console Log chung ở dưới cùng (GHIM TRƯỚC VỚI side=tk.BOTTOM ĐỂ ĐẢM BẢO ĐỦ CHIỀU CAO HIỂN THỊ)
        log_frame = ttk.LabelFrame(main_container, text="Nhật ký hoạt động & Thông báo", padding=4)
        log_frame.pack(side=tk.BOTTOM, fill=tk.X, expand=False, pady=(4, 0))

        self.log_text = ScrolledText(
            log_frame,
            height=7,
            font=FONT_LOG,
            bg=p.LOG_BG,
            fg=p.LOG_FG,
            insertbackground=p.LOG_FG,
            wrap=tk.WORD
        )
        self.log_text.pack(fill=tk.BOTH, expand=True)
        self.log_text.tag_config("info", foreground=p.LOG_FG)
        self.log_text.tag_config("success", foreground=p.SUCCESS)
        self.log_text.tag_config("warning", foreground=p.WARNING)
        self.log_text.tag_config("error", foreground=p.DANGER)

        # 4. Main Notebook (5 Tabs chính - chiếm phần diện tích phía trên)
        self.notebook = ttk.Notebook(main_container)
        self.notebook.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

        # Tab 1: Gộp vào mẫu chuẩn 28 cột
        self.tab_smart = ttk.Frame(self.notebook, padding=6)
        self.notebook.add(self.tab_smart, text="  📊 1. GỘP VÀO MẪU CHUẨN  ")
        self._build_tab_smart(self.tab_smart)

        # Tab 2: Kiểm tra & Chuẩn hoá (Mới)
        self.tab_inspect = ttk.Frame(self.notebook, padding=6)
        self.notebook.add(self.tab_inspect, text="  🔍 2. KIỂM TRA & CHUẨN HOÁ  ")
        self._build_tab_inspect(self.tab_inspect)

        # Tab 3: Tách / Ghép
        self.tab_split_merge = ttk.Frame(self.notebook, padding=4)
        self.notebook.add(self.tab_split_merge, text="  📤 3. TÁCH / GHÉP  ")
        sub_nb_sm = ttk.Notebook(self.tab_split_merge)
        sub_nb_sm.pack(fill=tk.BOTH, expand=True)

        self.tab_split = ttk.Frame(sub_nb_sm, padding=4)
        sub_nb_sm.add(self.tab_split, text="  📤 3.1. Tách tệp theo cột  ")
        self._build_tab_split(self.tab_split)

        self.tab_merge = ttk.Frame(sub_nb_sm, padding=4)
        sub_nb_sm.add(self.tab_merge, text="  📥 3.2. Ghép cùng cấu trúc  ")
        self._build_tab_merge(self.tab_merge)

        # Tab 4: Đổi đuôi & Chuyển mã
        self.tab_convert_enc = ttk.Frame(self.notebook, padding=4)
        self.notebook.add(self.tab_convert_enc, text="  🔄 4. ĐỔI ĐUÔI & CHUYỂN MÃ  ")
        sub_nb_ce = ttk.Notebook(self.tab_convert_enc)
        sub_nb_ce.pack(fill=tk.BOTH, expand=True)

        self.tab_convert = ttk.Frame(sub_nb_ce, padding=4)
        sub_nb_ce.add(self.tab_convert, text="  📄 4.1. Chuyển đổi sang docx/xlsx  ")
        self._build_tab_convert(self.tab_convert)

        self.tab_encoding = ttk.Frame(sub_nb_ce, padding=4)
        sub_nb_ce.add(self.tab_encoding, text="  🔤 4.2. Chuyển VNI, TCVN ➔ UNICODE  ")
        self._build_tab_encoding(self.tab_encoding)

        # Tab 5: OCR trích xuất ảnh sổ sang Excel
        self.tab_ocr = ttk.Frame(self.notebook, padding=6)
        self.notebook.add(self.tab_ocr, text="  📷 5. OCR  ")
        self._build_tab_ocr(self.tab_ocr)

    def apply_theme(self, is_dark: bool):
        """Cập nhật giao diện Sáng / Tối."""
        self.is_dark = is_dark
        self.palette = get_theme_palette(self.is_dark)
        p = self.palette
        if hasattr(self, 'hdr_frame'):
            self.hdr_frame.configure(bg=p.HEADER_BANNER_BG)
        if hasattr(self, 'title_lbl'):
            self.title_lbl.configure(bg=p.HEADER_BANNER_BG, fg=p.HEADER_BANNER_FG)
        if hasattr(self, 'log_text'):
            self.log_text.configure(
                bg=p.LOG_BG,
                fg=p.LOG_FG,
                insertbackground=p.LOG_FG
            )
            self.log_text.tag_config("info", foreground=p.LOG_FG)
            self.log_text.tag_config("success", foreground=p.SUCCESS)
            self.log_text.tag_config("warning", foreground=p.WARNING)
            self.log_text.tag_config("error", foreground=p.DANGER)

    # =========================================================================
    # TAB 1: TÁCH FILE THEO CỘT
    # =========================================================================
    def _build_tab_split(self, parent):
        p = self.palette

        # Khung 1: Chọn file
        f_card = ttk.LabelFrame(parent, text="1. Chọn tệp Excel nguồn cần tách", padding=8)
        f_card.pack(fill=tk.X, pady=(0, 6))

        r1 = ttk.Frame(f_card)
        r1.pack(fill=tk.X)
        ttk.Label(r1, text="Tệp Excel nguồn:", font=FONT_LABEL_BOLD).pack(side=tk.LEFT, padx=(0, 6))
        self.split_entry = ttk.Entry(r1, textvariable=self.split_file_var, font=FONT_LABEL)
        self.split_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 6))
        ttk.Button(r1, text="📂 Chọn tệp...", command=self._on_choose_split_file).pack(side=tk.LEFT)

        r2 = ttk.Frame(f_card)
        r2.pack(fill=tk.X, pady=(6, 0))
        ttk.Label(r2, text="Sheet dữ liệu:", font=FONT_LABEL).pack(side=tk.LEFT, padx=(0, 6))
        self.split_sheet_cb = ttk.Combobox(r2, textvariable=self.split_sheet_var, state="readonly", width=25)
        self.split_sheet_cb.pack(side=tk.LEFT, padx=(0, 16))
        self.split_sheet_cb.bind("<<ComboboxSelected>>", lambda e: self._analyze_split_file())

        self.split_file_info_lbl = ttk.Label(r2, text="", font=FONT_MUTED, foreground=p.PRIMARY)
        self.split_file_info_lbl.pack(side=tk.LEFT)

        # Khung 2: Cấu hình dòng tiêu đề & Cột làm căn cứ tách
        col_card = ttk.LabelFrame(parent, text="2. Cấu hình Cột & Điều kiện dòng hợp lệ", padding=8)
        col_card.pack(fill=tk.X, pady=(0, 6))

        c_row1 = ttk.Frame(col_card)
        c_row1.pack(fill=tk.X)

        ttk.Label(c_row1, text="Dòng tiêu đề:", font=FONT_LABEL).pack(side=tk.LEFT, padx=(0, 4))
        self.hdr_row_sp = ttk.Spinbox(c_row1, from_=1, to=50, width=5, textvariable=self.split_header_row_var)
        self.hdr_row_sp.pack(side=tk.LEFT, padx=(0, 16))
        self.hdr_row_sp.bind("<FocusOut>", lambda e: self._analyze_split_file())
        self.hdr_row_sp.bind("<Return>", lambda e: self._analyze_split_file())

        ttk.Label(c_row1, text="Cột STT:", font=FONT_LABEL).pack(side=tk.LEFT, padx=(0, 4))
        self.split_stt_cb = ttk.Combobox(c_row1, textvariable=self.split_stt_col_var, state="readonly", width=18)
        self.split_stt_cb.pack(side=tk.LEFT, padx=(0, 16))
        self.split_stt_cb.bind("<<ComboboxSelected>>", lambda e: self._run_preview_split())

        ttk.Label(c_row1, text="Cột Họ và tên:", font=FONT_LABEL).pack(side=tk.LEFT, padx=(0, 4))
        self.split_name_cb = ttk.Combobox(c_row1, textvariable=self.split_name_col_var, state="readonly", width=22)
        self.split_name_cb.pack(side=tk.LEFT, padx=(0, 16))
        self.split_name_cb.bind("<<ComboboxSelected>>", lambda e: self._run_preview_split())

        c_row2 = ttk.Frame(col_card)
        c_row2.pack(fill=tk.X, pady=(8, 0))

        ttk.Label(c_row2, text="👉 CHỌN CỘT LÀM CĂN CỨ TÁCH:", font=FONT_LABEL_BOLD, foreground=p.PRIMARY).pack(side=tk.LEFT, padx=(0, 6))
        self.split_target_cb = ttk.Combobox(c_row2, textvariable=self.split_target_col_var, state="readonly", width=38)
        self.split_target_cb.pack(side=tk.LEFT, padx=(0, 10))
        self.split_target_cb.bind("<<ComboboxSelected>>", lambda e: self._run_preview_split())

        ttk.Button(c_row2, text="🔄 Làm mới xem trước", command=self._run_preview_split).pack(side=tk.LEFT)

        # Khung 4: Tùy chọn & Nút bấm thực hiện (GHIM CỐ ĐỊNH Ở ĐÁY - LUÔN LUÔN NHÌN THẤY)
        ctrl_box = ttk.LabelFrame(parent, text="4. Tùy chọn & Bắt đầu tách", padding=8)
        ctrl_box.pack(side=tk.BOTTOM, fill=tk.X, pady=(4, 2))

        # Dòng 1 của ctrl_box: Checkbox tùy chọn + Label tóm tắt
        opt_row = ttk.Frame(ctrl_box)
        opt_row.pack(fill=tk.X, pady=(0, 6))

        ttk.Checkbutton(
            opt_row,
            text="Đánh lại STT liên tục (1, 2, 3...) cho từng file con",
            variable=self.split_renumber_var
        ).pack(side=tk.LEFT)

        self.split_summary_lbl = ttk.Label(opt_row, text="", font=FONT_LABEL_BOLD, foreground=p.SUCCESS)
        self.split_summary_lbl.pack(side=tk.RIGHT)

        # Dòng 2: Thanh tiến trình
        self.split_prog = ttk.Progressbar(ctrl_box, mode="determinate")
        self.split_prog.pack(fill=tk.X, pady=(0, 6))

        # Dòng 3: Nút hành động chính (Action Row)
        act_row = ttk.Frame(ctrl_box)
        act_row.pack(fill=tk.X)

        self.split_status_lbl = ttk.Label(act_row, text="Sẵn sàng tách file.", font=FONT_LABEL_BOLD, foreground=p.PRIMARY)
        self.split_status_lbl.pack(side=tk.LEFT)

        self.btn_open_split_dir = ttk.Button(act_row, text="📂 Mở thư mục kết quả", command=self._open_split_dir, state=tk.DISABLED)
        self.btn_open_split_dir.pack(side=tk.RIGHT, padx=(8, 0))

        # NÚT BẮT ĐẦU TÁCH FILE: To rõ, nổi bật, dùng Action.TButton đảm bảo luôn hiển thị trên mọi OS
        self.btn_run_split = ttk.Button(
            act_row,
            text="⚡ BẮT ĐẦU TÁCH FILE",
            command=self._start_split_thread,
            style="Action.TButton"
        )
        self.btn_run_split.pack(side=tk.RIGHT)

        # Khung 3: Bảng Xem trước (Preview Treeview - Ở giữa co dãn toàn bộ diện tích còn lại)
        prev_card = ttk.LabelFrame(parent, text="3. Xem trước các nhóm sẽ được tách ra", padding=4)
        prev_card.pack(side=tk.TOP, fill=tk.BOTH, expand=True, pady=(2, 2))

        cols = ("stt", "group_val", "row_count", "ratio")
        self.split_tree = ttk.Treeview(prev_card, columns=cols, show="headings", height=5)
        self.split_tree.heading("stt", text="STT")
        self.split_tree.heading("group_val", text="Giá trị trong cột (Tên file con dự kiến)")
        self.split_tree.heading("row_count", text="Số dòng dữ liệu")
        self.split_tree.heading("ratio", text="Tỷ lệ %")

        self.split_tree.column("stt", width=50, anchor=tk.CENTER)
        self.split_tree.column("group_val", width=480, anchor=tk.W)
        self.split_tree.column("row_count", width=130, anchor=tk.CENTER)
        self.split_tree.column("ratio", width=100, anchor=tk.CENTER)

        tree_sb = ttk.Scrollbar(prev_card, orient=tk.VERTICAL, command=self.split_tree.yview)
        self.split_tree.configure(yscrollcommand=tree_sb.set)
        self.split_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        tree_sb.pack(side=tk.RIGHT, fill=tk.Y)

    def _on_choose_split_file(self):
        f = filedialog.askopenfilename(
            title="Chọn tệp Excel nguồn cần tách",
            filetypes=[("Excel Files", "*.xlsx")],
            parent=self
        )
        if not f:
            return
        self.split_file_var.set(f)
        self._analyze_split_file()

    def _analyze_split_file(self):
        f_path = self.split_file_var.get()
        if not f_path or not Path(f_path).exists():
            return

        try:
            p = Path(f_path)
            wb = openpyxl.load_workbook(p, data_only=True, read_only=True)
            sheets = wb.sheetnames
            wb.close()

            self.split_sheet_cb["values"] = sheets
            curr_s = self.split_sheet_var.get()
            if not curr_s or curr_s not in sheets:
                curr_s = sheets[0]
                self.split_sheet_var.set(curr_s)

            wb = openpyxl.load_workbook(p, data_only=True)
            ws = wb[curr_s]
            info = detect_header_and_columns(ws)
            wb.close()

            self.split_header_row_var.set(str(info["header_row_idx"]))
            self.split_columns_data = info["formatted_headers"]

            # Cập nhật Comboboxes
            self.split_stt_cb["values"] = self.split_columns_data
            self.split_name_cb["values"] = self.split_columns_data
            self.split_target_cb["values"] = self.split_columns_data

            if 0 <= info["stt_col_idx"] < len(self.split_columns_data):
                self.split_stt_col_var.set(self.split_columns_data[info["stt_col_idx"]])
            if 0 <= info["name_col_idx"] < len(self.split_columns_data):
                self.split_name_col_var.set(self.split_columns_data[info["name_col_idx"]])

            # Ưu tiên chọn cột Đơn vị nếu có, nếu không chọn cột cuối cùng
            target_idx = None
            for idx, h_text in enumerate(self.split_columns_data):
                h_lower = h_text.lower()
                if "tên đơn vị" in h_lower or "đơn vị" in h_lower:
                    target_idx = idx
                    break
            if target_idx is None:
                for idx, h_text in enumerate(self.split_columns_data):
                    h_lower = h_text.lower()
                    if "tên trường" in h_lower or "trường" in h_lower:
                        target_idx = idx
                        break

            if target_idx is not None and target_idx < len(self.split_columns_data):
                self.split_target_col_var.set(self.split_columns_data[target_idx])
            elif self.split_columns_data:
                self.split_target_col_var.set(self.split_columns_data[-1])

            self.split_file_info_lbl.config(text=f"Phát hiện {len(self.split_columns_data)} cột (Dòng tiêu đề: {info['header_row_idx']})")
            self._run_preview_split()

        except Exception as e:
            self._log(f"❌ Lỗi khi đọc tệp: {e}", "error")

    def _get_selected_col_index(self, var_str: str) -> int:
        """Trích xuất index từ chuỗi dạng '[C] Tên trường' -> index 2."""
        val = var_str.strip()
        for idx, col_text in enumerate(self.split_columns_data):
            if col_text == val:
                return idx
        return 0

    def _run_preview_split(self):
        f_path = self.split_file_var.get()
        if not f_path or not Path(f_path).exists():
            return

        try:
            h_idx = int(self.split_header_row_var.get() or "2")
        except ValueError:
            h_idx = 2

        s_idx = self._get_selected_col_index(self.split_stt_col_var.get())
        n_idx = self._get_selected_col_index(self.split_name_col_var.get())
        sp_idx = self._get_selected_col_index(self.split_target_col_var.get())

        sheet = self.split_sheet_var.get()

        for item in self.split_tree.get_children():
            self.split_tree.delete(item)

        try:
            res = ExcelSplitter.preview_split(
                file_path=f_path,
                sheet_name=sheet,
                split_col_idx=sp_idx,
                header_row_idx=h_idx,
                stt_col_idx=s_idx,
                name_col_idx=n_idx
            )

            tot = res["total_valid_rows"]
            groups = res["groups"]

            for idx, (grp_name, count) in enumerate(groups.items(), start=1):
                ratio = f"{(count / tot * 100):.1f}%" if tot > 0 else "0%"
                self.split_tree.insert("", tk.END, values=(idx, grp_name, f"{count:,} dòng", ratio))

            self.split_summary_lbl.config(
                text=f"📊 Tìm thấy {len(groups)} nhóm, tổng cộng {tot:,} dòng hợp lệ (dừng khi 15 dòng STT trống)"
            )
        except Exception as e:
            self._log(f"⚠️ Lỗi khi xem trước: {e}", "warning")

    def _start_split_thread(self):
        try:
            f_path = self.split_file_var.get().strip()
            if not f_path or not Path(f_path).exists():
                messagebox.showwarning("Chưa chọn tệp", "Vui lòng chọn tệp Excel nguồn cần tách!", parent=self)
                return

            col_str = self.split_target_col_var.get().strip()
            if not col_str:
                messagebox.showwarning("Chưa chọn cột tách", "Vui lòng chọn cột làm căn cứ để tách file!", parent=self)
                return

            p = Path(f_path)
            sp_idx = self._get_selected_col_index(col_str)

            try:
                h_idx = int(self.split_header_row_var.get() or "2")
            except ValueError:
                h_idx = 2

            s_idx = self._get_selected_col_index(self.split_stt_col_var.get())
            n_idx = self._get_selected_col_index(self.split_name_col_var.get())
            renumber = self.split_renumber_var.get()
            sheet = self.split_sheet_var.get().strip()

            # Tạo thư mục đầu ra
            clean_col_name = col_str
            if "]" in clean_col_name:
                clean_col_name = clean_col_name.split("]", 1)[-1].strip()
            safe_folder_name = sanitize_filename(clean_col_name)

            out_dir = p.parent / f"Ket_Qua_Tach_Theo_{safe_folder_name}"
            self.last_split_output_dir = out_dir

            self.btn_run_split.config(state=tk.DISABLED)
            self.btn_open_split_dir.config(state=tk.DISABLED)
            self.split_prog["value"] = 0
            self.split_status_lbl.config(text="⏳ Đang phân tích và tách tệp...", foreground=self.palette.WARNING)
            self._log(f"⚡ Bắt đầu tách tệp: {p.name} theo cột '{col_str}'...", "info")

            def run():
                try:
                    def prog_cb(cur, tot, msg):
                        val = (cur / tot * 100) if tot > 0 else 0
                        self.msg_queue.put(("split_prog", val, f"[{cur}/{tot}] {msg}"))

                    res = ExcelSplitter.split_file(
                        file_path=f_path,
                        split_col_idx=sp_idx,
                        output_dir=out_dir,
                        sheet_name=sheet,
                        header_row_idx=h_idx,
                        stt_col_idx=s_idx,
                        name_col_idx=n_idx,
                        renumber_stt=renumber,
                        progress_callback=prog_cb,
                        log_callback=lambda m, l: self._log(m, l)
                    )

                    self.msg_queue.put(("split_done", res))
                except Exception as e:
                    import traceback
                    self.msg_queue.put(("split_error", str(e), traceback.format_exc()))

            threading.Thread(target=run, daemon=True).start()

        except Exception as ex:
            import traceback
            err_trace = traceback.format_exc()
            self._log(f"❌ Lỗi khởi chạy tiến trình tách tệp: {ex}\n{err_trace}", "error")
            messagebox.showerror("Lỗi khởi chạy", f"Không thể bắt đầu tách file:\n{ex}", parent=self)
            self.btn_run_split.config(state=tk.NORMAL)

    def _open_split_dir(self):
        if self.last_split_output_dir and self.last_split_output_dir.exists():
            open_path_in_os(self.last_split_output_dir)

    # =========================================================================
    # TAB 2: NHẬP / GỘP NHIỀU FILE
    # =========================================================================
    def _build_tab_merge(self, parent):
        p = self.palette

        # Khung 1: Cấu hình tệp đầu ra (Đặt ở trên)
        out_card = ttk.LabelFrame(parent, text="1. Tệp gộp đầu ra & Vị trí lưu", padding=6)
        out_card.pack(side=tk.TOP, fill=tk.X, pady=(0, 4))

        o_row1 = ttk.Frame(out_card)
        o_row1.pack(fill=tk.X)

        ttk.Label(o_row1, text="Tệp gộp đầu ra:", font=FONT_LABEL_BOLD).pack(side=tk.LEFT, padx=(0, 6))
        self.merge_out_entry = ttk.Entry(o_row1, textvariable=self.merge_out_var, font=FONT_LABEL)
        self.merge_out_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 6))
        ttk.Button(o_row1, text="💾 Lưu thành...", command=self._on_choose_merge_output_path).pack(side=tk.LEFT)

        # Khung 3: Tùy chọn & Bắt đầu gộp (BOTTOM, GHIM CỐ ĐỊNH Ở ĐÁY - LUÔN LUÔN NHÌN THẤY)
        m_ctrl_box = ttk.LabelFrame(parent, text="3. Tùy chọn & Bắt đầu gộp", padding=8)
        m_ctrl_box.pack(side=tk.BOTTOM, fill=tk.X, pady=(4, 2))

        m_opt_row = ttk.Frame(m_ctrl_box)
        m_opt_row.pack(fill=tk.X, pady=(0, 6))

        ttk.Checkbutton(
            m_opt_row,
            text="Đánh lại STT (1, 2, 3...)",
            variable=self.merge_renumber_var
        ).pack(side=tk.LEFT, padx=(0, 16))

        ttk.Checkbutton(
            m_opt_row,
            text="Thêm cột 'Tệp nguồn'",
            variable=self.merge_source_col_var
        ).pack(side=tk.LEFT, padx=(0, 16))

        ttk.Checkbutton(
            m_opt_row,
            text="Quét tất cả Sheet có dữ liệu học sinh",
            variable=self.merge_all_sheets_var
        ).pack(side=tk.LEFT)

        # Thanh tiến trình
        self.merge_prog = ttk.Progressbar(m_ctrl_box, mode="determinate")
        self.merge_prog.pack(fill=tk.X, pady=(0, 6))

        # Dòng nút hành động (Action Row)
        m_act_row = ttk.Frame(m_ctrl_box)
        m_act_row.pack(fill=tk.X)

        self.merge_status_lbl = ttk.Label(m_act_row, text="Sẵn sàng gộp file.", font=FONT_LABEL_BOLD, foreground=p.PRIMARY)
        self.merge_status_lbl.pack(side=tk.LEFT)

        self.btn_open_merge_file = ttk.Button(m_act_row, text="📂 Mở tệp gộp", command=self._open_merge_file, state=tk.DISABLED)
        self.btn_open_merge_file.pack(side=tk.RIGHT, padx=(8, 0))

        # NÚT BẮT ĐẦU GỘP FILE: To rõ, nổi bật, dùng Action.TButton đảm bảo luôn hiển thị trên mọi OS
        self.btn_run_merge = ttk.Button(
            m_act_row,
            text="⚡ BẮT ĐẦU GỘP FILE",
            command=self._start_merge_thread,
            style="Action.TButton"
        )
        self.btn_run_merge.pack(side=tk.RIGHT)

        # Khung 2: Danh sách tệp nguồn cần gộp (Ở giữa, chiếm trọn khoảng trống co dãn)
        top_card = ttk.LabelFrame(parent, text="2. Danh sách tệp nguồn cần gộp", padding=6)
        top_card.pack(side=tk.TOP, fill=tk.BOTH, expand=True, pady=(2, 2))

        btn_row = ttk.Frame(top_card)
        btn_row.pack(fill=tk.X, pady=(0, 4))

        ttk.Button(btn_row, text="📂 Chọn nhiều tệp...", command=self._on_choose_merge_files).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(btn_row, text="📁 Chọn cả thư mục...", command=self._on_choose_merge_folder).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(btn_row, text="❌ Xóa danh sách", command=self._clear_merge_list).pack(side=tk.LEFT, padx=(0, 16))

        self.merge_files_count_lbl = ttk.Label(btn_row, text="Chưa chọn tệp nào", font=FONT_LABEL_BOLD)
        self.merge_files_count_lbl.pack(side=tk.LEFT)

        # Bảng danh sách tệp
        m_cols = ("stt", "file_name", "status", "cols", "header_row")
        self.merge_tree = ttk.Treeview(top_card, columns=m_cols, show="headings", height=5)
        self.merge_tree.heading("stt", text="STT")
        self.merge_tree.heading("file_name", text="Tên tệp Excel")
        self.merge_tree.heading("status", text="Độ tương thích tiêu đề")
        self.merge_tree.heading("cols", text="Số cột")
        self.merge_tree.heading("header_row", text="Dòng tiêu đề")

        self.merge_tree.column("stt", width=50, anchor=tk.CENTER)
        self.merge_tree.column("file_name", width=420, anchor=tk.W)
        self.merge_tree.column("status", width=200, anchor=tk.CENTER)
        self.merge_tree.column("cols", width=80, anchor=tk.CENTER)
        self.merge_tree.column("header_row", width=100, anchor=tk.CENTER)

        m_sb = ttk.Scrollbar(top_card, orient=tk.VERTICAL, command=self.merge_tree.yview)
        self.merge_tree.configure(yscrollcommand=m_sb.set)
        self.merge_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        m_sb.pack(side=tk.RIGHT, fill=tk.Y)

    def _on_choose_merge_files(self):
        files = filedialog.askopenfilenames(
            title="Chọn các tệp Excel cần gộp",
            filetypes=[("Excel Files", "*.xlsx")],
            parent=self
        )
        if not files:
            return
        self.merge_files_list = [Path(f) for f in files]
        self._refresh_merge_files_list()

    def _on_choose_merge_folder(self):
        d = filedialog.askdirectory(title="Chọn thư mục chứa các tệp Excel cần gộp", parent=self)
        if not d:
            return
        p = Path(d)
        files = [
            f for f in p.rglob("*.xlsx")
            if not f.name.startswith("~$") and not f.name.startswith("._")
        ]
        if not files:
            messagebox.showinfo("Thông báo", "Không tìm thấy tệp Excel (.xlsx) nào trong thư mục này!", parent=self)
            return
        self.merge_files_list = files
        self._refresh_merge_files_list()

    def _clear_merge_list(self):
        self.merge_files_list = []
        for item in self.merge_tree.get_children():
            self.merge_tree.delete(item)
        self.merge_files_count_lbl.config(text="Chưa chọn tệp nào")

    def _refresh_merge_files_list(self):
        for item in self.merge_tree.get_children():
            self.merge_tree.delete(item)

        if not self.merge_files_list:
            self.merge_files_count_lbl.config(text="Chưa chọn tệp nào")
            return

        self.merge_files_count_lbl.config(text=f"Đã chọn {len(self.merge_files_list)} tệp Excel")

        # Đặt đường dẫn file gộp mặc định
        first_parent = self.merge_files_list[0].parent
        self.merge_out_var.set(str(first_parent / "Ket_Qua_Gop_Du_Lieu.xlsx"))

        # Xác thực cấu trúc tiêu đề
        val_res = ExcelMerger.validate_headers(self.merge_files_list)
        reports = val_res.get("reports", [])

        for idx, rep in enumerate(reports, start=1):
            st = rep.get("status")
            if st == "master":
                status_str = "✅ Tệp mẫu chuẩn (Master)"
            elif st == "match":
                status_str = "✅ Khớp cấu trúc tiêu đề"
            else:
                status_str = "⚠️ Lệch số lượng cột"

            self.merge_tree.insert(
                "",
                tk.END,
                values=(idx, rep.get("file"), status_str, rep.get("col_count"), rep.get("header_row"))
            )

        if not val_res.get("compatible"):
            self._log("⚠️ Cảnh báo: Một số tệp có số lượng cột không khớp với tệp mẫu đầu tiên!", "warning")
        else:
            self._log(f"✅ Tuyệt vời: Toàn bộ {len(self.merge_files_list)} tệp đều có cấu trúc tiêu đề tương thích!", "success")

    def _on_choose_merge_output_path(self):
        f = filedialog.asksaveasfilename(
            title="Lưu tệp Excel gộp thành",
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx")],
            parent=self
        )
        if f:
            self.merge_out_var.set(f)

    def _start_merge_thread(self):
        try:
            if not self.merge_files_list:
                messagebox.showwarning("Chưa chọn tệp", "Vui lòng chọn các tệp Excel cần gộp!", parent=self)
                return

            out_path = self.merge_out_var.get().strip()
            if not out_path:
                messagebox.showwarning("Chưa chọn nơi lưu", "Vui lòng chọn đường dẫn tệp gộp đầu ra!", parent=self)
                return

            self.btn_run_merge.config(state=tk.DISABLED)
            self.btn_open_merge_file.config(state=tk.DISABLED)
            self.merge_prog["value"] = 0
            self.merge_status_lbl.config(text="⏳ Đang gộp các tệp Excel...", foreground=self.palette.WARNING)
            self._log(f"⚡ Bắt đầu gộp {len(self.merge_files_list)} tệp vào: {Path(out_path).name}...", "info")

            renumber = self.merge_renumber_var.get()
            add_source = self.merge_source_col_var.get()
            merge_all = self.merge_all_sheets_var.get()

            def run():
                try:
                    def prog_cb(cur, tot, msg):
                        val = (cur / tot * 100) if tot > 0 else 0
                        self.msg_queue.put(("merge_prog", val, f"[{cur}/{tot}] {msg}"))

                    res = ExcelMerger.merge_files(
                        file_paths=self.merge_files_list,
                        output_file_path=out_path,
                        renumber_stt=renumber,
                        add_source_col=add_source,
                        merge_all_sheets=merge_all,
                        progress_callback=prog_cb,
                        log_callback=lambda m, l: self._log(m, l)
                    )

                    self.msg_queue.put(("merge_done", res, Path(out_path)))
                except Exception as e:
                    import traceback
                    self.msg_queue.put(("merge_error", str(e), traceback.format_exc()))

            threading.Thread(target=run, daemon=True).start()

        except Exception as ex:
            import traceback
            err_trace = traceback.format_exc()
            self._log(f"❌ Lỗi khởi chạy tiến trình gộp tệp: {ex}\n{err_trace}", "error")
            messagebox.showerror("Lỗi khởi chạy", f"Không thể bắt đầu gộp file:\n{ex}", parent=self)
            self.btn_run_merge.config(state=tk.NORMAL)

    def _open_merge_file(self):
        if self.last_merge_output_file and self.last_merge_output_file.exists():
            open_path_in_os(self.last_merge_output_file)

    # =========================================================================
    # TAB 3: GHÉP & CHUẨN HÓA THÔNG MINH THEO MẪU CHUẨN 28 CỘT
    # =========================================================================
    def _build_tab_smart(self, parent):
        p = self.palette

        # Khung 1: Tệp kết quả tổng hợp đầu ra
        o_card = ttk.LabelFrame(parent, text="1. Tệp kết quả tổng hợp (Mẫu chuẩn Mau_DS_Cap_Bang_THCS.xlsx - 28 cột)", padding=8)
        o_card.pack(fill=tk.X, pady=(0, 4))

        o_row = ttk.Frame(o_card)
        o_row.pack(fill=tk.X)
        ttk.Label(o_row, text="Đường dẫn tệp đích:", font=FONT_LABEL_BOLD).pack(side=tk.LEFT, padx=(0, 6))
        self.smart_out_entry = ttk.Entry(o_row, textvariable=self.smart_out_var, font=FONT_LABEL)
        self.smart_out_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 6))
        ttk.Button(o_row, text="💾 Lưu thành...", command=self._on_choose_smart_output_path).pack(side=tk.LEFT)

        # Khung 3: Tùy chọn & Nút bấm thực hiện (GHIM CỐ ĐỊNH Ở ĐÁY - LUÔN LUÔN NHÌN THẤY)
        s_ctrl_box = ttk.LabelFrame(parent, text="3. Tùy chọn & Bắt đầu ghép chuẩn hóa", padding=8)
        s_ctrl_box.pack(side=tk.BOTTOM, fill=tk.X, pady=(4, 2))

        s_opt_row = ttk.Frame(s_ctrl_box)
        s_opt_row.pack(fill=tk.X, pady=(0, 6))

        ttk.Checkbutton(
            s_opt_row,
            text="Tự động đánh lại STT liên tục (1, 2, 3...)",
            variable=self.smart_renumber_var
        ).pack(side=tk.LEFT, padx=(0, 16))

        ttk.Checkbutton(
            s_opt_row,
            text="Tự động điền giá trị mặc định cho cột để trống (Tên văn bằng, Quốc tịch...)",
            variable=self.smart_fill_defaults_var
        ).pack(side=tk.LEFT, padx=(0, 16))

        ttk.Checkbutton(
            s_opt_row,
            text="Quét tất cả Sheet có dữ liệu học sinh",
            variable=self.smart_all_sheets_var
        ).pack(side=tk.LEFT)

        # Thanh tiến trình
        self.smart_prog = ttk.Progressbar(s_ctrl_box, mode="determinate")
        self.smart_prog.pack(fill=tk.X, pady=(0, 6))

        # Dòng nút hành động
        s_act_row = ttk.Frame(s_ctrl_box)
        s_act_row.pack(fill=tk.X)

        self.smart_status_lbl = ttk.Label(s_act_row, text="Sẵn sàng ghép và chuẩn hóa.", font=FONT_LABEL_BOLD, foreground=p.PRIMARY)
        self.smart_status_lbl.pack(side=tk.LEFT)

        self.btn_open_smart_file = ttk.Button(s_act_row, text="📂 Mở tệp kết quả", command=self._open_smart_file, state=tk.DISABLED)
        self.btn_open_smart_file.pack(side=tk.RIGHT, padx=(8, 0))

        self.btn_cancel_smart_merge = ttk.Button(
            s_act_row,
            text="⏹️ Dừng",
            command=self._on_cancel_smart_merge,
            state=tk.DISABLED
        )
        self.btn_cancel_smart_merge.pack(side=tk.RIGHT, padx=(8, 0))

        self.btn_run_smart_merge = ttk.Button(
            s_act_row,
            text="⚡ BẮT ĐẦU GHÉP & CHUẨN HÓA",
            command=self._start_smart_merge_thread,
            style="Action.TButton"
        )
        self.btn_run_smart_merge.pack(side=tk.RIGHT, padx=(8, 0))

        ttk.Button(s_act_row, text="⚙️ Quản lý Profiles đã lưu", command=self._open_manage_profiles_dialog).pack(side=tk.RIGHT)

        # Khung 2: Danh sách tệp nguồn cần ghép (Ở giữa, chiếm trọn khoảng trống co dãn)
        s_top_card = ttk.LabelFrame(parent, text="2. Danh sách tệp nguồn cần ghép (có thể khác hoặc cùng cấu trúc cột)", padding=6)
        s_top_card.pack(side=tk.TOP, fill=tk.BOTH, expand=True, pady=(2, 2))

        btn_row = ttk.Frame(s_top_card)
        btn_row.pack(fill=tk.X, pady=(0, 4))

        ttk.Button(btn_row, text="📂 Chọn nhiều tệp...", command=self._on_choose_smart_files).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(btn_row, text="📁 Chọn cả thư mục...", command=self._on_choose_smart_folder).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(btn_row, text="❌ Xóa danh sách", command=self._clear_smart_list).pack(side=tk.LEFT, padx=(0, 16))

        self.smart_files_count_lbl = ttk.Label(btn_row, text="Chưa chọn tệp nào", font=FONT_LABEL_BOLD)
        self.smart_files_count_lbl.pack(side=tk.LEFT)

        # Bảng danh sách tệp nguồn
        sm_cols = ("stt", "file_name", "size", "path")
        self.smart_tree = ttk.Treeview(s_top_card, columns=sm_cols, show="headings", height=5)
        self.smart_tree.heading("stt", text="STT")
        self.smart_tree.heading("file_name", text="Tên tệp Excel")
        self.smart_tree.heading("size", text="Kích thước")
        self.smart_tree.heading("path", text="Đường dẫn tệp")

        self.smart_tree.column("stt", width=50, anchor=tk.CENTER)
        self.smart_tree.column("file_name", width=360, anchor=tk.W)
        self.smart_tree.column("size", width=110, anchor=tk.CENTER)
        self.smart_tree.column("path", width=380, anchor=tk.W)

        sm_sb = ttk.Scrollbar(s_top_card, orient=tk.VERTICAL, command=self.smart_tree.yview)
        self.smart_tree.configure(yscrollcommand=sm_sb.set)
        self.smart_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sm_sb.pack(side=tk.RIGHT, fill=tk.Y)

    def _on_choose_smart_files(self):
        files = filedialog.askopenfilenames(
            title="Chọn các tệp Excel nguồn cần ghép & chuẩn hóa",
            filetypes=[("Excel Files", "*.xlsx")],
            parent=self
        )
        if not files:
            return
        self.smart_files_list = [Path(f) for f in files]
        self._refresh_smart_files_list()

    def _on_choose_smart_folder(self):
        d = filedialog.askdirectory(title="Chọn thư mục chứa các tệp Excel cần ghép & chuẩn hóa", parent=self)
        if not d:
            return
        p = Path(d)
        files = [
            f for f in p.rglob("*.xlsx")
            if not f.name.startswith("~$") and not f.name.startswith("._")
        ]
        if not files:
            messagebox.showinfo("Thông báo", "Không tìm thấy tệp Excel (.xlsx) nào trong thư mục này!", parent=self)
            return
        self.smart_files_list = files
        self._refresh_smart_files_list()

    def _clear_smart_list(self):
        self.smart_files_list = []
        for item in self.smart_tree.get_children():
            self.smart_tree.delete(item)
        self.smart_files_count_lbl.config(text="Chưa chọn tệp nào")

    def _refresh_smart_files_list(self):
        for item in self.smart_tree.get_children():
            self.smart_tree.delete(item)

        if not self.smart_files_list:
            self.smart_files_count_lbl.config(text="Chưa chọn tệp nào")
            return

        self.smart_files_count_lbl.config(text=f"Đã chọn {len(self.smart_files_list)} tệp Excel")

        # Đặt đường dẫn file đích mặc định
        first_parent = self.smart_files_list[0].parent
        self.smart_out_var.set(str(first_parent / "Ket_Qua_Ghep_Chuan_Hoa_Mau_28_Cot.xlsx"))

        for idx, f_p in enumerate(self.smart_files_list, start=1):
            try:
                sz_bytes = f_p.stat().st_size
                sz_str = f"{sz_bytes / 1024:.1f} KB" if sz_bytes < 1024 * 1024 else f"{sz_bytes / (1024 * 1024):.1f} MB"
            except Exception:
                sz_str = "---"

            self.smart_tree.insert(
                "",
                tk.END,
                values=(idx, f_p.name, sz_str, str(f_p))
            )

    def _on_choose_smart_output_path(self):
        f = filedialog.asksaveasfilename(
            title="Lưu tệp tổng hợp 28 cột chuẩn thành",
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx")],
            parent=self
        )
        if f:
            self.smart_out_var.set(f)

    def _start_smart_merge_thread(self):
        try:
            if not self.smart_files_list:
                messagebox.showwarning("Chưa chọn tệp", "Vui lòng chọn các tệp Excel nguồn cần ghép & chuẩn hóa!", parent=self)
                return

            out_path = self.smart_out_var.get().strip()
            if not out_path:
                messagebox.showwarning("Chưa chọn nơi lưu", "Vui lòng chọn đường dẫn tệp kết quả đầu ra!", parent=self)
                return

            self.smart_cancel_requested = False
            self.btn_run_smart_merge.config(state=tk.DISABLED)
            self.btn_cancel_smart_merge.config(state=tk.NORMAL)
            self.btn_open_smart_file.config(state=tk.DISABLED)
            self.smart_prog["value"] = 0
            self.smart_status_lbl.config(text="⏳ Đang quét và chuẩn hóa các tệp...", foreground=self.palette.WARNING)
            self._log(f"⚡ Bắt đầu tiến trình Ghép & Chuẩn hóa {len(self.smart_files_list)} tệp vào: {Path(out_path).name}...", "info")

            renumber = self.smart_renumber_var.get()
            fill_defs = self.smart_fill_defaults_var.get()
            merge_all = self.smart_all_sheets_var.get()

            def run():
                try:
                    def prog_cb(cur, tot, msg):
                        val = (cur / tot * 100) if tot > 0 else 0
                        self.msg_queue.put(("smart_prog", val, f"[{cur}/{tot}] {msg}"))

                    def mapping_callback_sync(payload):
                        resp_event = threading.Event()
                        resp_holder = {"result": ("skip", None, None)}
                        self.msg_queue.put(("ask_mapping", payload, resp_event, resp_holder))
                        resp_event.wait()
                        return resp_holder["result"]

                    res = self.smart_merger.process_merge_and_standardize(
                        source_files=self.smart_files_list,
                        output_file_path=out_path,
                        mapping_callback=mapping_callback_sync,
                        renumber_stt=renumber,
                        fill_defaults=fill_defs,
                        merge_all_sheets=merge_all,
                        progress_callback=prog_cb,
                        log_callback=lambda m, l: self._log(m, l),
                        check_cancel=lambda: self.smart_cancel_requested
                    )

                    self.msg_queue.put(("smart_done", res, Path(out_path)))
                except Exception as e:
                    import traceback
                    self.msg_queue.put(("smart_error", str(e), traceback.format_exc()))

            threading.Thread(target=run, daemon=True).start()

        except Exception as ex:
            import traceback
            err_trace = traceback.format_exc()
            self._log(f"❌ Lỗi khởi chạy tiến trình ghép & chuẩn hóa: {ex}\n{err_trace}", "error")
            messagebox.showerror("Lỗi khởi chạy", f"Không thể bắt đầu ghép và chuẩn hóa file:\n{ex}", parent=self)
            self.btn_run_smart_merge.config(state=tk.NORMAL)
            self.btn_cancel_smart_merge.config(state=tk.DISABLED)

    def _on_cancel_smart_merge(self):
        """Yêu cầu dừng toàn bộ tiến trình ghép file."""
        if messagebox.askyesno("Xác nhận dừng", "Bạn có chắc chắn muốn DỪNG TOÀN BỘ quá trình ghép không?", parent=self):
            self.smart_cancel_requested = True
            self.btn_cancel_smart_merge.config(state=tk.DISABLED)
            self.smart_status_lbl.config(text="⏳ Đang dừng toàn bộ quá trình...", foreground=self.palette.DANGER)
            self._log("⏹ Người dùng đã gửi lệnh dừng toàn bộ tiến trình...", "warning")

    def _open_smart_file(self):
        if self.last_smart_output_file and self.last_smart_output_file.exists():
            open_path_in_os(self.last_smart_output_file)

    def _open_manage_profiles_dialog(self):
        """Mở hộp thoại quản lý các Profile ánh xạ cột đã lưu."""
        dlg = ManageProfilesDialog(
            parent=self.winfo_toplevel(),
            profile_mgr=self.smart_merger.profile_mgr,
            is_dark=self.is_dark
        )

    # =========================================================================
    # TAB 4: CHUYỂN ĐỔI ĐỊNH DẠNG TỆP CŨ (.doc ➔ .docx, .xls ➔ .xlsx)
    # =========================================================================
    def _build_tab_convert(self, parent):
        p = self.palette

        # Khung 1: Thư mục lưu tệp kết quả (Top row)
        o_card = ttk.LabelFrame(parent, text="1. Thư mục lưu kết quả tệp mới (.docx, .xlsx)", padding=8)
        o_card.pack(fill=tk.X, pady=(0, 4))

        r_row = ttk.Frame(o_card)
        r_row.pack(fill=tk.X)

        ttk.Radiobutton(
            r_row,
            text="Lưu tệp mới ngay cùng thư mục với tệp gốc",
            variable=self.convert_use_custom_dir,
            value=False
        ).pack(side=tk.LEFT, padx=(0, 20))

        ttk.Radiobutton(
            r_row,
            text="Lưu vào thư mục riêng:",
            variable=self.convert_use_custom_dir,
            value=True
        ).pack(side=tk.LEFT, padx=(0, 6))

        self.convert_dir_entry = ttk.Entry(r_row, textvariable=self.convert_custom_dir_var, font=FONT_LABEL, width=35)
        self.convert_dir_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 6))
        ttk.Button(r_row, text="📁 Chọn thư mục...", command=self._on_choose_convert_custom_dir).pack(side=tk.LEFT)

        # Khung 3: Tùy chọn & Bắt đầu chuyển đổi (GHIM CỐ ĐỊNH Ở ĐÁY)
        c_ctrl_box = ttk.LabelFrame(parent, text="3. Tùy chọn & Thực thi", padding=8)
        c_ctrl_box.pack(side=tk.BOTTOM, fill=tk.X, pady=(4, 2))

        c_opt_row = ttk.Frame(c_ctrl_box)
        c_opt_row.pack(fill=tk.X, pady=(0, 6))

        ttk.Checkbutton(
            c_opt_row,
            text="Bỏ qua nếu tệp đích (.docx, .xlsx) đã tồn tại",
            variable=self.convert_skip_existing_var
        ).pack(side=tk.LEFT, padx=(0, 20))

        ttk.Checkbutton(
            c_opt_row,
            text="⚠️ Xóa tệp gốc cũ (.doc, .xls) sau khi chuyển đổi thành công",
            variable=self.convert_delete_source_var
        ).pack(side=tk.LEFT)

        # Thanh tiến trình
        self.convert_prog = ttk.Progressbar(c_ctrl_box, mode="determinate")
        self.convert_prog.pack(fill=tk.X, pady=(0, 6))

        # Dòng nút hành động
        c_act_row = ttk.Frame(c_ctrl_box)
        c_act_row.pack(fill=tk.X)

        self.convert_status_lbl = ttk.Label(c_act_row, text="Sẵn sàng chuyển đổi định dạng tệp cũ.", font=FONT_LABEL_BOLD, foreground=p.PRIMARY)
        self.convert_status_lbl.pack(side=tk.LEFT)

        self.btn_open_convert_dir = ttk.Button(c_act_row, text="📂 Mở thư mục kết quả", command=self._open_convert_dir, state=tk.DISABLED)
        self.btn_open_convert_dir.pack(side=tk.RIGHT, padx=(8, 0))

        self.btn_run_convert = ttk.Button(
            c_act_row,
            text="⚡ BẮT ĐẦU CHUYỂN ĐỔI",
            command=self._start_convert_thread,
            style="Action.TButton"
        )
        self.btn_run_convert.pack(side=tk.RIGHT)

        # Khung 2: Danh sách tệp nguồn cần chuyển đổi (Ở giữa, dãn nở)
        c_top_card = ttk.LabelFrame(parent, text="2. Danh sách tệp Word (.doc) và Excel (.xls) cần chuyển đổi", padding=6)
        c_top_card.pack(side=tk.TOP, fill=tk.BOTH, expand=True, pady=(2, 2))

        btn_row = ttk.Frame(c_top_card)
        btn_row.pack(fill=tk.X, pady=(0, 4))

        ttk.Button(btn_row, text="📂 Chọn nhiều tệp...", command=self._on_choose_convert_files).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(btn_row, text="📁 Quét cả thư mục...", command=self._on_choose_convert_folder).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(btn_row, text="❌ Xóa danh sách", command=self._clear_convert_list).pack(side=tk.LEFT, padx=(0, 16))

        self.convert_count_lbl = ttk.Label(btn_row, text="Chưa chọn tệp nào", font=FONT_LABEL_BOLD)
        self.convert_count_lbl.pack(side=tk.LEFT)

        # Bảng danh sách tệp
        cv_cols = ("stt", "file_name", "type", "size", "status", "path")
        self.convert_tree = ttk.Treeview(c_top_card, columns=cv_cols, show="headings", height=5)
        self.convert_tree.heading("stt", text="STT")
        self.convert_tree.heading("file_name", text="Tên tệp nguồn")
        self.convert_tree.heading("type", text="Loại tệp")
        self.convert_tree.heading("size", text="Kích thước")
        self.convert_tree.heading("status", text="Trạng thái")
        self.convert_tree.heading("path", text="Đường dẫn tệp")

        self.convert_tree.column("stt", width=45, anchor=tk.CENTER)
        self.convert_tree.column("file_name", width=260, anchor=tk.W)
        self.convert_tree.column("type", width=120, anchor=tk.CENTER)
        self.convert_tree.column("size", width=95, anchor=tk.CENTER)
        self.convert_tree.column("status", width=140, anchor=tk.CENTER)
        self.convert_tree.column("path", width=300, anchor=tk.W)

        cv_sb = ttk.Scrollbar(c_top_card, orient=tk.VERTICAL, command=self.convert_tree.yview)
        self.convert_tree.configure(yscrollcommand=cv_sb.set)
        self.convert_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        cv_sb.pack(side=tk.RIGHT, fill=tk.Y)

    def _on_choose_convert_files(self):
        files = filedialog.askopenfilenames(
            title="Chọn các tệp .doc hoặc .xls cần chuyển đổi",
            filetypes=[
                ("Tệp cũ (*.doc, *.xls)", "*.doc *.xls"),
                ("Word 97-2003 (*.doc)", "*.doc"),
                ("Excel 97-2003 (*.xls)", "*.xls")
            ],
            parent=self
        )
        if not files:
            return
        scan_res = FileFormatConverter.scan_files([Path(f) for f in files])
        self.convert_files_list = [item["path"] for item in scan_res]
        self._refresh_convert_files_list(scan_res)

    def _on_choose_convert_folder(self):
        d = filedialog.askdirectory(title="Chọn thư mục chứa các tệp .doc và .xls", parent=self)
        if not d:
            return
        scan_res = FileFormatConverter.scan_files([Path(d)], recursive=True)
        if not scan_res:
            messagebox.showinfo("Thông báo", "Không tìm thấy tệp Word (.doc) hoặc Excel (.xls) nào trong thư mục này!", parent=self)
            return
        self.convert_files_list = [item["path"] for item in scan_res]
        self._refresh_convert_files_list(scan_res)

    def _clear_convert_list(self):
        self.convert_files_list = []
        for item in self.convert_tree.get_children():
            self.convert_tree.delete(item)
        self.convert_count_lbl.config(text="Chưa chọn tệp nào")

    def _refresh_convert_files_list(self, scan_data: list[dict] | None = None):
        for item in self.convert_tree.get_children():
            self.convert_tree.delete(item)

        if not self.convert_files_list:
            self.convert_count_lbl.config(text="Chưa chọn tệp nào")
            return

        items = scan_data or FileFormatConverter.scan_files(self.convert_files_list)
        doc_count = sum(1 for it in items if it["ext"] == ".doc")
        xls_count = sum(1 for it in items if it["ext"] == ".xls")

        self.convert_count_lbl.config(
            text=f"Đã chọn {len(items)} tệp (Word .doc: {doc_count} | Excel .xls: {xls_count})"
        )

        for idx, it in enumerate(items, start=1):
            self.convert_tree.insert(
                "",
                tk.END,
                values=(
                    idx,
                    it["name"],
                    it["type_label"],
                    it["size_str"],
                    it["status"],
                    str(it["path"])
                )
            )

    def _on_choose_convert_custom_dir(self):
        d = filedialog.askdirectory(title="Chọn thư mục lưu các tệp mới (.docx, .xlsx)", parent=self)
        if d:
            self.convert_custom_dir_var.set(d)
            self.convert_use_custom_dir.set(True)

    def _start_convert_thread(self):
        try:
            if not self.convert_files_list:
                messagebox.showwarning("Chưa chọn tệp", "Vui lòng chọn các tệp .doc hoặc .xls cần chuyển đổi!", parent=self)
                return

            use_custom = self.convert_use_custom_dir.get()
            custom_dir_str = self.convert_custom_dir_var.get().strip()

            out_dir: Path | None = None
            if use_custom:
                if not custom_dir_str:
                    messagebox.showwarning("Chưa chọn thư mục", "Vui lòng chọn thư mục riêng để lưu kết quả!", parent=self)
                    return
                out_dir = Path(custom_dir_str)
                out_dir.mkdir(parents=True, exist_ok=True)

            self.btn_run_convert.config(state=tk.DISABLED)
            self.btn_open_convert_dir.config(state=tk.DISABLED)
            self.convert_prog["value"] = 0
            self.convert_status_lbl.config(text="⏳ Đang chuyển đổi định dạng các tệp...", foreground=self.palette.WARNING)
            self._log(f"⚡ Bắt đầu chuyển đổi {len(self.convert_files_list)} tệp .doc / .xls...", "info")

            delete_src = self.convert_delete_source_var.get()
            skip_ex = self.convert_skip_existing_var.get()

            def run():
                try:
                    def prog_cb(cur, tot, msg):
                        val = (cur / tot * 100) if tot > 0 else 0
                        self.msg_queue.put(("convert_prog", val, msg))

                    def file_st_cb(f_path, status):
                        self.msg_queue.put(("convert_file_status", f_path, status))

                    res = FileFormatConverter.batch_convert(
                        files=self.convert_files_list,
                        output_dir=out_dir,
                        delete_source=delete_src,
                        skip_existing=skip_ex,
                        progress_callback=prog_cb,
                        file_status_callback=file_st_cb,
                        log_callback=lambda m, l: self._log(m, l)
                    )

                    target_dir = out_dir or (self.convert_files_list[0].parent if self.convert_files_list else None)
                    self.msg_queue.put(("convert_done", res, target_dir))
                except Exception as e:
                    import traceback
                    self.msg_queue.put(("convert_error", str(e), traceback.format_exc()))

            threading.Thread(target=run, daemon=True).start()

        except Exception as ex:
            import traceback
            err_trace = traceback.format_exc()
            self._log(f"❌ Lỗi khởi chạy chuyển đổi tệp: {ex}\n{err_trace}", "error")
            messagebox.showerror("Lỗi khởi chạy", f"Không thể bắt đầu chuyển đổi file:\n{ex}", parent=self)
            self.btn_run_convert.config(state=tk.NORMAL)

    def _open_convert_dir(self):
        if self.last_convert_dir and self.last_convert_dir.exists():
            open_path_in_os(self.last_convert_dir)

    # =========================================================================
    # LOG & QUEUE POLLER
    # =========================================================================
    def _log(self, msg: str, level: str = "info"):
        self.msg_queue.put(("log", msg, level))

    def _start_queue_poller(self):
        while not self.msg_queue.empty():
            item = self.msg_queue.get_nowait()
            msg_type = item[0]
            if msg_type == "log":
                _, msg, level = item
                self.log_text.insert(tk.END, msg + "\n", level)
                self.log_text.see(tk.END)
            elif msg_type == "split_prog":
                val, msg = item[1], item[2]
                self.split_prog["value"] = val
                self._log(f"  {msg}", "info")
            elif msg_type == "split_done":
                res = item[1]
                self.split_prog["value"] = 100
                self.btn_run_split.config(state=tk.NORMAL)
                self.btn_open_split_dir.config(state=tk.NORMAL)
                total_groups = res.get("total_groups", len(res.get("files", [])))
                total_rows = res.get("total_rows", 0)

                if res.get("status") == "empty" or total_rows == 0:
                    self.split_status_lbl.config(
                        text="⚠️ Không tìm thấy dòng dữ liệu hợp lệ nào!",
                        foreground=self.palette.WARNING
                    )
                    messagebox.showwarning(
                        "Không có dữ liệu hợp lệ",
                        "Không tìm thấy dòng dữ liệu nào thỏa mãn điều kiện hợp lệ (phải có STT và Họ tên)!",
                        parent=self
                    )
                else:
                    self.split_status_lbl.config(
                        text=f"✅ Đã tách xong {total_groups} file ({total_rows:,} dòng)!",
                        foreground=self.palette.SUCCESS
                    )
                    folder_name = self.last_split_output_dir.name if self.last_split_output_dir else ""
                    messagebox.showinfo(
                        "Hoàn tất Tách tệp",
                        f"Đã tách thành công {total_groups} tệp Excel với {total_rows:,} dòng dữ liệu hợp lệ!\n\nThư mục: {folder_name}",
                        parent=self
                    )
            elif msg_type == "split_error":
                err, trace = item[1], item[2]
                self.btn_run_split.config(state=tk.NORMAL)
                self.btn_open_split_dir.config(state=tk.NORMAL)
                self.split_status_lbl.config(text="❌ Lỗi khi tách tệp!", foreground=self.palette.DANGER)
                self._log(f"❌ Lỗi khi tách tệp: {err}\n{trace}", "error")
                messagebox.showerror("Lỗi tách tệp", f"Đã xảy ra lỗi khi tách tệp:\n{err}", parent=self)
            elif msg_type == "merge_prog":
                val, msg = item[1], item[2]
                self.merge_prog["value"] = val
                self._log(f"  {msg}", "info")
            elif msg_type == "merge_done":
                res, out_p = item[1], item[2]
                self.merge_prog["value"] = 100
                self.btn_run_merge.config(state=tk.NORMAL)
                self.btn_open_merge_file.config(state=tk.NORMAL)
                self.last_merge_output_file = out_p
                total_files = res.get("total_files", 0)
                total_rows = res.get("total_rows", 0)

                if res.get("status") == "empty" or total_rows == 0:
                    self.merge_status_lbl.config(
                        text="⚠️ Không tìm thấy dòng dữ liệu hợp lệ nào để gộp!",
                        foreground=self.palette.WARNING
                    )
                    messagebox.showwarning(
                        "Không có dữ liệu hợp lệ",
                        "Không tìm thấy dòng dữ liệu nào thỏa mãn điều kiện hợp lệ trong các tệp nguồn!",
                        parent=self
                    )
                else:
                    self.merge_status_lbl.config(
                        text=f"✅ Đã gộp thành công {total_files} tệp ({total_rows:,} dòng)!",
                        foreground=self.palette.SUCCESS
                    )
                    messagebox.showinfo(
                        "Hoàn tất Gộp tệp",
                        f"Đã gộp thành công {total_files} tệp với tổng cộng {total_rows:,} dòng dữ liệu hợp lệ!\n\nTệp: {out_p.name}",
                        parent=self
                    )
            elif msg_type == "merge_error":
                err, trace = item[1], item[2]
                self.btn_run_merge.config(state=tk.NORMAL)
                self.btn_open_merge_file.config(state=tk.NORMAL)
                self.merge_status_lbl.config(text="❌ Lỗi khi gộp tệp!", foreground=self.palette.DANGER)
                self._log(f"❌ Lỗi khi gộp tệp: {err}\n{trace}", "error")
                messagebox.showerror("Lỗi gộp tệp", f"Đã xảy ra lỗi khi gộp tệp:\n{err}", parent=self)
            elif msg_type == "ask_mapping":
                payload, resp_event, resp_holder = item[1], item[2], item[3]
                try:
                    dlg = ColumnMappingDialog(
                        parent=self.winfo_toplevel(),
                        file_name=payload["file_name"],
                        sheet_name=payload["sheet_name"],
                        raw_headers=payload["raw_headers"],
                        formatted_headers=payload.get("formatted_headers"),
                        suggested_mapping=payload.get("suggested_mapping"),
                        suggested_inverted_mapping=payload.get("suggested_inverted_mapping"),
                        profile_mgr=self.smart_merger.profile_mgr,
                        is_dark=self.is_dark,
                        suggested_defaults=payload.get("suggested_defaults"),
                        matched_profile_name=payload.get("matched_profile_name"),
                        sample_values=payload.get("sample_values"),
                        extra_cols_with_data=payload.get("extra_cols_with_data")
                    )
                    dlg.wait_window()
                    resp_holder["result"] = dlg.result
                except Exception as ex:
                    self._log(f"❌ Lỗi hiển thị hộp thoại ánh xạ: {ex}", "error")
                    resp_holder["result"] = ("skip", None, None)
                finally:
                    resp_event.set()
            elif msg_type == "smart_prog":
                val, msg = item[1], item[2]
                self.smart_prog["value"] = val
                self._log(f"  {msg}", "info")
            elif msg_type == "smart_done":
                res, out_p = item[1], item[2]
                self.btn_run_smart_merge.config(state=tk.NORMAL)
                self.btn_cancel_smart_merge.config(state=tk.DISABLED)
                total_rows = res.get("total_rows", 0)
                sheets_p = res.get("sheets_processed", 0)

                if res.get("status") == "cancelled":
                    self.smart_status_lbl.config(
                        text=f"⏹️ Đã dừng ghép ({total_rows:,} dòng, {sheets_p} sheet)!",
                        foreground=self.palette.WARNING
                    )
                    messagebox.showinfo("Đã dừng tiến trình", f"Tiến trình ghép file đã dừng theo yêu cầu của bạn.\nĐã xử lý: {sheets_p} sheet ({total_rows:,} dòng).", parent=self)
                else:
                    self.smart_prog["value"] = 100
                    self.btn_open_smart_file.config(state=tk.NORMAL)
                    self.last_smart_output_file = out_p
                    self.smart_status_lbl.config(
                        text=f"✅ Đã ghép xong ({total_rows:,} dòng, {sheets_p} sheet)!",
                        foreground=self.palette.SUCCESS
                    )
                    messagebox.showinfo(
                        "Hoàn tất Ghép & Chuẩn hóa",
                        f"Đã ghép thành công {total_rows:,} dòng dữ liệu học sinh từ {sheets_p} sheet vào mẫu chuẩn 28 cột!\n\nTệp: {out_p.name}",
                        parent=self
                    )
            elif msg_type == "smart_error":
                err, trace = item[1], item[2]
                self.btn_run_smart_merge.config(state=tk.NORMAL)
                self.btn_open_smart_file.config(state=tk.NORMAL)
                self.smart_status_lbl.config(text="❌ Lỗi khi ghép & chuẩn hóa!", foreground=self.palette.DANGER)
                self._log(f"❌ Lỗi khi ghép & chuẩn hóa: {err}\n{trace}", "error")
                messagebox.showerror("Lỗi ghép file", f"Đã xảy ra lỗi khi ghép và chuẩn hóa file:\n{err}", parent=self)
            elif msg_type == "convert_prog":
                val, msg = item[1], item[2]
                self.convert_prog["value"] = val
                self.convert_status_lbl.config(text=f"⏳ {msg}...", foreground=self.palette.WARNING)
                self._log(f"  {msg}", "info")
            elif msg_type == "convert_file_status":
                f_p, st = item[1], item[2]
                for child in self.convert_tree.get_children():
                    vals = list(self.convert_tree.item(child, "values"))
                    if str(f_p) in vals:
                        if "Thành công" in st:
                            icon_st = f"✅ {st}"
                        elif "Bỏ qua" in st:
                            icon_st = f"⏭️ {st}"
                        elif "Đang" in st:
                            icon_st = f"⏳ {st}"
                        else:
                            icon_st = f"❌ {st}"
                        vals[4] = icon_st
                        self.convert_tree.item(child, values=vals)
            elif msg_type == "convert_done":
                res, target_dir = item[1], item[2]
                self.convert_prog["value"] = 100
                self.btn_run_convert.config(state=tk.NORMAL)
                self.btn_open_convert_dir.config(state=tk.NORMAL)
                self.last_convert_dir = target_dir
                succ = res.get("success", 0)
                tot = res.get("total", 0)
                skip = res.get("skipped", 0)
                fail = res.get("failed", 0)

                # Cập nhật trạng thái từng dòng trên Treeview
                for r_item in res.get("results", []):
                    f_p = r_item.get("path")
                    st = r_item.get("status")
                    for child in self.convert_tree.get_children():
                        vals = list(self.convert_tree.item(child, "values"))
                        if str(f_p) in vals:
                            vals[4] = f"✅ {st}" if "Thành công" in st else (f"⏭️ {st}" if "Bỏ qua" in st else f"❌ {st}")
                            self.convert_tree.item(child, values=vals)

                self.convert_status_lbl.config(
                    text=f"✅ Đã chuyển xong {succ}/{tot} tệp ({skip} bỏ qua, {fail} lỗi)!",
                    foreground=self.palette.SUCCESS if fail == 0 else self.palette.WARNING
                )
                messagebox.showinfo(
                    "Hoàn tất Chuyển đổi",
                    f"Đã hoàn thành chuyển đổi định dạng tệp!\n\n"
                    f"• Thành công: {succ} tệp\n"
                    f"• Bỏ qua (đã có): {skip} tệp\n"
                    f"• Thất bại: {fail} tệp\n\n"
                    f"Thư mục: {target_dir.name if target_dir else 'Thư mục gốc của tệp'}",
                    parent=self
                )
            elif msg_type == "convert_error":
                err, trace = item[1], item[2]
                self.btn_run_convert.config(state=tk.NORMAL)
                self.btn_open_convert_dir.config(state=tk.NORMAL)
                self.convert_status_lbl.config(text="❌ Lỗi khi chuyển đổi định dạng!", foreground=self.palette.DANGER)
                self._log(f"❌ Lỗi khi chuyển đổi định dạng: {err}\n{trace}", "error")
                messagebox.showerror("Lỗi chuyển đổi", f"Đã xảy ra lỗi khi chuyển đổi tệp:\n{err}", parent=self)
            elif msg_type == "inspect_prog":
                val, msg = item[1], item[2]
                self.inspect_prog_bar["value"] = val
                self.inspect_status_lbl.config(text=f"⏳ {msg}...", foreground=self.palette.PRIMARY)
                self._log(f"  {msg}", "info")
            elif msg_type == "inspect_file_done":
                f_p, res = item[1], item[2]
                st = res.get("status")
                r_cnt = res.get("total_rows", 0)
                e_cnt = res.get("error_count", 0)
                w_cnt = res.get("warning_count", 0)
                for child in self.inspect_tree.get_children():
                    vals = list(self.inspect_tree.item(child, "values"))
                    if str(f_p) in vals or f_p.name in vals:
                        vals[2] = "✅ Hoàn tất" if st == "success" else f"❌ {st}"
                        vals[3] = f"{r_cnt:,}"
                        vals[4] = f"{e_cnt}" if e_cnt > 0 else "0"
                        vals[5] = f"{w_cnt}" if w_cnt > 0 else "0"
                        self.inspect_tree.item(child, values=vals)
            elif msg_type == "inspect_done":
                out_p, total_files, t_rows, t_err, t_warn, t_valid = item[1], item[2], item[3], item[4], item[5], item[6]
                self.inspect_prog_bar["value"] = 100
                self.btn_run_inspect.config(state=tk.NORMAL)
                self.btn_cancel_inspect.config(state=tk.DISABLED)
                if out_p:
                    self.btn_open_inspect_res.config(state=tk.NORMAL)
                    self.last_inspect_output_file = out_p
                    self.last_created_file = out_p
                self.inspect_status_lbl.config(text=f"✅ Đã kiểm tra xong ({t_rows:,} dòng)!", foreground=self.palette.SUCCESS)
                self.inspect_summary_lbl.config(
                    text=f"Tổng: {t_rows:,} dòng | Hợp lệ: {t_valid:,} | Cảnh báo: {t_warn:,} | Lỗi: {t_err:,}",
                    foreground=self.palette.DANGER if t_err > 0 else (self.palette.WARNING if t_warn > 0 else self.palette.SUCCESS)
                )
                self._log(f"🎉 Hoàn tất kiểm tra & chuẩn hoá: {t_rows:,} dòng (Hợp lệ: {t_valid:,}, Cảnh báo: {t_warn:,}, Lỗi: {t_err:,})", "success" if t_err == 0 else "warning")
                msg_body = f"Đã kiểm tra xong dữ liệu 28 cột!\n\n• Tổng số dòng: {t_rows:,}\n• Hợp lệ: {t_valid:,}\n• Cảnh báo: {t_warn:,}\n• Lỗi: {t_err:,}"
                if out_p:
                    msg_body += f"\n\nTệp kết quả: {out_p.name}"
                messagebox.showinfo("Hoàn tất Kiểm tra & Chuẩn hoá", msg_body, parent=self)
            elif msg_type == "inspect_error":
                err, trace = item[1], item[2]
                self.btn_run_inspect.config(state=tk.NORMAL)
                self.btn_cancel_inspect.config(state=tk.DISABLED)
                self.inspect_status_lbl.config(text="❌ Lỗi khi kiểm tra dữ liệu!", foreground=self.palette.DANGER)
                self._log(f"❌ Lỗi khi kiểm tra dữ liệu: {err}\n{trace}", "error")
                messagebox.showerror("Lỗi kiểm tra", f"Đã xảy ra lỗi khi kiểm tra và chuẩn hoá file:\n{err}", parent=self)
            elif msg_type == "encoding_prog":
                val, msg = item[1], item[2]
                self.encoding_prog_bar["value"] = val
                self.encoding_status_lbl.config(text=f"⏳ {msg}...", foreground=self.palette.PRIMARY)
                self._log(f"  {msg}", "info")
            elif msg_type == "encoding_file_done":
                f_p, ok, out_p, msg_text, cell_count = item[1], item[2], item[3], item[4], item[5]
                for child in self.encoding_tree.get_children():
                    vals = list(self.encoding_tree.item(child, "values"))
                    if str(f_p) in vals or f_p.name in vals:
                        vals[2] = "✅ Thành công" if ok else "❌ Lỗi"
                        vals[3] = f"{cell_count:,} ô" if ok else msg_text
                        self.encoding_tree.item(child, values=vals)
            elif msg_type == "encoding_done":
                succ, fail, tot_cells, last_out = item[1], item[2], item[3], item[4]
                self.encoding_prog_bar["value"] = 100
                self.btn_run_encoding.config(state=tk.NORMAL)
                self.btn_cancel_encoding.config(state=tk.DISABLED)
                if last_out:
                    self.btn_open_encoding_res.config(state=tk.NORMAL)
                    self.last_encoding_output_file = last_out
                    self.last_created_file = last_out
                self.encoding_status_lbl.config(text=f"✅ Đã chuyển đổi {tot_cells:,} ô sang Unicode!", foreground=self.palette.SUCCESS)
                self._log(f"🎉 Hoàn tất chuyển mã: {succ} tệp thành công, {fail} lỗi, tổng {tot_cells:,} ô đã chuyển sang Unicode.", "success" if fail == 0 else "warning")
                messagebox.showinfo(
                    "Hoàn tất Chuyển mã Unicode",
                    f"Đã hoàn thành chuyển đổi bảng mã sang Unicode!\n\n"
                    f"• Số tệp thành công: {succ}\n"
                    f"• Số tệp lỗi: {fail}\n"
                    f"• Tổng số ô đã đổi sang Unicode: {tot_cells:,}",
                    parent=self
                )
            elif msg_type == "encoding_error":
                err, trace = item[1], item[2]
                self.btn_run_encoding.config(state=tk.NORMAL)
                self.btn_cancel_encoding.config(state=tk.DISABLED)
                self.encoding_status_lbl.config(text="❌ Lỗi khi chuyển mã!", foreground=self.palette.DANGER)
                self._log(f"❌ Lỗi khi chuyển đổi bảng mã: {err}\n{trace}", "error")
                messagebox.showerror("Lỗi chuyển mã", f"Đã xảy ra lỗi khi chuyển đổi bảng mã:\n{err}", parent=self)
            elif msg_type == "ocr_prog":
                val, msg = item[1], item[2]
                self.ocr_prog_bar["value"] = val
                self.ocr_status_lbl.config(text=f"⏳ {msg}...", foreground=self.palette.PRIMARY)
            elif msg_type == "ocr_file_status":
                img_p, st_text, st_count = item[1], item[2], item[3]
                for child in self.ocr_tree.get_children():
                    vals = list(self.ocr_tree.item(child, "values"))
                    if str(img_p) in vals or img_p.name in vals:
                        vals[3] = st_text
                        vals[4] = f"{st_count} HS" if st_count > 0 else "---"
                        self.ocr_tree.item(child, values=vals)
            elif msg_type == "ocr_done":
                res = item[1]
                self.ocr_prog_bar["value"] = 100
                self.btn_run_ocr.config(state=tk.NORMAL)
                self.btn_cancel_ocr.config(state=tk.DISABLED)
                st = res.get("status")
                out_p = res.get("output_file")
                tot_stu = res.get("total_students", 0)
                succ_img = res.get("success_images", 0)
                tot_img = res.get("total_images", 0)
                suspects = res.get("suspect_rows_count", 0)

                if st == "success" and out_p:
                    self.btn_open_ocr_file.config(state=tk.NORMAL)
                    self.last_ocr_output_file = Path(out_p)
                    self.last_created_file = Path(out_p)
                    self.ocr_status_lbl.config(
                        text=f"✅ Hoàn tất OCR: {tot_stu:,} học sinh ({succ_img}/{tot_img} ảnh)!",
                        foreground=self.palette.SUCCESS
                    )
                    msg_lines = [
                        f"Đã trích xuất thành công {tot_stu:,} học sinh từ {succ_img}/{tot_img} ảnh!",
                        f"• Dòng cần kiểm tra lại (viết tay/mờ): {suspects}",
                        f"• Tệp Excel: {Path(out_p).name}",
                    ]
                    messagebox.showinfo("Hoàn tất OCR", "\n".join(msg_lines), parent=self)
                else:
                    self.ocr_status_lbl.config(
                        text="⚠️ Không trích xuất được học sinh nào!",
                        foreground=self.palette.WARNING
                    )
                    messagebox.showwarning(
                        "Không có dữ liệu OCR",
                        res.get("message", "Không trích xuất được dữ liệu học sinh hợp lệ."),
                        parent=self
                    )
            elif msg_type == "ocr_error":
                err, trace = item[1], item[2]
                self.btn_run_ocr.config(state=tk.NORMAL)
                self.btn_cancel_ocr.config(state=tk.DISABLED)
                self.ocr_status_lbl.config(text="❌ Lỗi khi chạy OCR!", foreground=self.palette.DANGER)
                self._log(f"❌ Lỗi khi chạy OCR: {err}\n{trace}", "error")
                messagebox.showerror("Lỗi OCR", f"Đã xảy ra lỗi trong quá trình chạy OCR:\n{err}", parent=self)

        self.after(100, self._start_queue_poller)

    # =========================================================================
    # TAB 2: KIỂM TRA & CHUẨN HOÁ DỮ LIỆU (28 CỘT CHUẨN)
    # =========================================================================
    def _build_tab_inspect(self, parent):
        p = self.palette

        # Khung Điều khiển & Tiến trình (Ở ĐÁY)
        ctrl_card = ttk.LabelFrame(parent, text="2. Tùy chọn & Bắt đầu kiểm tra", padding=8)
        ctrl_card.pack(side=tk.BOTTOM, fill=tk.X, pady=(4, 2))

        r_opt = ttk.Frame(ctrl_card)
        r_opt.pack(fill=tk.X, pady=(0, 4))

        ttk.Label(r_opt, text="Thư mục lưu:", font=FONT_LABEL).pack(side=tk.LEFT, padx=(0, 4))
        self.inspect_out_entry = ttk.Entry(r_opt, textvariable=self.inspect_out_dir_var, font=FONT_LABEL)
        self.inspect_out_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 6))
        ttk.Button(r_opt, text="📁 Chọn...", command=self._on_choose_inspect_out_dir).pack(side=tk.LEFT, padx=(0, 10))

        ttk.Checkbutton(
            r_opt,
            text="Tự động mở tệp Excel kết quả sau khi kiểm tra xong",
            variable=self.inspect_auto_open_var
        ).pack(side=tk.LEFT)

        r_prog = ttk.Frame(ctrl_card)
        r_prog.pack(fill=tk.X, pady=(2, 4))

        self.inspect_prog_bar = ttk.Progressbar(r_prog, orient=tk.HORIZONTAL, mode='determinate')
        self.inspect_prog_bar.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))

        self.inspect_status_lbl = ttk.Label(r_prog, text="Sẵn sàng", font=FONT_LABEL_BOLD, foreground=p.PRIMARY)
        self.inspect_status_lbl.pack(side=tk.RIGHT)

        r_btn = ttk.Frame(ctrl_card)
        r_btn.pack(fill=tk.X, pady=(2, 0))

        self.btn_open_inspect_res = ttk.Button(
            r_btn,
            text="📄 Mở file Excel vừa tạo",
            command=self._open_last_inspect_file,
            state="disabled"
        )
        self.btn_open_inspect_res.pack(side=tk.LEFT, padx=(0, 6))

        self.inspect_summary_lbl = ttk.Label(r_btn, text="", font=FONT_LABEL_BOLD)
        self.inspect_summary_lbl.pack(side=tk.LEFT, padx=10)

        self.btn_cancel_inspect = ttk.Button(
            r_btn,
            text="🛑 HỦY",
            command=self._cancel_inspect,
            state="disabled"
        )
        self.btn_cancel_inspect.pack(side=tk.RIGHT, padx=(6, 0))

        self.btn_run_inspect = ttk.Button(
            r_btn,
            text="⚡ BẮT ĐẦU KIỂM TRA & CHUẨN HOÁ",
            command=self._start_inspect_thread,
            style="Action.TButton"
        )
        self.btn_run_inspect.pack(side=tk.RIGHT)

        # Khung Danh sách tệp nguồn (Ở TRÊN, CO DÃN)
        top_card = ttk.LabelFrame(parent, text="1. Danh sách tệp Excel 28 cột cần kiểm tra", padding=6)
        top_card.pack(side=tk.TOP, fill=tk.BOTH, expand=True, pady=(2, 2))

        btn_row = ttk.Frame(top_card)
        btn_row.pack(fill=tk.X, pady=(0, 4))

        ttk.Button(btn_row, text="📂 Chọn tệp Excel...", command=self._on_choose_inspect_files).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(btn_row, text="📁 Chọn cả thư mục...", command=self._on_choose_inspect_folder).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(btn_row, text="❌ Xóa danh sách", command=self._clear_inspect_list).pack(side=tk.LEFT, padx=(0, 16))

        self.inspect_files_count_lbl = ttk.Label(btn_row, text="Chưa chọn tệp nào", font=FONT_LABEL_BOLD)
        self.inspect_files_count_lbl.pack(side=tk.LEFT)

        i_cols = ("stt", "file_name", "status", "rows", "errors", "warnings")
        self.inspect_tree = ttk.Treeview(top_card, columns=i_cols, show="headings", height=5)
        self.inspect_tree.heading("stt", text="STT")
        self.inspect_tree.heading("file_name", text="Tên tệp Excel")
        self.inspect_tree.heading("status", text="Trạng thái kiểm tra")
        self.inspect_tree.heading("rows", text="Số dòng")
        self.inspect_tree.heading("errors", text="Số lỗi")
        self.inspect_tree.heading("warnings", text="Số cảnh báo")

        self.inspect_tree.column("stt", width=45, anchor=tk.CENTER)
        self.inspect_tree.column("file_name", width=360, anchor=tk.W)
        self.inspect_tree.column("status", width=180, anchor=tk.CENTER)
        self.inspect_tree.column("rows", width=90, anchor=tk.CENTER)
        self.inspect_tree.column("errors", width=80, anchor=tk.CENTER)
        self.inspect_tree.column("warnings", width=100, anchor=tk.CENTER)

        i_sb = ttk.Scrollbar(top_card, orient=tk.VERTICAL, command=self.inspect_tree.yview)
        self.inspect_tree.configure(yscrollcommand=i_sb.set)
        self.inspect_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        i_sb.pack(side=tk.RIGHT, fill=tk.Y)

    def _on_choose_inspect_files(self):
        files = filedialog.askopenfilenames(
            title="Chọn các tệp Excel 28 cột cần kiểm tra",
            filetypes=[("Excel Files", "*.xlsx;*.xls")],
            parent=self
        )
        if not files:
            return
        for f in files:
            p = Path(f)
            if p not in self.inspect_files_list:
                self.inspect_files_list.append(p)
        self._refresh_inspect_tree()

    def _on_choose_inspect_folder(self):
        f_dir = filedialog.askdirectory(title="Chọn thư mục chứa tệp Excel cần kiểm tra", parent=self)
        if not f_dir:
            return
        p_dir = Path(f_dir)
        found = list(p_dir.glob("*.xlsx")) + list(p_dir.glob("*.xls"))
        for f in found:
            if not f.name.startswith("~$") and f not in self.inspect_files_list:
                self.inspect_files_list.append(f)
        self._refresh_inspect_tree()

    def _clear_inspect_list(self):
        self.inspect_files_list.clear()
        self._refresh_inspect_tree()
        self.inspect_summary_lbl.config(text="")
        self.btn_open_inspect_res.config(state=tk.DISABLED)

    def _refresh_inspect_tree(self):
        for item in self.inspect_tree.get_children():
            self.inspect_tree.delete(item)
        for idx, f in enumerate(self.inspect_files_list, start=1):
            self.inspect_tree.insert("", tk.END, values=(idx, f.name, "Chờ kiểm tra", "---", "---", "---"))
        count = len(self.inspect_files_list)
        self.inspect_files_count_lbl.config(
            text=f"Đã chọn {count} tệp Excel" if count > 0 else "Chưa chọn tệp nào",
            foreground=self.palette.SUCCESS if count > 0 else self.palette.TEXT_MUTED
        )

    def _on_choose_inspect_out_dir(self):
        d = filedialog.askdirectory(title="Chọn thư mục lưu tệp kết quả kiểm tra", parent=self)
        if d:
            self.inspect_out_dir_var.set(d)

    def _open_last_inspect_file(self):
        if self.last_inspect_output_file and self.last_inspect_output_file.exists():
            open_path_in_os(self.last_inspect_output_file)
        else:
            messagebox.showinfo("Thông báo", "Chưa có tệp kết quả nào hoặc tệp không tồn tại.", parent=self)

    def _cancel_inspect(self):
        self.inspect_cancel_requested = True
        self.inspect_status_lbl.config(text="Đang dừng tiến trình...", foreground=self.palette.WARNING)

    def _start_inspect_thread(self):
        if not self.inspect_files_list:
            messagebox.showwarning("Chưa chọn tệp", "Vui lòng chọn ít nhất một tệp Excel để kiểm tra!", parent=self)
            return

        self.inspect_cancel_requested = False
        self.btn_run_inspect.config(state=tk.DISABLED)
        self.btn_cancel_inspect.config(state=tk.NORMAL)
        self.inspect_prog_bar["value"] = 0
        self.inspect_status_lbl.config(text="Bắt đầu kiểm tra...", foreground=self.palette.PRIMARY)

        files = list(self.inspect_files_list)
        out_dir_str = self.inspect_out_dir_var.get().strip()
        out_dir = Path(out_dir_str) if out_dir_str else None
        auto_open = self.inspect_auto_open_var.get()

        t = threading.Thread(
            target=self._run_inspect_worker,
            args=(files, out_dir, auto_open),
            daemon=True
        )
        t.start()

    def _run_inspect_worker(self, files: list[Path], out_dir: Path | None, auto_open: bool):
        try:
            total_files = len(files)
            last_out: Path | None = None
            sum_rows = 0
            sum_errors = 0
            sum_warnings = 0
            sum_valid = 0

            for idx, f in enumerate(files, start=1):
                if self.inspect_cancel_requested:
                    self.msg_queue.put(("log", f"🛑 Đã dừng kiểm tra theo yêu cầu tại tệp {f.name}.", "warning"))
                    break

                prog_pct = int(((idx - 1) / total_files) * 100)
                self.msg_queue.put(("inspect_prog", prog_pct, f"[{idx}/{total_files}] Đang kiểm tra: {f.name}"))
                self.msg_queue.put(("log", f"🔍 [{idx}/{total_files}] Bắt đầu kiểm tra dữ liệu: {f.name}...", "info"))

                if out_dir:
                    out_dir.mkdir(parents=True, exist_ok=True)
                    out_file = out_dir / f"{f.stem}_DaKiemTra.xlsx"
                else:
                    out_file = f.parent / f"{f.stem}_DaKiemTra.xlsx"

                def p_cb(cur, tot, msg):
                    file_prog = int(((idx - 1) / total_files + (cur / tot) / total_files) * 100)
                    self.msg_queue.put(("inspect_prog", file_prog, f"[{idx}/{total_files}] {f.name}: dòng {cur}/{tot}"))

                res = self.inspector.process_file(
                    input_file=f,
                    output_file=out_file,
                    progress_callback=p_cb,
                    cancel_token=lambda: self.inspect_cancel_requested
                )

                if res.get("status") == "success":
                    last_out = out_file
                    r_cnt = res.get("total_rows", 0)
                    e_cnt = res.get("error_count", 0)
                    w_cnt = res.get("warning_count", 0)
                    v_cnt = res.get("valid_count", 0)

                    sum_rows += r_cnt
                    sum_errors += e_cnt
                    sum_warnings += w_cnt
                    sum_valid += v_cnt

                    self.msg_queue.put(("inspect_file_done", f, res))
                    self.msg_queue.put(("log", f"  ✅ {f.name}: {r_cnt:,} dòng (Hợp lệ: {v_cnt:,}, Cảnh báo: {w_cnt:,}, Lỗi: {e_cnt:,}) ➔ Lưu tại: {out_file.name}", "success" if e_cnt == 0 else "warning"))
                else:
                    self.msg_queue.put(("inspect_file_done", f, res))
                    self.msg_queue.put(("log", f"  ❌ Lỗi khi kiểm tra {f.name}: {res.get('message')}", "error"))

            self.msg_queue.put(("inspect_done", last_out, total_files, sum_rows, sum_errors, sum_warnings, sum_valid))

            if auto_open and last_out and last_out.exists() and total_files == 1:
                open_path_in_os(last_out)

        except Exception as e:
            trace = traceback.format_exc()
            self.msg_queue.put(("inspect_error", str(e), trace))

    # =========================================================================
    # SUB-TAB 4.2: CHUYỂN VNI, TCVN ➔ UNICODE
    # =========================================================================
    def _build_tab_encoding(self, parent):
        p = self.palette

        ctrl_card = ttk.LabelFrame(parent, text="2. Tùy chọn & Bắt đầu chuyển mã", padding=8)
        ctrl_card.pack(side=tk.BOTTOM, fill=tk.X, pady=(4, 2))

        r_opt = ttk.Frame(ctrl_card)
        r_opt.pack(fill=tk.X, pady=(0, 4))

        ttk.Checkbutton(
            r_opt,
            text="Ghi đè trực tiếp lên tệp gốc (mặc định tạo tệp mới có đuôi _Unicode.xlsx)",
            variable=self.encoding_overwrite_var
        ).pack(side=tk.LEFT)

        r_prog = ttk.Frame(ctrl_card)
        r_prog.pack(fill=tk.X, pady=(2, 4))

        self.encoding_prog_bar = ttk.Progressbar(r_prog, orient=tk.HORIZONTAL, mode='determinate')
        self.encoding_prog_bar.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))

        self.encoding_status_lbl = ttk.Label(r_prog, text="Sẵn sàng", font=FONT_LABEL_BOLD, foreground=p.PRIMARY)
        self.encoding_status_lbl.pack(side=tk.RIGHT)

        r_btn = ttk.Frame(ctrl_card)
        r_btn.pack(fill=tk.X, pady=(2, 0))

        self.btn_open_encoding_res = ttk.Button(
            r_btn,
            text="📄 Mở tệp vừa chuyển đổi",
            command=self._open_last_encoding_file,
            state="disabled"
        )
        self.btn_open_encoding_res.pack(side=tk.LEFT, padx=(0, 6))

        self.btn_cancel_encoding = ttk.Button(
            r_btn,
            text="🛑 HỦY",
            command=self._cancel_encoding,
            state="disabled"
        )
        self.btn_cancel_encoding.pack(side=tk.RIGHT, padx=(6, 0))

        self.btn_run_encoding = ttk.Button(
            r_btn,
            text="⚡ BẮT ĐẦU CHUYỂN MÃ UNICODE",
            command=self._start_encoding_thread,
            style="Action.TButton"
        )
        self.btn_run_encoding.pack(side=tk.RIGHT)

        # Khung danh sách tệp
        top_card = ttk.LabelFrame(parent, text="1. Danh sách tệp Excel (.xlsx) cần chuyển mã sang Unicode", padding=6)
        top_card.pack(side=tk.TOP, fill=tk.BOTH, expand=True, pady=(2, 2))

        btn_row = ttk.Frame(top_card)
        btn_row.pack(fill=tk.X, pady=(0, 4))

        ttk.Button(btn_row, text="📂 Chọn tệp Excel...", command=self._on_choose_encoding_files).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(btn_row, text="📁 Chọn cả thư mục...", command=self._on_choose_encoding_folder).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(btn_row, text="❌ Xóa danh sách", command=self._clear_encoding_list).pack(side=tk.LEFT, padx=(0, 16))

        self.encoding_files_count_lbl = ttk.Label(btn_row, text="Chưa chọn tệp nào", font=FONT_LABEL_BOLD)
        self.encoding_files_count_lbl.pack(side=tk.LEFT)

        e_cols = ("stt", "file_name", "status", "cells")
        self.encoding_tree = ttk.Treeview(top_card, columns=e_cols, show="headings", height=5)
        self.encoding_tree.heading("stt", text="STT")
        self.encoding_tree.heading("file_name", text="Tên tệp Excel")
        self.encoding_tree.heading("status", text="Trạng thái")
        self.encoding_tree.heading("cells", text="Số ô đã chuyển sang Unicode")

        self.encoding_tree.column("stt", width=50, anchor=tk.CENTER)
        self.encoding_tree.column("file_name", width=420, anchor=tk.W)
        self.encoding_tree.column("status", width=200, anchor=tk.CENTER)
        self.encoding_tree.column("cells", width=180, anchor=tk.CENTER)

        e_sb = ttk.Scrollbar(top_card, orient=tk.VERTICAL, command=self.encoding_tree.yview)
        self.encoding_tree.configure(yscrollcommand=e_sb.set)
        self.encoding_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        e_sb.pack(side=tk.RIGHT, fill=tk.Y)

    def _on_choose_encoding_files(self):
        files = filedialog.askopenfilenames(
            title="Chọn các tệp Excel cần chuyển mã sang Unicode",
            filetypes=[("Excel OpenXML", "*.xlsx")],
            parent=self
        )
        if not files:
            return
        for f in files:
            p = Path(f)
            if p not in self.encoding_files_list:
                self.encoding_files_list.append(p)
        self._refresh_encoding_tree()

    def _on_choose_encoding_folder(self):
        f_dir = filedialog.askdirectory(title="Chọn thư mục chứa tệp Excel cần chuyển mã", parent=self)
        if not f_dir:
            return
        p_dir = Path(f_dir)
        found = list(p_dir.glob("*.xlsx"))
        for f in found:
            if not f.name.startswith("~$") and f not in self.encoding_files_list:
                self.encoding_files_list.append(f)
        self._refresh_encoding_tree()

    def _clear_encoding_list(self):
        self.encoding_files_list.clear()
        self._refresh_encoding_tree()
        self.btn_open_encoding_res.config(state=tk.DISABLED)

    def _refresh_encoding_tree(self):
        for item in self.encoding_tree.get_children():
            self.encoding_tree.delete(item)
        for idx, f in enumerate(self.encoding_files_list, start=1):
            self.encoding_tree.insert("", tk.END, values=(idx, f.name, "Chờ chuyển mã", "---"))
        count = len(self.encoding_files_list)
        self.encoding_files_count_lbl.config(
            text=f"Đã chọn {count} tệp Excel" if count > 0 else "Chưa chọn tệp nào",
            foreground=self.palette.SUCCESS if count > 0 else self.palette.TEXT_MUTED
        )

    def _open_last_encoding_file(self):
        if self.last_encoding_output_file and self.last_encoding_output_file.exists():
            open_path_in_os(self.last_encoding_output_file)
        else:
            messagebox.showinfo("Thông báo", "Chưa có tệp kết quả nào hoặc tệp không tồn tại.", parent=self)

    def _cancel_encoding(self):
        self.encoding_cancel_requested = True
        self.encoding_status_lbl.config(text="Đang dừng tiến trình...", foreground=self.palette.WARNING)

    def _start_encoding_thread(self):
        if not self.encoding_files_list:
            messagebox.showwarning("Chưa chọn tệp", "Vui lòng chọn ít nhất một tệp Excel để chuyển mã!", parent=self)
            return

        self.encoding_cancel_requested = False
        self.btn_run_encoding.config(state=tk.DISABLED)
        self.btn_cancel_encoding.config(state=tk.NORMAL)
        self.encoding_prog_bar["value"] = 0
        self.encoding_status_lbl.config(text="Bắt đầu chuyển mã...", foreground=self.palette.PRIMARY)

        files = list(self.encoding_files_list)
        overwrite = self.encoding_overwrite_var.get()

        t = threading.Thread(
            target=self._run_encoding_worker,
            args=(files, overwrite),
            daemon=True
        )
        t.start()

    def _run_encoding_worker(self, files: list[Path], overwrite: bool):
        try:
            total_files = len(files)
            last_out: Path | None = None
            succ_count = 0
            fail_count = 0
            total_cells = 0

            for idx, f in enumerate(files, start=1):
                if self.encoding_cancel_requested:
                    self.msg_queue.put(("log", f"🛑 Đã dừng chuyển mã theo yêu cầu tại tệp {f.name}.", "warning"))
                    break

                prog_pct = int(((idx - 1) / total_files) * 100)
                self.msg_queue.put(("encoding_prog", prog_pct, f"[{idx}/{total_files}] Đang chuyển mã: {f.name}"))
                self.msg_queue.put(("log", f"🔤 [{idx}/{total_files}] Đang quét và chuyển mã Unicode: {f.name}...", "info"))

                out_p = f if overwrite else f.parent / f"{f.stem}_Unicode.xlsx"

                ok, final_out, msg_text, cell_count = FileFormatConverter.convert_excel_encoding_to_unicode(f, out_p)
                if ok:
                    succ_count += 1
                    total_cells += cell_count
                    last_out = final_out
                    self.msg_queue.put(("encoding_file_done", f, True, final_out, msg_text, cell_count))
                    self.msg_queue.put(("log", f"  ✅ {f.name}: Đã chuyển đổi {cell_count:,} ô sang Unicode chuẩn. ➔ Lưu: {final_out.name}", "success"))
                else:
                    fail_count += 1
                    self.msg_queue.put(("encoding_file_done", f, False, None, msg_text, 0))
                    self.msg_queue.put(("log", f"  ❌ Lỗi khi chuyển mã {f.name}: {msg_text}", "error"))

            self.msg_queue.put(("encoding_done", succ_count, fail_count, total_cells, last_out))

        except Exception as e:
            trace = traceback.format_exc()
            self.msg_queue.put(("encoding_error", str(e), trace))

    # =========================================================================
    # TAB 5: OCR TRÍCH XUẤT ẢNH SỔ SANG EXCEL
    # =========================================================================
    def _build_tab_ocr(self, parent):
        p = self.palette

        # 1. Khung Cấu hình API & Prompt (Đặt ở trên cùng)
        cfg_card = ttk.LabelFrame(parent, text="1. Cấu hình AI OCR (API Key, Model & Prompt)", padding=6)
        cfg_card.pack(side=tk.TOP, fill=tk.X, pady=(0, 4))

        r_api = ttk.Frame(cfg_card)
        r_api.pack(fill=tk.X, pady=(0, 4))

        ttk.Label(r_api, text="API Key:", font=FONT_LABEL_BOLD).pack(side=tk.LEFT, padx=(0, 4))
        self.ocr_api_key_entry = ttk.Entry(r_api, textvariable=self.ocr_api_key_var, font=FONT_LABEL, show="•", width=26)
        self.ocr_api_key_entry.pack(side=tk.LEFT, padx=(0, 4))

        self.btn_toggle_ocr_key = ttk.Button(r_api, text="👁️", width=3, command=self._toggle_ocr_api_key_visibility)
        self.btn_toggle_ocr_key.pack(side=tk.LEFT, padx=(0, 12))

        ttk.Label(r_api, text="Base URL:", font=FONT_LABEL).pack(side=tk.LEFT, padx=(0, 4))
        self.ocr_base_url_entry = ttk.Entry(r_api, textvariable=self.ocr_base_url_var, font=FONT_LABEL, width=22)
        self.ocr_base_url_entry.pack(side=tk.LEFT, padx=(0, 12))

        ttk.Label(r_api, text="Model:", font=FONT_LABEL_BOLD).pack(side=tk.LEFT, padx=(0, 4))
        self.ocr_model_entry = ttk.Entry(r_api, textvariable=self.ocr_model_var, font=FONT_LABEL, width=26)
        self.ocr_model_entry.pack(side=tk.LEFT, padx=(0, 8))

        ttk.Button(r_api, text="💾 Lưu cấu hình", command=self._on_save_ocr_config).pack(side=tk.RIGHT)

        r_prompt = ttk.Frame(cfg_card)
        r_prompt.pack(fill=tk.X)

        ttk.Label(r_prompt, text="Prompt đang dùng:", font=FONT_LABEL_BOLD).pack(side=tk.LEFT, padx=(0, 6))
        self.ocr_prompt_lbl = ttk.Label(r_prompt, textvariable=self.ocr_prompt_source_var, font=FONT_MUTED, foreground=p.PRIMARY)
        self.ocr_prompt_lbl.pack(side=tk.LEFT, fill=tk.X, expand=True)

        ttk.Button(r_prompt, text="📂 Nạp file prompt.txt...", command=self._on_choose_ocr_prompt_file).pack(side=tk.RIGHT, padx=(6, 0))
        ttk.Button(r_prompt, text="📝 Xem / Sửa Prompt...", command=self._open_ocr_prompt_dialog).pack(side=tk.RIGHT)

        # 3. Khung Tùy chọn xuất file & Bắt đầu OCR (GHIM CỐ ĐỊNH Ở ĐÁY)
        ctrl_card = ttk.LabelFrame(parent, text="3. Đầu ra & Thực thi OCR", padding=8)
        ctrl_card.pack(side=tk.BOTTOM, fill=tk.X, pady=(4, 2))

        r_out = ttk.Frame(ctrl_card)
        r_out.pack(fill=tk.X, pady=(0, 4))

        ttk.Label(r_out, text="Tệp Excel kết quả:", font=FONT_LABEL_BOLD).pack(side=tk.LEFT, padx=(0, 6))
        self.ocr_out_entry = ttk.Entry(r_out, textvariable=self.ocr_out_var, font=FONT_LABEL)
        self.ocr_out_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 6))
        ttk.Button(r_out, text="💾 Lưu thành...", command=self._on_choose_ocr_output_path).pack(side=tk.LEFT)

        self.ocr_prog_bar = ttk.Progressbar(ctrl_card, orient=tk.HORIZONTAL, mode="determinate")
        self.ocr_prog_bar.pack(fill=tk.X, pady=(2, 6))

        r_act = ttk.Frame(ctrl_card)
        r_act.pack(fill=tk.X)

        self.ocr_status_lbl = ttk.Label(r_act, text="Sẵn sàng chạy OCR.", font=FONT_LABEL_BOLD, foreground=p.PRIMARY)
        self.ocr_status_lbl.pack(side=tk.LEFT)

        self.btn_open_ocr_file = ttk.Button(
            r_act,
            text="📄 Mở file Excel vừa tạo",
            command=self._open_last_ocr_file,
            state=tk.DISABLED
        )
        self.btn_open_ocr_file.pack(side=tk.RIGHT, padx=(8, 0))

        self.btn_cancel_ocr = ttk.Button(
            r_act,
            text="🛑 DỪNG",
            command=self._cancel_ocr,
            state=tk.DISABLED
        )
        self.btn_cancel_ocr.pack(side=tk.RIGHT, padx=(8, 0))

        self.btn_run_ocr = ttk.Button(
            r_act,
            text="⚡ BẮT ĐẦU CHẠY OCR",
            command=self._start_ocr_thread,
            style="Action.TButton"
        )
        self.btn_run_ocr.pack(side=tk.RIGHT)

        # 2. Khung Danh sách ảnh nguồn (Ở giữa, co dãn)
        mid_card = ttk.LabelFrame(parent, text="2. Danh sách ảnh sổ gốc cần trích xuất (.jpg, .jpeg, .png, .heic, .heif)", padding=6)
        mid_card.pack(side=tk.TOP, fill=tk.BOTH, expand=True, pady=(2, 2))

        btn_row = ttk.Frame(mid_card)
        btn_row.pack(fill=tk.X, pady=(0, 4))

        ttk.Button(btn_row, text="📁 Chọn thư mục ảnh trường...", command=self._on_choose_ocr_folder).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(btn_row, text="🖼️ Chọn nhiều ảnh lẻ...", command=self._on_choose_ocr_files).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(btn_row, text="❌ Xóa danh sách", command=self._clear_ocr_list).pack(side=tk.LEFT, padx=(0, 16))

        self.ocr_files_count_lbl = ttk.Label(btn_row, text="Chưa chọn ảnh nào", font=FONT_LABEL_BOLD)
        self.ocr_files_count_lbl.pack(side=tk.LEFT)

        o_cols = ("stt", "file_name", "size", "status", "students", "path")
        self.ocr_tree = ttk.Treeview(mid_card, columns=o_cols, show="headings", height=5)
        self.ocr_tree.heading("stt", text="STT")
        self.ocr_tree.heading("file_name", text="Tên file ảnh (Sắp xếp tự nhiên)")
        self.ocr_tree.heading("size", text="Kích thước")
        self.ocr_tree.heading("status", text="Trạng thái OCR")
        self.ocr_tree.heading("students", text="Số HS trích xuất")
        self.ocr_tree.heading("path", text="Đường dẫn")

        self.ocr_tree.column("stt", width=45, anchor=tk.CENTER)
        self.ocr_tree.column("file_name", width=240, anchor=tk.W)
        self.ocr_tree.column("size", width=95, anchor=tk.CENTER)
        self.ocr_tree.column("status", width=180, anchor=tk.CENTER)
        self.ocr_tree.column("students", width=120, anchor=tk.CENTER)
        self.ocr_tree.column("path", width=280, anchor=tk.W)

        o_sb = ttk.Scrollbar(mid_card, orient=tk.VERTICAL, command=self.ocr_tree.yview)
        self.ocr_tree.configure(yscrollcommand=o_sb.set)
        self.ocr_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        o_sb.pack(side=tk.RIGHT, fill=tk.Y)

    def _toggle_ocr_api_key_visibility(self):
        self.ocr_show_key = not self.ocr_show_key
        self.ocr_api_key_entry.config(show="" if self.ocr_show_key else "•")

    def _on_save_ocr_config(self):
        api_key = self.ocr_api_key_var.get().strip()
        base_url = self.ocr_base_url_var.get().strip() or DEFAULT_BASE_URL
        model_name = self.ocr_model_var.get().strip() or DEFAULT_MODEL_NAME
        self.ocr_config_mgr.save_config(
            api_key=api_key,
            base_url=base_url,
            model_name=model_name,
        )
        self._log(f"💾 Đã lưu cấu hình OCR (Model: {model_name}, Base URL: {base_url}).", "success")
        messagebox.showinfo("Đã lưu cấu hình", "Đã lưu cấu hình API Key, Base URL và Model thành công!", parent=self)

    def _on_choose_ocr_prompt_file(self):
        f = filedialog.askopenfilename(
            title="Chọn tệp prompt.txt cho OCR",
            filetypes=[("Text Files", "*.txt"), ("All Files", "*.*")],
            parent=self
        )
        if not f:
            return
        try:
            p = Path(f)
            txt = p.read_text(encoding="utf-8").strip()
            if txt:
                self.ocr_prompt_text = txt
                self.ocr_prompt_source_var.set(f"Tệp tùy chọn ({p.name})")
                self._log(f"📝 Đã nạp nội dung Prompt từ tệp: {p}", "info")
        except Exception as e:
            messagebox.showerror("Lỗi đọc file", f"Không thể đọc tệp prompt:\n{e}", parent=self)

    def _open_ocr_prompt_dialog(self):
        dlg = OCRPromptDialog(
            parent=self.winfo_toplevel(),
            current_prompt=self.ocr_prompt_text,
            default_prompt=self.ocr_config_mgr.get_default_prompt(),
            is_dark=self.is_dark,
        )
        dlg.wait_window()
        if dlg.saved_prompt is not None:
            self.ocr_prompt_text = dlg.saved_prompt
            self.ocr_prompt_source_var.set("Prompt tùy chỉnh trên giao diện")
            self.ocr_config_mgr.save_config(
                api_key=self.ocr_api_key_var.get().strip(),
                base_url=self.ocr_base_url_var.get().strip(),
                model_name=self.ocr_model_var.get().strip(),
                custom_prompt=self.ocr_prompt_text,
            )
            self._log("📝 Đã cập nhật nội dung Prompt OCR.", "success")

    def _on_choose_ocr_folder(self):
        d = filedialog.askdirectory(title="Chọn thư mục chứa ảnh sổ gốc (vd: 01, 02)", parent=self)
        if not d:
            return
        folder_p = Path(d)
        scan_res = OCRProcessor.scan_images([folder_p])
        if not scan_res:
            messagebox.showinfo(
                "Thông báo",
                f"Thư mục '{folder_p.name}' trống hoặc không chứa tệp ảnh hợp lệ (.jpg, .jpeg, .png, .heic, .heif)!",
                parent=self
            )
            return

        self.ocr_images_list = [item["path"] for item in scan_res]
        # Tự động tìm kiếm Prompt theo 3 cấp độ giống dự án so-hoa-vbcc
        prompt_txt, prompt_src = self.ocr_config_mgr.resolve_prompt_for_folder(folder_p)
        self.ocr_prompt_text = prompt_txt
        self.ocr_prompt_source_var.set(prompt_src)

        # Gợi ý đường dẫn lưu file kết quả mặc định: truong_<ten_thu_muc>_ketqua.xlsx
        default_out = folder_p.parent / f"truong_{folder_p.name}_ketqua.xlsx"
        self.ocr_out_var.set(str(default_out))

        self._refresh_ocr_tree(scan_res)
        self._log(
            f"📁 Đã nạp {len(scan_res)} ảnh từ thư mục '{folder_p.name}' (Năm học: {folder_p.parent.name}). Nguồn prompt: {prompt_src}",
            "info"
        )

    def _on_choose_ocr_files(self):
        files = filedialog.askopenfilenames(
            title="Chọn các tệp ảnh sổ gốc cần chạy OCR",
            filetypes=[
                ("Image Files", "*.jpg *.jpeg *.png *.heic *.heif *.JPG *.JPEG *.PNG *.HEIC *.HEIF"),
                ("All Files", "*.*"),
            ],
            parent=self
        )
        if not files:
            return
        scan_res = OCRProcessor.scan_images([Path(f) for f in files])
        if not scan_res:
            return
        self.ocr_images_list = [item["path"] for item in scan_res]
        first_parent = self.ocr_images_list[0].parent
        prompt_txt, prompt_src = self.ocr_config_mgr.resolve_prompt_for_folder(first_parent)
        self.ocr_prompt_text = prompt_txt
        self.ocr_prompt_source_var.set(prompt_src)
        self.ocr_out_var.set(str(first_parent / f"truong_{first_parent.name}_ketqua.xlsx"))
        self._refresh_ocr_tree(scan_res)

    def _clear_ocr_list(self):
        self.ocr_images_list.clear()
        for item in self.ocr_tree.get_children():
            self.ocr_tree.delete(item)
        self.ocr_files_count_lbl.config(text="Chưa chọn ảnh nào", foreground=self.palette.TEXT_MUTED)
        self.btn_open_ocr_file.config(state=tk.DISABLED)

    def _refresh_ocr_tree(self, scan_data: list[dict] | None = None):
        for item in self.ocr_tree.get_children():
            self.ocr_tree.delete(item)

        if not self.ocr_images_list:
            self.ocr_files_count_lbl.config(text="Chưa chọn ảnh nào", foreground=self.palette.TEXT_MUTED)
            return

        items = scan_data or OCRProcessor.scan_images(self.ocr_images_list)
        self.ocr_files_count_lbl.config(
            text=f"Đã chọn {len(items)} ảnh (đã sắp xếp tự nhiên 1, 2, ..., 9, 10)",
            foreground=self.palette.SUCCESS
        )
        for idx, it in enumerate(items, start=1):
            self.ocr_tree.insert(
                "",
                tk.END,
                values=(idx, it["name"], it["size_str"], it["status"], "---", str(it["path"]))
            )

    def _on_choose_ocr_output_path(self):
        init_name = Path(self.ocr_out_var.get()).name if self.ocr_out_var.get() else "ketqua_ocr.xlsx"
        f = filedialog.asksaveasfilename(
            title="Lưu tệp Excel kết quả OCR thành",
            initialfile=init_name,
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx")],
            parent=self
        )
        if f:
            self.ocr_out_var.set(f)

    def _open_last_ocr_file(self):
        if self.last_ocr_output_file and self.last_ocr_output_file.exists():
            open_path_in_os(self.last_ocr_output_file)
        else:
            messagebox.showinfo("Thông báo", "Chưa có tệp Excel kết quả OCR nào.", parent=self)

    def _cancel_ocr(self):
        if messagebox.askyesno("Xác nhận dừng OCR", "Bạn có chắc chắn muốn dừng tiến trình OCR đang chạy không?\n(Các ảnh đã đọc thành công trước đó vẫn sẽ được xuất ra file)", parent=self):
            self.ocr_cancel_requested = True
            self.btn_cancel_ocr.config(state=tk.DISABLED)
            self.ocr_status_lbl.config(text="⏳ Đang dừng tiến trình OCR...", foreground=self.palette.WARNING)

    def _start_ocr_thread(self):
        if not self.ocr_images_list:
            messagebox.showwarning("Chưa chọn ảnh", "Vui lòng chọn thư mục chứa ảnh hoặc các tệp ảnh cần chạy OCR!", parent=self)
            return

        api_key = self.ocr_api_key_var.get().strip()
        if not api_key:
            messagebox.showwarning("Thiếu API Key", "Vui lòng nhập API Key (CKEY_API_KEY) ở mục Cấu hình AI OCR!", parent=self)
            return

        out_path_str = self.ocr_out_var.get().strip()
        if not out_path_str:
            self._on_choose_ocr_output_path()
            out_path_str = self.ocr_out_var.get().strip()
            if not out_path_str:
                return

        base_url = self.ocr_base_url_var.get().strip() or DEFAULT_BASE_URL
        model_name = self.ocr_model_var.get().strip() or DEFAULT_MODEL_NAME
        prompt_text = self.ocr_prompt_text.strip() or self.ocr_config_mgr.get_default_prompt()

        # Tự động lưu lại cấu hình hiện hành
        self.ocr_config_mgr.save_config(api_key=api_key, base_url=base_url, model_name=model_name)

        self.ocr_cancel_requested = False
        self.btn_run_ocr.config(state=tk.DISABLED)
        self.btn_cancel_ocr.config(state=tk.NORMAL)
        self.btn_open_ocr_file.config(state=tk.DISABLED)
        self.ocr_prog_bar["value"] = 0
        self.ocr_status_lbl.config(text="⏳ Đang kết nối AI và trích xuất dữ liệu từ ảnh...", foreground=self.palette.WARNING)

        images = list(self.ocr_images_list)
        out_p = Path(out_path_str)

        def run():
            try:
                def prog_cb(cur, tot, msg):
                    val = (cur / tot * 100) if tot > 0 else 0
                    self.msg_queue.put(("ocr_prog", val, msg))

                def file_st_cb(img_p, st_str, count):
                    self.msg_queue.put(("ocr_file_status", img_p, st_str, count))

                res = self.ocr_processor.process_ocr_batch(
                    image_paths=images,
                    output_excel_path=out_p,
                    prompt_text=prompt_text,
                    api_key=api_key,
                    base_url=base_url,
                    model_name=model_name,
                    progress_callback=prog_cb,
                    file_status_callback=file_st_cb,
                    log_callback=lambda m, l: self._log(m, l),
                    cancel_token=lambda: self.ocr_cancel_requested,
                )
                self.msg_queue.put(("ocr_done", res))
            except Exception as e:
                trace = traceback.format_exc()
                self.msg_queue.put(("ocr_error", str(e), trace))

        threading.Thread(target=run, daemon=True).start()


class OCRPromptDialog(tk.Toplevel):
    """Hộp thoại Xem và Chỉnh sửa nội dung Prompt cho hệ thống AI OCR."""
    def __init__(self, parent, current_prompt: str, default_prompt: str, is_dark: bool | None = None):
        super().__init__(parent)
        self.palette = get_theme_palette(is_dark)
        self.default_prompt = default_prompt
        self.saved_prompt: str | None = None

        self.title("📝 Xem & Chỉnh sửa Prompt OCR")
        self.geometry("820x560")
        self.minsize(680, 420)
        self.configure(bg=self.palette.WINDOW_BG)
        self.transient(parent)
        self.grab_set()

        self._build_ui(current_prompt)
        self._center_window(parent)

    def _center_window(self, parent):
        self.update_idletasks()
        w, h = 820, 560
        if parent and parent.winfo_exists():
            x = parent.winfo_rootx() + (parent.winfo_width() // 2) - (w // 2)
            y = parent.winfo_rooty() + (parent.winfo_height() // 2) - (h // 2)
            self.geometry(f"{w}x{h}+{max(0, x)}+{max(0, y)}")

    def _build_ui(self, current_prompt: str):
        p = self.palette

        h_frame = tk.Frame(self, bg=p.HEADER_BANNER_BG, padx=14, pady=8)
        h_frame.pack(side=tk.TOP, fill=tk.X)
        tk.Label(
            h_frame,
            text="📝 NỘI DUNG PROMPT CHỈ THỊ CHO AI OCR",
            font=FONT_TITLE,
            bg=p.HEADER_BANNER_BG,
            fg=p.HEADER_BANNER_FG
        ).pack(anchor=tk.W)

        b_frame = ttk.Frame(self, padding=(12, 8))
        b_frame.pack(side=tk.BOTTOM, fill=tk.X)

        ttk.Button(b_frame, text="🔄 Khôi phục Mặc định", command=self._reset_default).pack(side=tk.LEFT)
        ttk.Button(b_frame, text="Đóng", command=self.destroy).pack(side=tk.RIGHT)
        ttk.Button(b_frame, text="✅ Lưu & Áp dụng Prompt", command=self._save, style="Action.TButton").pack(side=tk.RIGHT, padx=(0, 8))

        body = ttk.Frame(self, padding=10)
        body.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

        self.txt = ScrolledText(
            body,
            font=FONT_LOG,
            bg=p.LOG_BG,
            fg=p.LOG_FG,
            insertbackground=p.LOG_FG,
            wrap=tk.WORD
        )
        self.txt.pack(fill=tk.BOTH, expand=True)
        self.txt.insert("1.0", current_prompt)

    def _reset_default(self):
        self.txt.delete("1.0", tk.END)
        self.txt.insert("1.0", self.default_prompt)

    def _save(self):
        content = self.txt.get("1.0", tk.END).strip()
        if not content:
            messagebox.showwarning("Prompt trống", "Nội dung Prompt không được để trống!", parent=self)
            return
        self.saved_prompt = content
        self.destroy()



class ManageProfilesDialog(tk.Toplevel):
    """Hộp thoại Quản lý Kho Tri Thức Ánh Xạ & Quy Tắc Cột Dự Phòng trong data/mapping_profiles.json."""
    def __init__(self, parent, profile_mgr: MappingProfileManager, is_dark: bool | None = None):
        super().__init__(parent)
        self.palette = get_theme_palette(is_dark)
        self.profile_mgr = profile_mgr
        self.title("⚙️ Quản lý Tri Thức Ánh Xạ & Quy Tắc Dự Phòng")
        self.geometry("920x620")
        self.minsize(780, 500)
        self.configure(bg=self.palette.WINDOW_BG)
        self.transient(parent)
        self.grab_set()

        self._build_ui()
        self._refresh_profiles()
        self._refresh_fallbacks()
        self._center_window(parent)

    def _center_window(self, parent):
        self.update_idletasks()
        w, h = 920, 620
        if parent and parent.winfo_exists():
            x = parent.winfo_rootx() + (parent.winfo_width() // 2) - (w // 2)
            y = parent.winfo_rooty() + (parent.winfo_height() // 2) - (h // 2)
            self.geometry(f"{w}x{h}+{max(0, x)}+{max(0, y)}")
        else:
            sw = self.winfo_screenwidth()
            sh = self.winfo_screenheight()
            self.geometry(f"{w}x{h}+{(sw - w) // 2}+{(sh - h) // 2}")

    def _build_ui(self):
        p = self.palette

        # Header
        h_frame = tk.Frame(self, bg=p.HEADER_BANNER_BG, padx=14, pady=10)
        h_frame.pack(side=tk.TOP, fill=tk.X)
        tk.Label(
            h_frame,
            text="⚙️ QUẢN LÝ TRI THỨC ÁNH XẠ & QUY TẮC CỘT DỰ PHÒNG",
            font=FONT_TITLE,
            bg=p.HEADER_BANNER_BG,
            fg=p.HEADER_BANNER_FG
        ).pack(anchor=tk.W)
        tk.Label(
            h_frame,
            text="Kho tri thức lưu trữ các biến thể tiêu đề đã học và các quy tắc tự động lấy dữ liệu dự phòng.",
            font=FONT_SUBTITLE,
            bg=p.HEADER_BANNER_BG,
            fg=p.HEADER_SUB_FG
        ).pack(anchor=tk.W, pady=(2, 0))

        # Bottom close button (pack FIRST so it stays pinned at bottom)
        b_frame = ttk.Frame(self, padding=(10, 8))
        b_frame.pack(side=tk.BOTTOM, fill=tk.X)
        ttk.Button(b_frame, text="Đóng", command=self.destroy).pack(side=tk.RIGHT)

        # Notebook tabs (takes remaining space)
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=8, pady=(8, 4))

        # Tab 1: Từ điển biến thể cột đã học
        self.tab_profiles = ttk.Frame(self.notebook, padding=6)
        self.notebook.add(self.tab_profiles, text="📚 Từ điển Biến thể Cột đã học")
        self._build_tab_profiles(self.tab_profiles)

        # Tab 2: Quy tắc cột dự phòng (Fallback)
        self.tab_fallback = ttk.Frame(self.notebook, padding=6)
        self.notebook.add(self.tab_fallback, text="🔄 Quy tắc Cột dự phòng (Fallback)")
        self._build_tab_fallback(self.tab_fallback)

    def _build_tab_profiles(self, container):
        p = self.palette

        # Bottom button bar (pinned to bottom of tab)
        btn_bar = ttk.Frame(container, padding=(0, 6))
        btn_bar.pack(side=tk.BOTTOM, fill=tk.X)
        ttk.Button(btn_bar, text="🗑️ Xóa các Profile đã chọn", command=self._delete_selected_profiles).pack(side=tk.LEFT)
        ttk.Button(btn_bar, text="⚠️ Xóa TẤT CẢ Profile", command=self._delete_all_profiles).pack(side=tk.LEFT, padx=(8, 0))
        ttk.Button(btn_bar, text="🔄 Làm mới", command=self._refresh_profiles).pack(side=tk.LEFT, padx=(8, 0))

        # Treeview in middle
        tree_card = ttk.Frame(container)
        tree_card.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

        cols = ("stt", "name", "cols", "mapped", "date")
        self.tree = ttk.Treeview(tree_card, columns=cols, show="headings", selectmode="extended", height=8)
        self.tree.heading("stt", text="STT")
        self.tree.heading("name", text="Tên Profile đã lưu")
        self.tree.heading("cols", text="Số cột nguồn")
        self.tree.heading("mapped", text="Chi tiết ánh xạ")
        self.tree.heading("date", text="Ngày cập nhật")

        self.tree.column("stt", width=45, anchor=tk.CENTER)
        self.tree.column("name", width=250, anchor=tk.W)
        self.tree.column("cols", width=100, anchor=tk.CENTER)
        self.tree.column("mapped", width=260, anchor=tk.W)
        self.tree.column("date", width=140, anchor=tk.CENTER)

        sb = ttk.Scrollbar(tree_card, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sb.pack(side=tk.RIGHT, fill=tk.Y)

    def _build_tab_fallback(self, container):
        p = self.palette

        # Info card at top
        info_lbl = ttk.Label(
            container,
            text="💡 Quy tắc dự phòng cho phép một cột đích tự động lấy dữ liệu từ cột khác nếu file Excel nguồn không có cột riêng biệt.\nVí dụ chuẩn: Cột 17 (Hội đồng thi) tự động dự phòng lấy từ Cột 14 (Tên trường đã học).",
            font=FONT_MUTED,
            foreground=p.PRIMARY
        )
        info_lbl.pack(side=tk.TOP, anchor=tk.W, pady=(0, 6))

        # Form thêm/sửa quy tắc (pinned to bottom of tab, above close bar)
        form_frame = ttk.LabelFrame(container, text="➕ Thêm / Cập nhật Quy tắc Dự phòng", padding=10)
        form_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=(6, 2))

        col_options = [f"Cột {c['idx']}: {c['name']}" for c in TARGET_COLUMNS]

        row1 = ttk.Frame(form_frame)
        row1.pack(fill=tk.X, pady=2)

        ttk.Label(row1, text="Nếu Cột đích:", font=FONT_LABEL).pack(side=tk.LEFT)
        self.cb_fb_target = ttk.Combobox(row1, values=col_options, state="readonly", width=25)
        self.cb_fb_target.pack(side=tk.LEFT, padx=(4, 10))
        # Mặc định chọn Cột 17
        for opt in col_options:
            if opt.startswith("Cột 17:"):
                self.cb_fb_target.set(opt)
                break

        ttk.Label(row1, text="trống ➔ Dự phòng lấy từ:", font=FONT_LABEL).pack(side=tk.LEFT)
        self.cb_fb_source = ttk.Combobox(row1, values=col_options, state="readonly", width=25)
        self.cb_fb_source.pack(side=tk.LEFT, padx=(4, 10))
        # Mặc định chọn Cột 14
        for opt in col_options:
            if opt.startswith("Cột 14:"):
                self.cb_fb_source.set(opt)
                break

        ttk.Button(row1, text="💾 Lưu quy tắc", command=self._add_fallback_rule).pack(side=tk.LEFT, padx=(6, 0))
        ttk.Button(row1, text="🗑️ Xóa quy tắc đã chọn", command=self._delete_selected_fallback).pack(side=tk.RIGHT)

        # Fallback treeview (takes remaining middle space)
        tree_card = ttk.Frame(container)
        tree_card.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

        cols = ("target", "fallback", "desc")
        self.tree_fb = ttk.Treeview(tree_card, columns=cols, show="headings", height=7)
        self.tree_fb.heading("target", text="Cột đích cần dữ liệu")
        self.tree_fb.heading("fallback", text="Dự phòng lấy từ Cột")
        self.tree_fb.heading("desc", text="Mô tả hành vi khi ghép file")

        self.tree_fb.column("target", width=220, anchor=tk.W)
        self.tree_fb.column("fallback", width=220, anchor=tk.W)
        self.tree_fb.column("desc", width=340, anchor=tk.W)

        sb = ttk.Scrollbar(tree_card, orient=tk.VERTICAL, command=self.tree_fb.yview)
        self.tree_fb.configure(yscrollcommand=sb.set)
        self.tree_fb.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sb.pack(side=tk.RIGHT, fill=tk.Y)


    def _refresh_profiles(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        profiles = self.profile_mgr.get_all_profiles()
        for idx, p in enumerate(profiles, start=1):
            src_h = p.get("source_headers", [])
            mapping = p.get("mapping", {})
            mapped_count = sum(1 for v in mapping.values() if v is not None)
            ignored_count = len(src_h) - mapped_count
            mapped_desc = f"{mapped_count} cột ghép, {ignored_count} bỏ qua"
            self.tree.insert("", tk.END, iid=p.get("id"), values=(
                idx,
                p.get("name", "Không tên"),
                f"{len(src_h)} cột",
                mapped_desc,
                p.get("updated_at", "---")
            ))

    def _delete_selected_profiles(self):
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("Chưa chọn", "Vui lòng chọn ít nhất một Profile để xóa!", parent=self)
            return
        count = len(selected)
        if messagebox.askyesno("Xác nhận xóa", f"Bạn có chắc chắn muốn xóa {count} Profile đã chọn không?", parent=self):
            self.profile_mgr.delete_profiles(list(selected))
            self._refresh_profiles()

    def _delete_all_profiles(self):
        profiles = self.profile_mgr.get_all_profiles()
        if not profiles:
            messagebox.showinfo("Thông báo", "Hiện không có Profile nào trong danh sách.", parent=self)
            return
        if messagebox.askyesno(
            "Cảnh báo xóa tất cả",
            f"CẢNH BÁO: Bạn có chắc chắn muốn xóa TOÀN BỘ {len(profiles)} Profile đã lưu không?\n"
            "Toàn bộ cấu hình ánh xạ sẽ bị xóa và hành động này không thể hoàn tác!",
            parent=self
        ):
            self.profile_mgr.delete_all_profiles()
            self._refresh_profiles()

    def _refresh_fallbacks(self):
        for item in self.tree_fb.get_children():
            self.tree_fb.delete(item)
        fbs = self.profile_mgr.get_fallback_columns()
        col_map = {c["idx"]: c["name"] for c in TARGET_COLUMNS}
        for t_idx, fb_idx in sorted(fbs.items()):
            t_name = f"Cột {t_idx}: {col_map.get(t_idx, f'Cột {t_idx}')}"
            fb_name = f"Cột {fb_idx}: {col_map.get(fb_idx, f'Cột {fb_idx}')}"
            desc = f"Nếu Cột {t_idx} trống ➔ Tự động lấy dữ liệu từ Cột {fb_idx} ({col_map.get(fb_idx, '')})"
            self.tree_fb.insert("", tk.END, iid=str(t_idx), values=(t_name, fb_name, desc))

    def _add_fallback_rule(self):
        t_str = self.cb_fb_target.get()
        s_str = self.cb_fb_source.get()
        if not t_str or not s_str:
            messagebox.showwarning("Thiếu thông tin", "Vui lòng chọn cả Cột đích và Cột nguồn dự phòng!", parent=self)
            return
        try:
            t_idx = int(t_str.split(":")[0].replace("Cột", "").strip())
            s_idx = int(s_str.split(":")[0].replace("Cột", "").strip())
        except Exception:
            messagebox.showerror("Lỗi", "Không thể xác định chỉ số cột!", parent=self)
            return
        if t_idx == s_idx:
            messagebox.showwarning("Không hợp lệ", "Cột đích và Cột nguồn dự phòng không được trùng nhau!", parent=self)
            return
        self.profile_mgr.add_fallback_column(t_idx, s_idx)
        self._refresh_fallbacks()
        messagebox.showinfo("Thành công", f"Đã lưu quy tắc: Cột {t_idx} dự phòng lấy từ Cột {s_idx}!", parent=self)

    def _delete_selected_fallback(self):
        selected = self.tree_fb.selection()
        if not selected:
            messagebox.showwarning("Chưa chọn", "Vui lòng chọn một quy tắc dự phòng cần xóa trong bảng!", parent=self)
            return
        t_idx_str = selected[0]
        try:
            t_idx = int(t_idx_str)
        except Exception:
            return
        if messagebox.askyesno("Xác nhận xóa", f"Bạn có chắc chắn muốn xóa quy tắc dự phòng của Cột {t_idx} không?", parent=self):
            self.profile_mgr.delete_fallback_column(t_idx)
            self._refresh_fallbacks()


class CustomToolsDialog(tk.Toplevel):
    """
    Cửa sổ Hộp thoại độc lập của Mô-đun 3 (Tách / Gộp Excel).
    Hosting CustomToolsView bên trong.
    """
    def __init__(self, parent, is_dark: bool | None = None):
        super().__init__(parent)
        self.is_dark = IS_DARK if is_dark is None else is_dark
        self.palette = get_theme_palette(self.is_dark)
        self.title("Mô-đun 3: Bộ tiện ích Nhập & Tách Excel theo cột")
        self.geometry("1020x680")
        self.minsize(860, 560)
        self.transient(parent)
        self.grab_set()

        self.view = CustomToolsView(self, is_dark=is_dark)
        self.view.pack(fill=tk.BOTH, expand=True)
        self._center_window(parent)

    def _center_window(self, parent):
        self.update_idletasks()
        w, h = 1020, 680
        x = parent.winfo_rootx() + (parent.winfo_width() // 2) - (w // 2)
        y = parent.winfo_rooty() + (parent.winfo_height() // 2) - (h // 2)
        self.geometry(f"{w}x{h}+{max(0, x)}+{max(0, y)}")

    def __getattr__(self, name):
        return getattr(self.view, name)
