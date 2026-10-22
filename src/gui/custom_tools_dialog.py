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
    MappingProfileManager
)
from src.gui.column_mapping_dialog import ColumnMappingDialog
from src.file_converter import FileFormatConverter

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

        # Biến Tab 3 (Ghép & Chuẩn hóa vào Mẫu chuẩn 28 cột)
        self.smart_merger = SmartMerger()
        self.smart_files_list: list[Path] = []
        self.smart_out_var = tk.StringVar()
        self.smart_renumber_var = tk.BooleanVar(value=True)
        self.smart_fill_defaults_var = tk.BooleanVar(value=True)
        self.smart_all_sheets_var = tk.BooleanVar(value=True)
        self.last_smart_output_file: Path | None = None

        # Biến Tab 4 (Chuyển đổi định dạng tệp cũ .doc, .xls)
        self.convert_files_list: list[Path] = []
        self.convert_use_custom_dir = tk.BooleanVar(value=False)
        self.convert_custom_dir_var = tk.StringVar()
        self.convert_delete_source_var = tk.BooleanVar(value=False)
        self.convert_skip_existing_var = tk.BooleanVar(value=True)
        self.last_convert_dir: Path | None = None

        self._build_ui()
        self._start_queue_poller()

    def _build_ui(self):
        p = self.palette

        # 1. Header Banner
        hdr_frame = tk.Frame(self, bg=p.HEADER_BANNER_BG, padx=16, pady=8)
        hdr_frame.pack(fill=tk.X)

        title_lbl = tk.Label(
            hdr_frame,
            text="🛠️ CÔNG CỤ & BỘ TIỆN ÍCH MỞ RỘNG (EXCEL & WORD)",
            font=FONT_TITLE,
            bg=p.HEADER_BANNER_BG,
            fg=p.HEADER_BANNER_FG
        )
        title_lbl.pack(anchor=tk.W)

        sub_lbl = tk.Label(
            hdr_frame,
            text="Tách file theo cột, Gộp file cùng mẫu, Ghép & Chuẩn hóa mẫu 28 cột, và Chuyển đổi định dạng tệp cũ (.doc, .xls)",
            font=FONT_SUBTITLE,
            bg=p.HEADER_BANNER_BG,
            fg=p.HEADER_SUB_FG
        )
        sub_lbl.pack(anchor=tk.W, pady=(2, 0))

        # 2. Main Notebook (3 Tabs)
        main_container = ttk.Frame(self, style="Window.TFrame", padding=6)
        main_container.pack(fill=tk.BOTH, expand=True)

        self.notebook = ttk.Notebook(main_container)
        self.notebook.pack(fill=tk.BOTH, expand=True)

        # Tab 1: Tách file
        self.tab_split = ttk.Frame(self.notebook, padding=6)
        self.notebook.add(self.tab_split, text="  📤 TÁCH TỆP THEO CỘT  ")
        self._build_tab_split(self.tab_split)

        # Tab 2: Gộp file cùng cấu trúc
        self.tab_merge = ttk.Frame(self.notebook, padding=6)
        self.notebook.add(self.tab_merge, text="  📥 NHẬP / GỘP CÙNG CẤU TRÚC  ")
        self._build_tab_merge(self.tab_merge)

        # Tab 3: Ghép & Chuẩn hóa thông minh vào mẫu chuẩn 28 cột
        self.tab_smart = ttk.Frame(self.notebook, padding=6)
        self.notebook.add(self.tab_smart, text="  🧩 GHÉP & CHUẨN HÓA (MẪU CHUẨN)  ")
        self._build_tab_smart(self.tab_smart)

        # Tab 4: Chuyển đổi định dạng tệp cũ (.doc, .xls)
        self.tab_convert = ttk.Frame(self.notebook, padding=6)
        self.notebook.add(self.tab_convert, text="  🔄 CHUYỂN ĐỔI ĐỊNH DẠNG CŨ (.doc, .xls)  ")
        self._build_tab_convert(self.tab_convert)

        # 3. Vùng Console Log chung ở dưới cùng
        log_frame = ttk.LabelFrame(main_container, text="Nhật ký hoạt động & Thông báo", padding=4)
        log_frame.pack(fill=tk.BOTH, expand=False, pady=(4, 0))

        self.log_text = ScrolledText(
            log_frame,
            height=3,
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

            self.btn_run_smart_merge.config(state=tk.DISABLED)
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
                        resp_holder = {"result": (None, None, {}, True)}
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
                        log_callback=lambda m, l: self._log(m, l)
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
                        formatted_headers=payload["formatted_headers"],
                        suggested_mapping=payload["suggested_mapping"],
                        profile_mgr=self.smart_merger.profile_mgr,
                        is_dark=self.is_dark,
                        suggested_defaults=payload.get("suggested_defaults"),
                        matched_profile_name=payload.get("matched_profile_name")
                    )
                    dlg.wait_window()
                    resp_holder["result"] = dlg.result
                except Exception as ex:
                    self._log(f"❌ Lỗi hiển thị hộp thoại ánh xạ: {ex}", "error")
                    resp_holder["result"] = (None, None, {}, True)
                finally:
                    resp_event.set()
            elif msg_type == "smart_prog":
                val, msg = item[1], item[2]
                self.smart_prog["value"] = val
                self._log(f"  {msg}", "info")
            elif msg_type == "smart_done":
                res, out_p = item[1], item[2]
                self.smart_prog["value"] = 100
                self.btn_run_smart_merge.config(state=tk.NORMAL)
                self.btn_open_smart_file.config(state=tk.NORMAL)
                self.last_smart_output_file = out_p
                total_rows = res.get("total_rows", 0)
                sheets_p = res.get("sheets_processed", 0)

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

        self.after(100, self._start_queue_poller)


class ManageProfilesDialog(tk.Toplevel):
    """Hộp thoại Quản lý danh sách Profile ánh xạ cột đã lưu trong data/mapping_profiles.json."""
    def __init__(self, parent, profile_mgr: MappingProfileManager, is_dark: bool | None = None):
        super().__init__(parent)
        self.palette = get_theme_palette(is_dark)
        self.profile_mgr = profile_mgr
        self.title("⚙️ Quản lý Profiles Ánh Xạ Cột Đã Lưu")
        self.geometry("780x480")
        self.minsize(620, 360)
        self.configure(bg=self.palette.WINDOW_BG)
        self.transient(parent)
        self.grab_set()

        self._build_ui()
        self._refresh()

    def _build_ui(self):
        p = self.palette

        # Header
        h_frame = tk.Frame(self, bg=p.HEADER_BANNER_BG, padx=14, pady=10)
        h_frame.pack(fill=tk.X)
        tk.Label(
            h_frame,
            text="⚙️ QUẢN LÝ CÁC PROFILES ÁNH XẠ CỘT ĐÃ LƯU",
            font=FONT_TITLE,
            bg=p.HEADER_BANNER_BG,
            fg=p.HEADER_BANNER_FG
        ).pack(anchor=tk.W)
        tk.Label(
            h_frame,
            text="Khi quét file/sheet có cùng cấu trúc cột, hệ thống sẽ tự động ghép mà không cần hỏi lại.",
            font=FONT_SUBTITLE,
            bg=p.HEADER_BANNER_BG,
            fg=p.HEADER_SUB_FG
        ).pack(anchor=tk.W, pady=(2, 0))

        # Treeview
        tree_card = ttk.Frame(self, padding=8)
        tree_card.pack(fill=tk.BOTH, expand=True)

        cols = ("stt", "name", "cols", "mapped", "date")
        self.tree = ttk.Treeview(tree_card, columns=cols, show="headings")
        self.tree.heading("stt", text="STT")
        self.tree.heading("name", text="Tên Profile")
        self.tree.heading("cols", text="Số cột nguồn")
        self.tree.heading("mapped", text="Cột đã ánh xạ")
        self.tree.heading("date", text="Ngày cập nhật")

        self.tree.column("stt", width=45, anchor=tk.CENTER)
        self.tree.column("name", width=280, anchor=tk.W)
        self.tree.column("cols", width=95, anchor=tk.CENTER)
        self.tree.column("mapped", width=140, anchor=tk.CENTER)
        self.tree.column("date", width=160, anchor=tk.CENTER)

        sb = ttk.Scrollbar(tree_card, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sb.pack(side=tk.RIGHT, fill=tk.Y)

        # Bottom buttons
        b_frame = ttk.Frame(self, padding=10)
        b_frame.pack(fill=tk.X)
        ttk.Button(b_frame, text="🗑️ Xóa Profile đã chọn", command=self._delete_selected).pack(side=tk.LEFT)
        ttk.Button(b_frame, text="Đóng", command=self.destroy).pack(side=tk.RIGHT)

    def _refresh(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        profiles = self.profile_mgr.list_profiles()
        for idx, p in enumerate(profiles, start=1):
            raw_h = p.get("raw_headers", [])
            named_m = self.profile_mgr.get_named_mapping(p)
            mapped_count = sum(1 for v in named_m.values() if v is not None)
            self.tree.insert("", tk.END, iid=p.get("id"), values=(
                idx,
                p.get("name", "Không tên"),
                f"{len(raw_h)} cột",
                f"{mapped_count}/28 cột",
                p.get("updated_at", p.get("created_at", ""))
            ))

    def _delete_selected(self):
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("Chưa chọn", "Vui lòng chọn Profile cần xóa!", parent=self)
            return
        p_id = selected[0]
        if messagebox.askyesno("Xác nhận xóa", "Bạn có chắc chắn muốn xóa Profile này?", parent=self):
            self.profile_mgr.delete_profile(p_id)
            self._refresh()


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
