from __future__ import annotations

import sys
import re
import tkinter as tk
from tkinter import ttk, messagebox
from typing import Any

from src.gui.theme import (
    get_theme_palette, IS_DARK,
    FONT_TITLE, FONT_SUBTITLE, FONT_LABEL, FONT_LABEL_BOLD,
    FONT_MUTED
)
from src.smart_merger import (
    TARGET_COLUMNS, MappingProfileManager,
    normalize_text_for_match
)

IGNORE_OPTION = "-- Bỏ qua --"

TARGET_OPTIONS = [
    IGNORE_OPTION,
    "Cột 1: STT",
    "Cột 2: Số hiệu bằng",
    "Cột 3: Số vào sổ gốc cấp văn bằng",
    "Cột 4: Tên văn bằng",
    "Cột 5: Họ, chữ đệm và tên (Cả họ và tên)",
    "Cột 5: Họ đệm (Ghép Họ tên - phần 1)",
    "Cột 5: Tên (Ghép Họ tên - phần 2)",
    "Cột 6: Mã người học",
    "Cột 7: Số định danh cá nhân",
    "Cột 8: Hộ chiếu",
    "Cột 9: Ngày, tháng, năm sinh (Đầy đủ cả ngày tháng năm)",
    "Cột 9: Ngày sinh (Ghép Ngày sinh - phần 1)",
    "Cột 9: Tháng sinh (Ghép Ngày sinh - phần 2)",
    "Cột 9: Năm sinh (Ghép Ngày sinh - phần 3)",
    "Cột 10: Giới tính",
    "Cột 11: Dân tộc",
    "Cột 12: Quốc tịch",
    "Cột 13: Nơi sinh",
    "Cột 14: Tên trường đã học",
    "Cột 15: Niên khóa",
    "Cột 16: Điểm các môn thi",
    "Cột 17: Hội đồng thi",
    "Cột 18: Địa danh nơi cơ quan cấp bằng đặt trụ sở",
    "Cột 19: Ngày tháng năm cấp bằng",
    "Cột 20: Tên cơ quan cấp bằng",
    "Cột 21: Chức danh người ký bằng",
    "Cột 22: Họ, chữ đệm, tên người ký bằng",
    "Cột 23: Tình trạng văn bằng",
    "Cột 24: Ghi chú",
    "Cột 25: Trạng thái số hóa",
    "Cột 26: Thông tin chưa bảo đảm đầy đủ, chính xác và nguyên nhân",
    "Cột 27: Mã đơn vị",
    "Cột 28: Tên đơn vị"
]


def option_to_target_spec(opt: str) -> Any:
    """Chuyển đổi chuỗi tùy chọn trong combobox sang giá trị đặc tả target_spec."""
    if not opt or opt == IGNORE_OPTION:
        return None
    if "Họ đệm (Ghép Họ tên - phần 1)" in opt:
        return [5, 0]
    if "Tên (Ghép Họ tên - phần 2)" in opt:
        return [5, 1]
    if "Họ, chữ đệm và tên" in opt:
        return 5
    if "Ngày sinh (Ghép Ngày sinh - phần 1)" in opt:
        return [9, 0]
    if "Tháng sinh (Ghép Ngày sinh - phần 2)" in opt:
        return [9, 1]
    if "Năm sinh (Ghép Ngày sinh - phần 3)" in opt:
        return [9, 2]
    if "Ngày, tháng, năm sinh" in opt:
        return 9

    m = re.match(r"^Cột\s+(\d+):", opt)
    if m:
        return int(m.group(1))
    return None


def target_spec_to_option(spec: Any) -> str:
    """Chuyển đổi target_spec sang chuỗi hiển thị tương ứng trong combobox."""
    if spec is None:
        return IGNORE_OPTION
    if spec == [5, 0]:
        return "Cột 5: Họ đệm (Ghép Họ tên - phần 1)"
    if spec == [5, 1]:
        return "Cột 5: Tên (Ghép Họ tên - phần 2)"
    if spec == 5:
        return "Cột 5: Họ, chữ đệm và tên (Cả họ và tên)"
    if spec == [9, 0]:
        return "Cột 9: Ngày sinh (Ghép Ngày sinh - phần 1)"
    if spec == [9, 1]:
        return "Cột 9: Tháng sinh (Ghép Ngày sinh - phần 2)"
    if spec == [9, 2]:
        return "Cột 9: Năm sinh (Ghép Ngày sinh - phần 3)"
    if spec == 9:
        return "Cột 9: Ngày, tháng, năm sinh (Đầy đủ cả ngày tháng năm)"
    if isinstance(spec, int):
        for opt in TARGET_OPTIONS:
            if opt.startswith(f"Cột {spec}:"):
                return opt
    return IGNORE_OPTION


class ColumnMappingDialog(tk.Toplevel):
    """
    Hộp thoại Cấu hình Ánh xạ Cột thông minh (Mô hình Đảo ngược - Inverted Mapping):
    - Hiển thị danh sách N cột nguồn của tệp Excel hiện tại.
    - Cho phép người dùng chọn đích đến trong 28 cột mẫu hoặc chọn '-- Bỏ qua --'.
    - Hiển thị dữ liệu mẫu từ các dòng đầu tiên để người dùng dễ nhận diện.
    - Hỗ trợ 2 nút thoát rõ ràng: 'Bỏ qua Sheet này' và 'Dừng toàn bộ'.
    """

    def __init__(
        self,
        parent: tk.Tk | tk.Toplevel,
        file_name: str,
        sheet_name: str,
        raw_headers: list[str],
        formatted_headers: list[str] | None = None,
        suggested_mapping: dict[int, int | list[int] | None] | None = None,
        suggested_inverted_mapping: dict[str, Any] | None = None,
        profile_mgr: MappingProfileManager | None = None,
        is_dark: bool | None = None,
        suggested_defaults: dict[int, str] | None = None,
        matched_profile_name: str | None = None,
        sample_values: dict[int, str] | None = None,
        extra_cols_with_data: list[str] | None = None,
        **kwargs: Any
    ):
        super().__init__(parent)
        self.is_dark = IS_DARK if is_dark is None else is_dark
        self.palette = get_theme_palette(self.is_dark)
        self.profile_mgr = profile_mgr or MappingProfileManager()

        self.file_name = file_name
        self.sheet_name = sheet_name
        self.raw_headers = raw_headers
        self.formatted_headers = formatted_headers or [f"Cột {i+1}: {h}" for i, h in enumerate(raw_headers)]
        self.suggested_mapping = suggested_mapping or {}
        self.suggested_inverted_mapping = suggested_inverted_mapping or kwargs.get("suggested_inverted_mapping")
        self.suggested_defaults = suggested_defaults or {}
        self.matched_profile_name = matched_profile_name
        self.sample_values = sample_values or kwargs.get("sample_values") or {}
        self.extra_cols_with_data = set(extra_cols_with_data or kwargs.get("extra_cols_with_data") or [])

        # Kết quả: ("apply" | "skip" | "cancel", execution_mapping, profile_name, inverted_mapping)
        self.result: tuple[Any, ...] = ("skip", None, None)

        self.title("🧩 Cấu hình Ánh xạ Cột - Ghép & Chuẩn hóa")
        self.geometry("980x680")
        self.minsize(850, 500)
        self.configure(bg=self.palette.WINDOW_BG)

        # Chặn tương tác ngoài dialog (modal)
        self.transient(parent)
        self.grab_set()

        # Biến Profile Name
        clean_file_base = file_name.rsplit('.', 1)[0]
        if self.matched_profile_name:
            default_prof_name = f"Mẫu {self.matched_profile_name}"
        else:
            default_prof_name = f"Mẫu {clean_file_base}"
        self.profile_name_var = tk.StringVar(value=default_prof_name)
        self.saved_profile_var = tk.StringVar()

        # Biến chọn cho từng cột nguồn: index -> StringVar
        self.source_vars: dict[int, tk.StringVar] = {}
        self.target_vars = self.source_vars

        self._build_ui()
        self._init_suggestions()
        self._center_window(parent)

    def _center_window(self, parent):
        self.update_idletasks()
        pw = parent.winfo_width()
        ph = parent.winfo_height()
        px = parent.winfo_rootx()
        py = parent.winfo_rooty()
        w = self.winfo_width()
        h = self.winfo_height()
        x = max(0, px + (pw - w) // 2)
        y = max(0, py + (ph - h) // 2)
        self.geometry(f"{w}x{h}+{x}+{y}")

    def _build_ui(self):
        p = self.palette

        # 1. Header Banner
        hdr_frame = tk.Frame(self, bg=p.HEADER_BANNER_BG, padx=16, pady=10)
        hdr_frame.pack(fill=tk.X)

        tk.Label(
            hdr_frame,
            text="🧩 CẤU HÌNH ÁNH XẠ CỘT VÀO MẪU CHUẨN 28 CỘT",
            font=FONT_TITLE,
            bg=p.HEADER_BANNER_BG,
            fg=p.HEADER_BANNER_FG
        ).pack(anchor=tk.W)

        sub_text = f"Tệp nguồn: {self.file_name}  |  Sheet: '{self.sheet_name}'  |  Tìm thấy: {len(self.raw_headers)} cột"
        tk.Label(
            hdr_frame,
            text=sub_text,
            font=FONT_SUBTITLE,
            bg=p.HEADER_BANNER_BG,
            fg=p.HEADER_SUB_FG
        ).pack(anchor=tk.W, pady=(2, 0))

        # 2. Thanh Profile (Đặt tên & Nạp Profile cũ)
        prof_card = ttk.LabelFrame(self, text="Cấu hình Profile ánh xạ", padding=8)
        prof_card.pack(fill=tk.X, padx=12, pady=(8, 4))

        pr1 = ttk.Frame(prof_card)
        pr1.pack(fill=tk.X)

        ttk.Label(pr1, text="Tên Profile lưu lại:", font=FONT_LABEL_BOLD).pack(side=tk.LEFT, padx=(0, 6))
        prof_entry = ttk.Entry(pr1, textvariable=self.profile_name_var, font=FONT_LABEL, width=32)
        prof_entry.pack(side=tk.LEFT, padx=(0, 16))

        all_profiles = self.profile_mgr.get_all_profiles()
        prof_names = [f"{p_item.get('name', 'Profile')} ({p_item.get('id', '')})" for p_item in all_profiles]

        if prof_names:
            ttk.Label(pr1, text="Nạp từ Profile cũ:", font=FONT_LABEL).pack(side=tk.LEFT, padx=(0, 6))
            self.saved_prof_cb = ttk.Combobox(pr1, textvariable=self.saved_profile_var, values=prof_names, state="readonly", width=32)
            self.saved_prof_cb.pack(side=tk.LEFT, padx=(0, 8))
            self.saved_prof_cb.bind("<<ComboboxSelected>>", self._on_select_saved_profile)

        # 3. Thanh tác vụ nhanh
        quick_bar = ttk.Frame(self, padding=(12, 4))
        quick_bar.pack(fill=tk.X)

        self.show_only_extra = bool(self.extra_cols_with_data)
        if self.extra_cols_with_data:
            guide_lbl = f"⚠️ Kế thừa '{self.matched_profile_name or 'Profile'}': Đang chỉ hiển thị {len(self.extra_cols_with_data)} cột MỚI có dữ liệu. Vui lòng chọn ánh xạ hoặc bỏ qua."
            lbl_color = p.DANGER
        else:
            guide_lbl = "💡 Chọn cột đích tương ứng cho từng cột trong tệp Excel nguồn. Chọn '-- Bỏ qua --' cho các cột thừa."
            lbl_color = p.PRIMARY

        self.guide_label = ttk.Label(
            quick_bar,
            text=guide_lbl,
            font=FONT_MUTED,
            foreground=lbl_color
        )
        self.guide_label.pack(side=tk.LEFT)

        ttk.Button(quick_bar, text="🔄 Tự động gợi ý lại", command=self._on_auto_match).pack(side=tk.RIGHT, padx=(4, 0))
        ttk.Button(quick_bar, text="🧹 Bỏ qua tất cả", command=self._on_clear_all).pack(side=tk.RIGHT, padx=(4, 0))

        if self.extra_cols_with_data:
            self.btn_toggle_view = ttk.Button(
                quick_bar,
                text=f"📋 Xem tất cả ({len(self.raw_headers)} cột)",
                command=self._toggle_view_mode
            )
            self.btn_toggle_view.pack(side=tk.RIGHT, padx=(4, 0))

        # 4. Bảng cuộn danh sách N cột nguồn
        table_container = ttk.Frame(self, padding=(12, 4))
        table_container.pack(fill=tk.BOTH, expand=True)

        # Header của bảng
        t_header = tk.Frame(table_container, bg=p.TREE_HEADING_BG, padx=6, pady=6)
        t_header.pack(fill=tk.X)

        tk.Label(t_header, text="STT", font=FONT_LABEL_BOLD, bg=p.TREE_HEADING_BG, fg=p.TREE_HEADING_FG, width=4, anchor="center").pack(side=tk.LEFT)
        tk.Label(t_header, text="Cột trong tệp nguồn", font=FONT_LABEL_BOLD, bg=p.TREE_HEADING_BG, fg=p.TREE_HEADING_FG, width=24, anchor="w").pack(side=tk.LEFT, padx=(6, 0))
        tk.Label(t_header, text="Dữ liệu mẫu", font=FONT_LABEL_BOLD, bg=p.TREE_HEADING_BG, fg=p.TREE_HEADING_FG, width=28, anchor="w").pack(side=tk.LEFT, padx=(6, 0))
        tk.Label(t_header, text="👉 Ánh xạ sang Cột đích (Mẫu chuẩn 28)", font=FONT_LABEL_BOLD, bg=p.TREE_HEADING_BG, fg=p.TREE_HEADING_FG, anchor="w").pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(8, 0))

        # Canvas cuộn
        canvas_frame = ttk.Frame(table_container)
        canvas_frame.pack(fill=tk.BOTH, expand=True)

        self.canvas = tk.Canvas(canvas_frame, bg=p.WINDOW_BG, highlightthickness=0)
        scrollbar = ttk.Scrollbar(canvas_frame, orient=tk.VERTICAL, command=self.canvas.yview)
        self.scrollable_frame = tk.Frame(self.canvas, bg=p.WINDOW_BG)

        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )

        self.canvas_window = self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=scrollbar.set)

        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.canvas.bind("<Configure>", self._on_canvas_configure)
        self._bind_mousewheel(self.canvas)

        # Đổ N dòng cột nguồn vào scrollable_frame
        self.row_frames: dict[int, tk.Frame] = {}
        for i, header_name in enumerate(self.raw_headers):
            row_bg = p.CARD_BG if i % 2 == 0 else p.WINDOW_BG
            r_frame = tk.Frame(self.scrollable_frame, bg=row_bg, padx=6, pady=4)
            self.row_frames[i] = r_frame

            # Cột 1: STT
            tk.Label(r_frame, text=str(i + 1), font=FONT_LABEL, bg=row_bg, fg=p.TEXT, width=4, anchor="center").pack(side=tk.LEFT)

            # Cột 2: Tên cột nguồn
            is_extra_with_data = str(header_name) in self.extra_cols_with_data
            h_text = str(header_name or f"[Trống {i+1}]")
            h_fg = p.DANGER if is_extra_with_data else p.TEXT
            tk.Label(
                r_frame,
                text=h_text,
                font=FONT_LABEL_BOLD if is_extra_with_data else FONT_LABEL,
                bg=row_bg,
                fg=h_fg,
                width=24,
                anchor="w"
            ).pack(side=tk.LEFT, padx=(6, 0))

            # Cột 3: Dữ liệu mẫu (Sample value)
            sample_val = self.sample_values.get(i, "")
            sample_display = f'"{sample_val}"' if sample_val else "--- (ô rỗng) ---"
            if len(sample_display) > 28:
                sample_display = sample_display[:25] + "..."
            tk.Label(
                r_frame,
                text=sample_display,
                font=FONT_MUTED,
                bg=row_bg,
                fg=p.MUTED,
                width=28,
                anchor="w"
            ).pack(side=tk.LEFT, padx=(6, 0))

            # Cột 4: Dropdown chọn Cột đích
            var = tk.StringVar(value=IGNORE_OPTION)
            self.source_vars[i] = var
            cb = ttk.Combobox(r_frame, textvariable=var, values=TARGET_OPTIONS, state="readonly", width=36)
            cb.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(8, 0))

        # Hiển thị các dòng ban đầu theo chế độ lọc
        self._apply_row_visibility()

        # 5. Thanh nút bấm điều khiển ở đáy
        btn_bar = ttk.Frame(self, padding=(12, 10))
        btn_bar.pack(fill=tk.X)

        # Nút Dừng toàn bộ (Danger)
        ttk.Button(btn_bar, text="⏹️ Dừng toàn bộ", command=self._on_cancel).pack(side=tk.LEFT)

        # Nút Bỏ qua sheet này
        ttk.Button(btn_bar, text="⏭️ Bỏ qua Sheet này", command=self._on_skip).pack(side=tk.LEFT, padx=(8, 0))

        # Nút Lưu & Áp dụng
        save_btn = ttk.Button(btn_bar, text="✅ Lưu Profile & Áp dụng", command=self._on_save_and_apply)
        save_btn.pack(side=tk.RIGHT)
        save_btn.focus_set()

    def _apply_row_visibility(self):
        """Ẩn hoặc hiện các dòng theo chế độ lọc (chỉ hiện cột mới có dữ liệu hoặc hiện tất cả)."""
        for i, header_name in enumerate(self.raw_headers):
            r_frame = self.row_frames.get(i)
            if not r_frame:
                continue
            if self.show_only_extra and str(header_name) not in self.extra_cols_with_data:
                r_frame.pack_forget()
            else:
                r_frame.pack(fill=tk.X, expand=True)

    def _toggle_view_mode(self):
        """Chuyển đổi giữa chế độ chỉ xem cột mới có dữ liệu và xem tất cả các cột."""
        self.show_only_extra = not self.show_only_extra
        if hasattr(self, "btn_toggle_view"):
            if self.show_only_extra:
                self.btn_toggle_view.config(text=f"📋 Xem tất cả ({len(self.raw_headers)} cột)")
                if hasattr(self, "guide_label"):
                    self.guide_label.config(
                        text=f"⚠️ Kế thừa '{self.matched_profile_name or 'Profile'}': Đang chỉ hiển thị {len(self.extra_cols_with_data)} cột MỚI có dữ liệu.",
                        foreground=self.palette.DANGER
                    )
            else:
                self.btn_toggle_view.config(text=f"🔍 Chỉ xem cột mới có dữ liệu ({len(self.extra_cols_with_data)})")
                if hasattr(self, "guide_label"):
                    self.guide_label.config(
                        text=f"📋 Đang hiển thị toàn bộ {len(self.raw_headers)} cột. Các cột màu đỏ là cột mới có dữ liệu.",
                        foreground=self.palette.PRIMARY
                    )
        self._apply_row_visibility()

    def _on_canvas_configure(self, event):
        self.canvas.itemconfig(self.canvas_window, width=event.width)

    def _bind_mousewheel(self, widget):
        def _on_mousewheel(event):
            delta = int(-1 * (event.delta / 120)) if event.delta else 1
            self.canvas.yview_scroll(delta, "units")
        widget.bind_all("<MouseWheel>", _on_mousewheel)

    def _init_suggestions(self):
        """Khởi tạo giá trị gợi ý ban đầu vào các dropdown."""
        if self.suggested_inverted_mapping:
            for i, h in enumerate(self.raw_headers):
                if h is None:
                    continue
                nh = normalize_text_for_match(str(h))
                spec = self.suggested_inverted_mapping.get(nh)
                opt = target_spec_to_option(spec)
                if i in self.source_vars:
                    self.source_vars[i].set(opt)
        elif self.suggested_mapping:
            # Chuyển đổi từ execution_mapping sang dropdown
            for t_idx, s_val in self.suggested_mapping.items():
                if s_val is None:
                    continue
                if isinstance(s_val, list):
                    for ord_i, si in enumerate(s_val):
                        if si in self.source_vars:
                            opt = target_spec_to_option([t_idx, ord_i])
                            self.source_vars[si].set(opt)
                elif isinstance(s_val, int) and s_val in self.source_vars:
                    opt = target_spec_to_option(t_idx)
                    self.source_vars[s_val].set(opt)

    def _on_auto_match(self):
        """Tự động khớp mờ lại từ đầu."""
        new_inverted = self.profile_mgr.auto_suggest_inverted_mapping(self.raw_headers)
        self.suggested_inverted_mapping = new_inverted
        self._init_suggestions()

    def _on_clear_all(self):
        """Đặt tất cả về -- Bỏ qua --."""
        for var in self.source_vars.values():
            var.set(IGNORE_OPTION)

    def _on_select_saved_profile(self, event):
        """Nạp cấu hình từ Profile cũ."""
        selected_text = self.saved_profile_var.get()
        if not selected_text:
            return
        all_profiles = self.profile_mgr.get_all_profiles()
        for p_item in all_profiles:
            match_str = f"{p_item.get('name', 'Profile')} ({p_item.get('id', '')})"
            if match_str == selected_text:
                p_map = p_item.get("mapping", {})
                for i, h in enumerate(self.raw_headers):
                    if h is None:
                        continue
                    nh = normalize_text_for_match(str(h))
                    spec = p_map.get(nh)
                    opt = target_spec_to_option(spec)
                    if i in self.source_vars:
                        self.source_vars[i].set(opt)
                self.profile_name_var.set(p_item.get("name", self.profile_name_var.get()))
                break

    def _get_inverted_mapping(self) -> dict[str, Any]:
        """Lấy mapping đảo ngược từ giao diện: {norm_header: target_spec}."""
        res: dict[str, Any] = {}
        for i, h in enumerate(self.raw_headers):
            if h is None or not str(h).strip():
                continue
            nh = normalize_text_for_match(str(h))
            opt = self.source_vars[i].get()
            spec = option_to_target_spec(opt)
            res[nh] = spec
        return res

    def _on_save_and_apply(self):
        """Xác nhận cấu hình ánh xạ và lưu Profile."""
        inverted_map = self._get_inverted_mapping()
        execution_map = self.profile_mgr.resolve_execution_mapping(self.raw_headers, inverted_map)

        # Kiểm tra cột bắt buộc Họ và tên (Cột 5)
        if execution_map.get(5) is None:
            confirm = messagebox.askyesno(
                "Cảnh báo thiếu Họ và tên",
                "Bạn chưa chọn cột nào tương ứng với 'Họ, chữ đệm và tên' (Cột 5).\n\n"
                "Nếu không có cột Họ tên, hệ thống không thể phân biệt dòng học sinh hợp lệ.\n"
                "Bạn có chắc chắn muốn tiếp tục không?",
                parent=self
            )
            if not confirm:
                return

        prof_name = self.profile_name_var.get().strip() or f"Mẫu {self.file_name}"
        self.result = ("apply", execution_map, prof_name, inverted_map)
        self.grab_release()
        self.destroy()

    def _on_skip(self):
        """Bỏ qua sheet này."""
        self.result = ("skip", None, None)
        self.grab_release()
        self.destroy()

    def _on_cancel(self):
        """Dừng toàn bộ quá trình."""
        confirm = messagebox.askyesno(
            "Xác nhận dừng toàn bộ",
            "Bạn có chắc chắn muốn DỪNG TOÀN BỘ quá trình ghép không?\n"
            "Tất cả các tệp và sheet còn lại sẽ dừng lại ngay lập tức.",
            parent=self
        )
        if confirm:
            self.result = ("cancel", None, None)
            self.grab_release()
            self.destroy()
