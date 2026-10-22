from __future__ import annotations

import sys
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
    auto_suggest_mapping, auto_suggest_mapping_with_details,
    parse_column_expression, normalize_text_for_match, compute_similarity
)

EMPTY_OPTION = "-- Để trống --"


class ColumnMappingDialog(tk.Toplevel):
    """
    Hộp thoại Cấu hình Ánh xạ Cột thông minh (Ghép & Chuẩn hóa):
    - Hiển thị 28 cột chuẩn của mẫu Mau_DS_Cap_Bang_THCS.xlsx.
    - Điền sẵn gợi ý qua 2 tầng: Profile tương đồng trước -> Từ khóa & Khớp mờ sau.
    - Cho phép người dùng kiểm tra, chỉnh sửa từng cột bằng Combobox hoặc chọn để trống.
    - Cho phép đặt tên Profile và lưu lại theo tên cột để tự động hóa an toàn cho các lần sau.
    """

    def __init__(
        self,
        parent: tk.Tk | tk.Toplevel,
        file_name: str,
        sheet_name: str,
        raw_headers: list[str],
        formatted_headers: list[str],
        suggested_mapping: dict[int, int | None],
        profile_mgr: MappingProfileManager | None = None,
        is_dark: bool | None = None,
        suggested_defaults: dict[int, str] | None = None,
        matched_profile_name: str | None = None
    ):
        super().__init__(parent)
        self.is_dark = IS_DARK if is_dark is None else is_dark
        self.palette = get_theme_palette(self.is_dark)
        self.profile_mgr = profile_mgr or MappingProfileManager()

        self.file_name = file_name
        self.sheet_name = sheet_name
        self.raw_headers = raw_headers
        self.formatted_headers = formatted_headers
        self.suggested_mapping = suggested_mapping
        self.suggested_defaults = suggested_defaults or {}
        self.matched_profile_name = matched_profile_name

        # Kết quả trả về: (mapping, profile_name, default_values, should_skip)
        self.result: tuple[dict[int, int | list[int] | None] | None, str | None, dict[int, str], bool] = (None, None, {}, True)

        self.title("🧩 Cấu hình Ánh xạ Cột - Ghép & Chuẩn hóa")
        self.geometry("1020x720")
        self.minsize(900, 580)
        self.configure(bg=self.palette.WINDOW_BG)

        # Chặn tương tác ngoài dialog (modal)
        self.transient(parent)
        self.grab_set()

        # Biến giao diện
        if self.matched_profile_name:
            default_prof_name = f"Mẫu {self.matched_profile_name} ({len(raw_headers)} cột)"
        else:
            default_prof_name = f"Mẫu {file_name.rsplit('.', 1)[0]} ({len(raw_headers)} cột)"
        self.profile_name_var = tk.StringVar(value=default_prof_name)
        self.saved_profile_var = tk.StringVar()

        # Danh sách combobox ánh xạ và ô nhập giá trị mặc định cho 28 cột
        self.target_vars: dict[int, tk.StringVar] = {}
        self.default_vars: dict[int, tk.StringVar] = {}
        self.combo_options = [EMPTY_OPTION] + formatted_headers

        self._build_ui()
        self._apply_mapping(self.suggested_mapping)
        if self.suggested_defaults:
            self._apply_defaults(self.suggested_defaults)
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

        title_lbl = tk.Label(
            hdr_frame,
            text="🧩 CẤU HÌNH ÁNH XẠ CỘT VÀO MẪU CHUẨN 28 CỘT",
            font=FONT_TITLE,
            bg=p.HEADER_BANNER_BG,
            fg=p.HEADER_BANNER_FG
        )
        title_lbl.pack(anchor=tk.W)

        sub_text = f"Tệp nguồn: {self.file_name}  |  Sheet: '{self.sheet_name}'  |  Tìm thấy: {len(self.raw_headers)} cột"
        sub_lbl = tk.Label(
            hdr_frame,
            text=sub_text,
            font=FONT_SUBTITLE,
            bg=p.HEADER_BANNER_BG,
            fg=p.HEADER_SUB_FG
        )
        sub_lbl.pack(anchor=tk.W, pady=(2, 0))

        # 2. Profile Bar (Lưu & Nạp Profile)
        prof_card = ttk.LabelFrame(self, text="Cấu hình nhận diện Profile", padding=8)
        prof_card.pack(fill=tk.X, padx=12, pady=(8, 4))

        pr1 = ttk.Frame(prof_card)
        pr1.pack(fill=tk.X)

        ttk.Label(pr1, text="Tên Profile lưu lại:", font=FONT_LABEL_BOLD).pack(side=tk.LEFT, padx=(0, 6))
        prof_entry = ttk.Entry(pr1, textvariable=self.profile_name_var, font=FONT_LABEL, width=32)
        prof_entry.pack(side=tk.LEFT, padx=(0, 16))

        # Danh sách profile đã lưu
        all_profiles = self.profile_mgr.list_profiles()
        prof_names = [f"{p_item.get('name', 'Profile')} ({p_item.get('id', '')})" for p_item in all_profiles]

        if prof_names:
            ttk.Label(pr1, text="Nạp từ Profile cũ:", font=FONT_LABEL).pack(side=tk.LEFT, padx=(0, 6))
            self.saved_prof_cb = ttk.Combobox(pr1, textvariable=self.saved_profile_var, values=prof_names, state="readonly", width=30)
            self.saved_prof_cb.pack(side=tk.LEFT, padx=(0, 8))
            self.saved_prof_cb.bind("<<ComboboxSelected>>", self._on_select_saved_profile)

        # 3. Thanh tác vụ nhanh
        quick_bar = ttk.Frame(self, padding=(12, 4))
        quick_bar.pack(fill=tk.X)

        ttk.Label(
            quick_bar,
            text="💡 Hỗ trợ gộp nhiều cột nguồn bằng dấu '+' (ví dụ: 'B + C', 'Cột B + Cột C') hoặc bấm nút '➕' bên cạnh để tích chọn các cột cần ghép.",
            font=FONT_MUTED,
            foreground=p.PRIMARY
        ).pack(side=tk.LEFT)

        ttk.Button(quick_bar, text="🔄 Tự động khớp lại", command=self._on_auto_match).pack(side=tk.RIGHT, padx=(4, 0))
        ttk.Button(quick_bar, text="🧹 Để trống tất cả", command=self._on_clear_all).pack(side=tk.RIGHT)

        # 4. Bảng ánh xạ cuộn (Scrollable Canvas)
        table_container = ttk.Frame(self, padding=(12, 4))
        table_container.pack(fill=tk.BOTH, expand=True)

        # Tiêu đề bảng
        t_header = tk.Frame(table_container, bg=p.TREE_HEADING_BG, padx=6, pady=6)
        t_header.pack(fill=tk.X)

        tk.Label(t_header, text="STT", font=FONT_LABEL_BOLD, bg=p.TREE_HEADING_BG, fg=p.TREE_HEADING_FG, width=4, anchor="center").pack(side=tk.LEFT)
        tk.Label(t_header, text="Cột đích (Mẫu chuẩn 28)", font=FONT_LABEL_BOLD, bg=p.TREE_HEADING_BG, fg=p.TREE_HEADING_FG, width=24, anchor="w").pack(side=tk.LEFT, padx=(6, 0))
        tk.Label(t_header, text="Yêu cầu", font=FONT_LABEL_BOLD, bg=p.TREE_HEADING_BG, fg=p.TREE_HEADING_FG, width=10, anchor="center").pack(side=tk.LEFT, padx=(6, 0))
        tk.Label(t_header, text="👉 Cột nguồn tương ứng", font=FONT_LABEL_BOLD, bg=p.TREE_HEADING_BG, fg=p.TREE_HEADING_FG, width=28, anchor="w").pack(side=tk.LEFT, padx=(6, 0))
        tk.Label(t_header, text="✍️ Giá trị mặc định (nếu ô nguồn trống)", font=FONT_LABEL_BOLD, bg=p.TREE_HEADING_BG, fg=p.TREE_HEADING_FG, anchor="w").pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(8, 0))

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

        # Đổ 28 dòng cột vào scrollable_frame
        for col_def in TARGET_COLUMNS:
            t_idx = col_def["idx"]
            t_name = col_def["name"]
            t_req = col_def["required"]
            t_def = col_def["default"]

            row_bg = p.CARD_BG if t_idx % 2 == 0 else p.WINDOW_BG
            r_frame = tk.Frame(self.scrollable_frame, bg=row_bg, padx=6, pady=4)
            r_frame.pack(fill=tk.X, expand=True)

            # Cột 1: STT
            tk.Label(r_frame, text=str(t_idx), font=FONT_LABEL, bg=row_bg, fg=p.TEXT, width=4, anchor="center").pack(side=tk.LEFT)

            # Cột 2: Tên cột đích
            tk.Label(r_frame, text=t_name, font=FONT_LABEL_BOLD if t_req else FONT_LABEL, bg=row_bg, fg=p.TEXT, width=24, anchor="w").pack(side=tk.LEFT, padx=(6, 0))

            # Cột 3: Tính chất
            req_text = "⭐ Bắt buộc" if t_req else ("Khuyên dùng" if t_idx in (9, 14, 15) else "Tùy chọn")
            req_fg = p.DANGER if t_req else (p.WARNING if t_idx in (9, 14, 15) else p.MUTED)
            tk.Label(r_frame, text=req_text, font=FONT_MUTED, bg=row_bg, fg=req_fg, width=10, anchor="center").pack(side=tk.LEFT, padx=(6, 0))

            # Cột 4: Combobox Cột nguồn + Nút ➕ Ghép nhiều cột
            c_var = tk.StringVar(value=EMPTY_OPTION)
            self.target_vars[t_idx] = c_var

            cb = ttk.Combobox(r_frame, textvariable=c_var, values=self.combo_options, state="normal", width=22)
            cb.pack(side=tk.LEFT, padx=(6, 0))

            btn_multi = ttk.Button(
                r_frame,
                text="➕",
                width=3,
                command=lambda idx=t_idx, name=t_name: self._open_multi_column_selector(idx, name)
            )
            btn_multi.pack(side=tk.LEFT, padx=(2, 0))

            # Cột 5: Ô nhập Giá trị mặc định nếu ô nguồn trống
            d_var = tk.StringVar()
            if t_idx == 4:
                d_var.set("Bằng tốt nghiệp trung học cơ sở")
            elif t_idx == 12:
                d_var.set("Việt Nam")
            self.default_vars[t_idx] = d_var

            if t_idx == 1:
                tk.Label(r_frame, text="Tự động đánh số 1..N", font=FONT_MUTED, bg=row_bg, fg=p.MUTED, anchor="w").pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(8, 0))
            else:
                def_entry = ttk.Entry(r_frame, textvariable=d_var, font=FONT_LABEL)
                def_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(8, 0))

        # 5. Thanh nút bấm điều khiển ở đáy
        btn_bar = ttk.Frame(self, padding=(12, 10))
        btn_bar.pack(fill=tk.X)

        ttk.Button(btn_bar, text="⏹️ Hủy bỏ toàn bộ", command=self._on_cancel).pack(side=tk.LEFT)
        ttk.Button(btn_bar, text="⏭️ Bỏ qua Sheet này", command=self._on_skip).pack(side=tk.LEFT, padx=(8, 0))

        save_btn = ttk.Button(btn_bar, text="✅ Xác nhận & Lưu Profile", command=self._on_save_and_confirm)
        save_btn.pack(side=tk.RIGHT)
        save_btn.focus_set()

    def _on_canvas_configure(self, event):
        self.canvas.itemconfig(self.canvas_window, width=event.width)

    def _bind_mousewheel(self, widget):
        def _on_mousewheel(event):
            if sys.platform == "darwin":
                self.canvas.yview_scroll(int(-1 * event.delta), "units")
            else:
                self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        widget.bind_all("<MouseWheel>", _on_mousewheel)

    def _open_multi_column_selector(self, t_idx: int, t_name: str):
        c_var = self.target_vars[t_idx]
        MultiColumnSelectorDialog(
            parent=self,
            target_name=t_name,
            formatted_headers=self.formatted_headers,
            current_value=c_var.get(),
            callback=lambda res_str: c_var.set(res_str)
        )

    def _apply_mapping(self, mapping: dict[int, int | list[int] | None]):
        """Cập nhật giá trị hiển thị trên combobox theo mapping."""
        for t_idx, s_val in mapping.items():
            if t_idx not in self.target_vars:
                continue
            if isinstance(s_val, list):
                # Nhiều cột gộp: e.g. [1, 2] -> "Cột B: Họ và chữ đệm + Cột C: Tên"
                text_parts = []
                for sc in s_val:
                    if sc is not None and 0 <= sc < len(self.formatted_headers):
                        text_parts.append(self.formatted_headers[sc])
                if text_parts:
                    self.target_vars[t_idx].set(" + ".join(text_parts))
                else:
                    self.target_vars[t_idx].set(EMPTY_OPTION)
            elif isinstance(s_val, int) and 0 <= s_val < len(self.formatted_headers):
                self.target_vars[t_idx].set(self.formatted_headers[s_val])
            else:
                self.target_vars[t_idx].set(EMPTY_OPTION)

    def _get_current_mapping(self) -> dict[int, int | list[int] | None]:
        """Lấy bản đồ ánh xạ hiện tại từ các combobox hoặc biểu thức người dùng gõ."""
        mapping: dict[int, int | list[int] | None] = {}
        for t_idx, c_var in self.target_vars.items():
            val = c_var.get()
            mapping[t_idx] = parse_column_expression(val, self.formatted_headers)
        return mapping

    def _apply_defaults(self, defaults: dict[int | str, str]):
        """Cập nhật các ô nhập giá trị mặc định từ profile."""
        for t_idx, d_var in self.default_vars.items():
            str_key = str(t_idx)
            int_key = t_idx
            if str_key in defaults:
                d_var.set(defaults[str_key])
            elif int_key in defaults:
                d_var.set(defaults[int_key])
            else:
                if t_idx == 4:
                    d_var.set("Bằng tốt nghiệp trung học cơ sở")
                elif t_idx == 12:
                    d_var.set("Việt Nam")
                else:
                    d_var.set("")

    def _find_best_header_match(self, name: str) -> int | None:
        """Tìm cột trong raw_headers khớp chính xác hoặc khớp mờ nhất với name."""
        if not name:
            return None
        norm_target = normalize_text_for_match(name)
        if not norm_target:
            return None
        # 1. Khớp chính xác
        for idx, rh in enumerate(self.raw_headers):
            if normalize_text_for_match(rh) == norm_target:
                return idx
        # 2. Khớp mờ
        best_idx = None
        best_sc = 0.0
        for idx, rh in enumerate(self.raw_headers):
            sc, _ = compute_similarity(name, [name], rh)
            if sc > best_sc:
                best_sc = sc
                best_idx = idx
        if best_idx is not None and best_sc >= 0.65:
            return best_idx
        return None

    def _on_auto_match(self):
        """Khớp lại tự động bằng thuật toán 2 tầng (Profile tương đồng trước, Từ khóa sau)."""
        new_map, new_defs, matched_pname = auto_suggest_mapping_with_details(self.raw_headers, self.profile_mgr)
        self._apply_mapping(new_map)
        if new_defs:
            self._apply_defaults(new_defs)
        if matched_pname:
            self.profile_name_var.set(f"Theo {matched_pname} ({len(self.raw_headers)} cột)")

    def _on_clear_all(self):
        """Đặt tất cả về Để trống."""
        for c_var in self.target_vars.values():
            c_var.set(EMPTY_OPTION)

    def _on_select_saved_profile(self, event):
        """Khi người dùng chọn một Profile đã lưu: Ánh xạ linh hoạt theo TÊN CỘT thay vì index."""
        selected_text = self.saved_profile_var.get()
        if not selected_text:
            return
        all_profiles = self.profile_mgr.list_profiles()
        for p_item in all_profiles:
            match_str = f"{p_item.get('name', 'Profile')} ({p_item.get('id', '')})"
            if match_str == selected_text:
                named_map = self.profile_mgr.get_named_mapping(p_item)
                # Chuyển đổi named_map sang index thực tế của file hiện tại
                resolved_map: dict[int, int | list[int] | None] = {}
                for t_idx_str, col_val in named_map.items():
                    t_idx = int(t_idx_str)
                    if col_val is None:
                        resolved_map[t_idx] = None
                    elif isinstance(col_val, list):
                        sub_indices = []
                        for sub_name in col_val:
                            f_idx = self._find_best_header_match(sub_name)
                            if f_idx is not None and f_idx not in sub_indices:
                                sub_indices.append(f_idx)
                        resolved_map[t_idx] = sub_indices if sub_indices else None
                    else:
                        resolved_map[t_idx] = self._find_best_header_match(col_val)

                self._apply_mapping(resolved_map)

                saved_defaults = p_item.get("default_values", {})
                self._apply_defaults(saved_defaults)

                self.profile_name_var.set(p_item.get("name", self.profile_name_var.get()))
                break

    def _on_save_and_confirm(self):
        """Xác nhận cấu hình ánh xạ và lưu profile."""
        current_map = self._get_current_mapping()
        current_defaults = {t_idx: var.get().strip() for t_idx, var in self.default_vars.items() if var.get().strip()}

        # Kiểm tra cột bắt buộc Họ và tên (Cột 5)
        if current_map.get(5) is None:
            confirm = messagebox.askyesno(
                "Cảnh báo thiếu Họ và tên",
                "Bạn chưa chọn cột nguồn cho 'Họ, chữ đệm và tên' (Cột 5).\n\n"
                "Nếu không có cột Họ tên, hệ thống có thể không phân biệt được dòng học sinh.\n"
                "Bạn có chắc chắn muốn tiếp tục không?",
                parent=self
            )
            if not confirm:
                return

        prof_name = self.profile_name_var.get().strip() or f"Mẫu {self.file_name}"
        self.result = (current_map, prof_name, current_defaults, False)
        self.grab_release()
        self.destroy()

    def _on_skip(self):
        """Bỏ qua sheet này."""
        self.result = (None, None, {}, True)
        self.grab_release()
        self.destroy()

    def _on_cancel(self):
        """Hủy bỏ toàn bộ quá trình."""
        confirm = messagebox.askyesno(
            "Xác nhận hủy",
            "Bạn có chắc chắn muốn hủy bỏ toàn bộ quá trình ghép không?",
            parent=self
        )
        if confirm:
            self.result = (None, None, {}, False)
            self.grab_release()
            self.destroy()


class MultiColumnSelectorDialog(tk.Toplevel):
    """Hộp thoại tích chọn nhiều cột nguồn để ghép vào một cột đích."""

    def __init__(
        self,
        parent: tk.Toplevel,
        target_name: str,
        formatted_headers: list[str],
        current_value: str,
        callback: Any
    ):
        super().__init__(parent)
        self.callback = callback
        self.formatted_headers = formatted_headers
        is_dark = getattr(parent, 'is_dark', IS_DARK)
        self.palette = get_theme_palette(is_dark)

        self.title(f"🧩 Ghép nhiều cột: {target_name}")
        self.geometry("520x460")
        self.minsize(440, 340)
        self.configure(bg=self.palette.WINDOW_BG)
        self.transient(parent)
        self.grab_set()

        p = self.palette

        # Parse current_value để tìm các chỉ số cột đang được chọn trước đó
        parsed = parse_column_expression(current_value, formatted_headers)
        if isinstance(parsed, list):
            selected_indices = set(parsed)
        elif isinstance(parsed, int):
            selected_indices = {parsed}
        else:
            selected_indices = set()

        self.check_vars: list[tuple[int, tk.BooleanVar, str]] = []

        # Header info
        hdr_frame = tk.Frame(self, bg=p.HEADER_BANNER_BG, padx=14, pady=10)
        hdr_frame.pack(fill=tk.X)

        tk.Label(
            hdr_frame,
            text=f"Tích chọn các cột nguồn cần gộp vào: {target_name}",
            font=FONT_LABEL_BOLD,
            bg=p.HEADER_BANNER_BG,
            fg=p.HEADER_BANNER_FG
        ).pack(anchor=tk.W)

        tk.Label(
            hdr_frame,
            text="Hệ thống sẽ tự động ghép các cột được chọn lại với nhau (ngăn cách bằng 1 dấu cách).",
            font=FONT_MUTED,
            bg=p.HEADER_BANNER_BG,
            fg=p.HEADER_SUB_FG
        ).pack(anchor=tk.W, pady=(2, 0))

        # Scrollable checklist
        container = ttk.Frame(self, padding=(12, 8))
        container.pack(fill=tk.BOTH, expand=True)

        canvas = tk.Canvas(container, bg=p.WINDOW_BG, highlightthickness=0)
        scrollbar = ttk.Scrollbar(container, orient=tk.VERTICAL, command=canvas.yview)
        scroll_frame = tk.Frame(canvas, bg=p.WINDOW_BG)

        scroll_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        c_win = canvas.create_window((0, 0), window=scroll_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.bind("<Configure>", lambda e: canvas.itemconfig(c_win, width=e.width))

        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        for i, fh in enumerate(formatted_headers):
            var = tk.BooleanVar(value=(i in selected_indices))
            self.check_vars.append((i, var, fh))
            chk = ttk.Checkbutton(scroll_frame, text=fh, variable=var)
            chk.pack(anchor=tk.W, pady=3, padx=6)

        # Buttons
        btn_frame = ttk.Frame(self, padding=12)
        btn_frame.pack(fill=tk.X)

        ttk.Button(btn_frame, text="Hủy", command=self._on_close).pack(side=tk.LEFT)
        ttk.Button(btn_frame, text="🧹 Bỏ chọn hết", command=self._clear_all).pack(side=tk.LEFT, padx=(8, 0))
        ttk.Button(btn_frame, text="✅ Áp dụng ghép", command=self._confirm).pack(side=tk.RIGHT)

        self._center_window(parent)

    def _clear_all(self):
        for _, var, _ in self.check_vars:
            var.set(False)

    def _confirm(self):
        selected_headers = [fh for _, var, fh in self.check_vars if var.get()]
        if selected_headers:
            result_str = " + ".join(selected_headers)
        else:
            result_str = EMPTY_OPTION
        self.callback(result_str)
        self._on_close()

    def _on_close(self):
        self.grab_release()
        self.destroy()

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

