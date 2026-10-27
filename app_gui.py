# -*- coding: utf-8 -*-
"""
VietExcel FontFixer - GUI Application
Modern Desktop Application for fixing VNI and TCVN3 fonts in Excel files (.xls, .xlsx)
"""

import os
import sys
import threading
import tkinter as tk
from tkinter import filedialog, messagebox
import customtkinter as ctk

from excel_processor import convert_file

ctk.set_appearance_mode("Dark")  # Modes: "System", "Dark", "Light"
ctk.set_default_color_theme("blue")  # Themes: "blue", "green", "dark-blue"


class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("VietExcel FontFixer - Sửa Lỗi Font VNI & TCVN3 Cho Excel")
        self.geometry("900, 680")
        self.minsize(800, 580)

        # File list: list of dicts: {"path": str, "status": str, "output": str}
        self.file_items = []
        self.is_processing = False

        self._build_ui()

    def _build_ui(self):
        # Configure grid
        self.grid_rowconfigure(2, weight=1)
        self.grid_columnconfigure(0, weight=1)

        # 1. Header Frame
        header_frame = ctk.CTkFrame(self, corner_radius=10)
        header_frame.grid(row=0, column=0, padx=15, pady=(15, 10), sticky="ew")

        title_label = ctk.CTkLabel(
            header_frame,
            text="✨ VietExcel FontFixer",
            font=ctk.CTkFont(size=22, weight="bold")
        )
        title_label.pack(anchor="w", padx=15, pady=(10, 2))

        desc_label = ctk.CTkLabel(
            header_frame,
            text="Tự động sửa lỗi font VNI-Windows (VNI-Times,...) & TCVN3 (.VnTime,...) sang Unicode chuẩn • Bảo toàn 100% định dạng, viền ô, căn lề",
            font=ctk.CTkFont(size=13),
            text_color="gray70"
        )
        desc_label.pack(anchor="w", padx=15, pady=(0, 10))

        # 2. Options Frame
        opts_frame = ctk.CTkFrame(self, corner_radius=10)
        opts_frame.grid(row=1, column=0, padx=15, pady=(0, 10), sticky="ew")
        opts_frame.grid_columnconfigure(1, weight=1)

        # Mode Selection
        mode_label = ctk.CTkLabel(opts_frame, text="Bảng mã nguồn:", font=ctk.CTkFont(weight="bold"))
        mode_label.grid(row=0, column=0, padx=15, pady=10, sticky="w")

        self.mode_var = ctk.StringVar(value="Tự động nhận diện")
        self.mode_menu = ctk.CTkSegmentedButton(
            opts_frame,
            values=["Tự động nhận diện", "VNI-Windows", "TCVN3 (ABC)"],
            variable=self.mode_var
        )
        self.mode_menu.grid(row=0, column=1, padx=15, pady=10, sticky="w")

        # Output Folder Option
        out_label = ctk.CTkLabel(opts_frame, text="Thư mục lưu:", font=ctk.CTkFont(weight="bold"))
        out_label.grid(row=1, column=0, padx=15, pady=(0, 10), sticky="w")

        out_row = ctk.CTkFrame(opts_frame, fg_color="transparent")
        out_row.grid(row=1, column=1, padx=15, pady=(0, 10), sticky="ew")
        out_row.grid_columnconfigure(1, weight=1)

        self.out_mode_var = ctk.StringVar(value="same")
        self.radio_same = ctk.CTkRadioButton(
            out_row,
            text="Cùng thư mục file gốc (_fixed.xlsx)",
            variable=self.out_mode_var,
            value="same",
            command=self._on_out_mode_changed
        )
        self.radio_same.grid(row=0, column=0, padx=(0, 15), pady=0, sticky="w")

        self.radio_custom = ctk.CTkRadioButton(
            out_row,
            text="Thư mục tùy chọn:",
            variable=self.out_mode_var,
            value="custom",
            command=self._on_out_mode_changed
        )
        self.radio_custom.grid(row=0, column=1, padx=(0, 10), pady=0, sticky="w")

        self.custom_folder_path = ctk.StringVar(value="")
        self.entry_out_dir = ctk.CTkEntry(out_row, textvariable=self.custom_folder_path, state="disabled", width=220)
        self.entry_out_dir.grid(row=0, column=2, padx=(0, 10), pady=0, sticky="ew")

        self.btn_browse_out = ctk.CTkButton(
            out_row,
            text="Chọn...",
            width=65,
            state="disabled",
            command=self._choose_output_dir
        )
        self.btn_browse_out.grid(row=0, column=3, padx=0, pady=0)

        # 3. Main Center Area: Files list & Log
        main_split = ctk.CTkFrame(self, corner_radius=10)
        main_split.grid(row=2, column=0, padx=15, pady=(0, 10), sticky="nsew")
        main_split.grid_rowconfigure(1, weight=1)
        main_split.grid_columnconfigure(0, weight=3)
        main_split.grid_columnconfigure(1, weight=2)

        # Left: File List
        list_header = ctk.CTkFrame(main_split, fg_color="transparent")
        list_header.grid(row=0, column=0, padx=10, pady=(10, 5), sticky="ew")
        self.lbl_file_count = ctk.CTkLabel(list_header, text="Danh sách file (0 file)", font=ctk.CTkFont(weight="bold"))
        self.lbl_file_count.pack(side="left")

        self.file_scroll_frame = ctk.CTkScrollableFrame(main_split, label_text="Tên File • Trạng thái")
        self.file_scroll_frame.grid(row=1, column=0, padx=10, pady=(0, 10), sticky="nsew")

        # Right: Log Box
        log_header = ctk.CTkFrame(main_split, fg_color="transparent")
        log_header.grid(row=0, column=1, padx=10, pady=(10, 5), sticky="ew")
        lbl_log = ctk.CTkLabel(log_header, text="Nhật ký chuyển đổi", font=ctk.CTkFont(weight="bold"))
        lbl_log.pack(side="left")

        self.log_textbox = ctk.CTkTextbox(main_split, font=ctk.CTkFont(family="Consolas", size=12))
        self.log_textbox.grid(row=1, column=1, padx=10, pady=(0, 10), sticky="nsew")
        self.log_textbox.insert("end", "Sẵn sàng. Vui lòng thêm file hoặc thư mục để bắt đầu.\n")

        # 4. Progress and Status Bar
        status_frame = ctk.CTkFrame(self, corner_radius=10)
        status_frame.grid(row=3, column=0, padx=15, pady=(0, 10), sticky="ew")
        status_frame.grid_columnconfigure(0, weight=1)

        self.progress_bar = ctk.CTkProgressBar(status_frame)
        self.progress_bar.grid(row=0, column=0, padx=15, pady=(10, 5), sticky="ew")
        self.progress_bar.set(0)

        self.status_lbl = ctk.CTkLabel(status_frame, text="Chờ thao tác...", font=ctk.CTkFont(size=12), text_color="gray70")
        self.status_lbl.grid(row=1, column=0, padx=15, pady=(0, 10), sticky="w")

        # 5. Buttons Control Bar
        btn_frame = ctk.CTkFrame(self, corner_radius=10, fg_color="transparent")
        btn_frame.grid(row=4, column=0, padx=15, pady=(0, 15), sticky="ew")

        self.btn_add_files = ctk.CTkButton(
            btn_frame,
            text="➕ Thêm Files",
            width=120,
            command=self._add_files
        )
        self.btn_add_files.pack(side="left", padx=(0, 10))

        self.btn_add_folder = ctk.CTkButton(
            btn_frame,
            text="📁 Thêm Thư Mục",
            width=130,
            command=self._add_folder
        )
        self.btn_add_folder.pack(side="left", padx=(0, 10))

        self.btn_clear = ctk.CTkButton(
            btn_frame,
            text="🗑️ Xóa Danh Sách",
            width=110,
            fg_color="gray40",
            hover_color="gray50",
            command=self._clear_files
        )
        self.btn_clear.pack(side="left", padx=(0, 10))

        self.btn_open_folder = ctk.CTkButton(
            btn_frame,
            text="📂 Mở Thư Mục Kết Quả",
            width=150,
            fg_color="gray30",
            hover_color="gray40",
            command=self._open_output_folder
        )
        self.btn_open_folder.pack(side="right", padx=(10, 0))

        self.btn_start = ctk.CTkButton(
            btn_frame,
            text="🚀 BẮT ĐẦU CHUYỂN ĐỔI",
            width=200,
            height=38,
            font=ctk.CTkFont(size=14, weight="bold"),
            fg_color="#28a745",
            hover_color="#218838",
            command=self._start_process
        )
        self.btn_start.pack(side="right")

    def _on_out_mode_changed(self):
        if self.out_mode_var.get() == "custom":
            self.entry_out_dir.configure(state="normal")
            self.btn_browse_out.configure(state="normal")
        else:
            self.entry_out_dir.configure(state="disabled")
            self.btn_browse_out.configure(state="disabled")

    def _choose_output_dir(self):
        d = filedialog.askdirectory(title="Chọn thư mục lưu file kết quả")
        if d:
            self.custom_folder_path.set(d)

    def _add_files(self):
        files = filedialog.askopenfilenames(
            title="Chọn các file Excel (.xls, .xlsx)",
            filetypes=[("Excel Files", "*.xls *.xlsx"), ("All Files", "*.*")]
        )
        if files:
            for f in files:
                norm_p = os.path.normpath(f)
                if not any(item["path"] == norm_p for item in self.file_items):
                    self.file_items.append({"path": norm_p, "status": "Chờ xử lý", "output": ""})
            self._refresh_file_list()

    def _add_folder(self):
        folder = filedialog.askdirectory(title="Chọn thư mục chứa các file Excel")
        if folder:
            added = 0
            for root, _, filenames in os.walk(folder):
                for fn in filenames:
                    ext = os.path.splitext(fn)[1].lower()
                    if ext in ('.xls', '.xlsx') and not fn.startswith('~$') and not fn.endswith('_fixed.xlsx'):
                        full_p = os.path.normpath(os.path.join(root, fn))
                        if not any(item["path"] == full_p for item in self.file_items):
                            self.file_items.append({"path": full_p, "status": "Chờ xử lý", "output": ""})
                            added += 1
            self._log(f"Đã quét thư mục '{folder}', tìm thấy {added} file Excel mới.")
            self._refresh_file_list()

    def _clear_files(self):
        if self.is_processing:
            return
        self.file_items.clear()
        self._refresh_file_list()
        self._log("Đã xóa sạch danh sách file.")

    def _refresh_file_list(self):
        # Clear widgets inside file_scroll_frame
        for widget in self.file_scroll_frame.winfo_children():
            widget.destroy()

        self.lbl_file_count.configure(text=f"Danh sách file ({len(self.file_items)} file)")

        for idx, item in enumerate(self.file_items):
            row = ctk.CTkFrame(self.file_scroll_frame, fg_color="transparent")
            row.pack(fill="x", pady=2, padx=5)

            name = os.path.basename(item["path"])
            lbl_name = ctk.CTkLabel(row, text=f"{idx+1}. {name}", anchor="w", width=300)
            lbl_name.pack(side="left", padx=(0, 10))

            color = "gray70"
            if item["status"] == "Hoàn thành":
                color = "#28a745"
            elif item["status"] == "Đang xử lý":
                color = "#ffc107"
            elif "Lỗi" in item["status"]:
                color = "#dc3545"

            lbl_status = ctk.CTkLabel(row, text=item["status"], text_color=color, anchor="e")
            lbl_status.pack(side="right")

    def _log(self, text):
        self.log_textbox.insert("end", text + "\n")
        self.log_textbox.see("end")

    def _open_output_folder(self):
        target_dir = None
        if self.out_mode_var.get() == "custom" and self.custom_folder_path.get():
            target_dir = self.custom_folder_path.get()
        elif self.file_items:
            # Open folder of first item
            first_out = self.file_items[0].get("output")
            if first_out and os.path.exists(first_out):
                target_dir = os.path.dirname(first_out)
            else:
                target_dir = os.path.dirname(self.file_items[0]["path"])

        if target_dir and os.path.exists(target_dir):
            os.startfile(target_dir)
        else:
            messagebox.showinfo("Thông báo", "Chưa có thư mục kết quả để mở!")

    def _start_process(self):
        if self.is_processing:
            return

        if not self.file_items:
            messagebox.showwarning("Chưa có file", "Vui lòng chọn ít nhất 1 file Excel để bắt đầu!")
            return

        if self.out_mode_var.get() == "custom":
            custom_dir = self.custom_folder_path.get().strip()
            if not custom_dir or not os.path.isdir(custom_dir):
                messagebox.showerror("Lỗi", "Thư mục tùy chọn không hợp lệ hoặc chưa được chọn!")
                return

        self.is_processing = True
        self._set_ui_state(processing=True)

        mode_str = self.mode_var.get()
        mode = 'auto'
        if "VNI" in mode_str:
            mode = 'vni'
        elif "TCVN3" in mode_str:
            mode = 'tcvn3'

        thread = threading.Thread(target=self._process_worker, args=(mode,), daemon=True)
        thread.start()

    def _set_ui_state(self, processing):
        st = "disabled" if processing else "normal"
        self.btn_add_files.configure(state=st)
        self.btn_add_folder.configure(state=st)
        self.btn_clear.configure(state=st)
        self.btn_start.configure(state=st)
        self.mode_menu.configure(state=st)
        self.radio_same.configure(state=st)
        self.radio_custom.configure(state=st)
        if not processing and self.out_mode_var.get() == "custom":
            self.entry_out_dir.configure(state="normal")
            self.btn_browse_out.configure(state="normal")
        else:
            self.entry_out_dir.configure(state="disabled")
            self.btn_browse_out.configure(state="disabled")

    def _process_worker(self, mode):
        total = len(self.file_items)
        success_count = 0
        error_count = 0

        self._log(f"--- BẮT ĐẦU CHUYỂN ĐỔI ({total} files, Chế độ: {self.mode_var.get()}) ---")

        for idx, item in enumerate(self.file_items):
            src = item["path"]
            filename = os.path.basename(src)
            base_name, _ = os.path.splitext(filename)

            if self.out_mode_var.get() == "custom":
                dst = os.path.join(self.custom_folder_path.get().strip(), f"{base_name}_fixed.xlsx")
            else:
                dst = os.path.join(os.path.dirname(src), f"{base_name}_fixed.xlsx")

            item["status"] = "Đang xử lý"
            self.after(0, self._refresh_file_list)
            self.after(0, lambda f=filename: self.status_lbl.configure(text=f"Đang xử lý: {f}"))

            def on_progress(p, msg):
                prog = (idx + p) / total
                self.after(0, lambda: self.progress_bar.set(prog))
                self.after(0, lambda m=msg: self._log(f"  > {m}"))

            try:
                out_path = convert_file(src, dst, mode=mode, progress_cb=on_progress)
                item["status"] = "Hoàn thành"
                item["output"] = out_path
                success_count += 1
                self.after(0, lambda f=filename: self._log(f"✅ Hoàn thành: {f} -> {os.path.basename(out_path)}"))
            except Exception as e:
                item["status"] = "Lỗi"
                error_count += 1
                self.after(0, lambda f=filename, err=str(e): self._log(f"❌ Lỗi file {f}: {err}"))

            self.after(0, self._refresh_file_list)

        self.after(0, lambda: self.progress_bar.set(1.0))
        self.after(0, lambda: self.status_lbl.configure(text=f"Xong! {success_count} thành công, {error_count} lỗi."))
        self.after(0, lambda: self._log(f"=== KẾT THÚC: {success_count} file thành công, {error_count} lỗi ==="))
        self.after(0, lambda: self._set_ui_state(processing=False))
        self.is_processing = False

        self.after(0, lambda: messagebox.showinfo(
            "Hoàn thành",
            f"Đã xử lý xong {total} file!\n• Thành công: {success_count}\n• Lỗi: {error_count}"
        ))


if __name__ == "__main__":
    app = App()
    app.mainloop()
