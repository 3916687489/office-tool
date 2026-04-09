import threading
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog

import pandas as pd
from PyPDF2 import PdfMerger, PdfReader, PdfWriter


class OfficeToolApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("综合执法办公小工具")
        self.root.geometry("400x300")
        self.root.resizable(False, False)

        self._build_ui()

    def _build_ui(self) -> None:
        container = tk.Frame(self.root, padx=20, pady=20)
        container.pack(expand=True, fill="both")

        # 统一按钮样式，保证四个按钮大小一致
        btn_style = {"width": 18, "height": 2, "font": ("Arial", 12)}

        self.btn_excel_merge = tk.Button(
            container, text="Excel 合并", command=self.excel_merge, **btn_style
        )
        self.btn_excel_split = tk.Button(
            container, text="Excel 拆分", command=self.excel_split, **btn_style
        )
        self.btn_pdf_merge = tk.Button(
            container, text="PDF 合并", command=self.pdf_merge, **btn_style
        )
        self.btn_pdf_split = tk.Button(
            container, text="PDF 拆分", command=self.pdf_split, **btn_style
        )

        self.btn_excel_merge.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")
        self.btn_excel_split.grid(row=0, column=1, padx=10, pady=10, sticky="nsew")
        self.btn_pdf_merge.grid(row=1, column=0, padx=10, pady=10, sticky="nsew")
        self.btn_pdf_split.grid(row=1, column=1, padx=10, pady=10, sticky="nsew")

        container.grid_rowconfigure(0, weight=1)
        container.grid_rowconfigure(1, weight=1)
        container.grid_columnconfigure(0, weight=1)
        container.grid_columnconfigure(1, weight=1)

    def _desktop_path(self) -> Path:
        desktop = Path.home() / "Desktop"
        return desktop if desktop.exists() else Path.home()

    def _timestamp(self) -> str:
        return datetime.now().strftime("%Y%m%d%H%M%S")

    def _set_buttons_state(self, enabled: bool) -> None:
        state = tk.NORMAL if enabled else tk.DISABLED
        for button in [
            self.btn_excel_merge,
            self.btn_excel_split,
            self.btn_pdf_merge,
            self.btn_pdf_split,
        ]:
            button.config(state=state)

    def _run_in_thread(self, task, success_message: str) -> None:
        # 所有耗时任务放到子线程，避免主界面卡顿
        self._set_buttons_state(False)

        def wrapper() -> None:
            try:
                task()
                self.root.after(
                    0, lambda: messagebox.showinfo("处理完成", success_message)
                )
            except Exception as exc:  # noqa: BLE001
                self.root.after(
                    0, lambda: messagebox.showerror("处理失败", f"发生错误：{exc}")
                )
            finally:
                self.root.after(0, lambda: self._set_buttons_state(True))

        threading.Thread(target=wrapper, daemon=True).start()

    def excel_merge(self) -> None:
        folder = filedialog.askdirectory(title="请选择包含 Excel 文件的文件夹")
        if not folder:
            messagebox.showinfo("已取消", "未选择文件夹，操作已取消。")
            return

        folder_path = Path(folder)

        def task() -> None:
            excel_files = sorted(
                [
                    p
                    for p in folder_path.glob("*.xlsx")
                    if p.is_file() and not p.name.startswith("~$")
                ],
                key=lambda x: x.name.lower(),
            )
            if not excel_files:
                raise ValueError("所选文件夹中未找到可用的 .xlsx 文件。")

            # 读取每个文件第一个工作表，并按行拼接
            all_frames = []
            for file in excel_files:
                df = pd.read_excel(file, sheet_name=0, engine="openpyxl")
                all_frames.append(df)

            merged_df = pd.concat(all_frames, ignore_index=True)
            out_file = self._desktop_path() / f"合并结果_{self._timestamp()}.xlsx"
            merged_df.to_excel(out_file, index=False, engine="openpyxl")

        self._run_in_thread(task, "Excel 合并完成，结果已输出到桌面。")

    def excel_split(self) -> None:
        file_path = filedialog.askopenfilename(
            title="请选择 Excel 文件",
            filetypes=[("Excel 文件", "*.xlsx")],
        )
        if not file_path:
            messagebox.showinfo("已取消", "未选择文件，操作已取消。")
            return

        rows_per_file = simpledialog.askinteger(
            "输入拆分行数",
            "请输入每个文件的行数（默认500，最小1）：",
            minvalue=1,
            initialvalue=500,
        )
        if rows_per_file is None:
            messagebox.showinfo("已取消", "未输入拆分行数，操作已取消。")
            return

        src = Path(file_path)

        def task() -> None:
            df = pd.read_excel(src, sheet_name=0, engine="openpyxl")
            total_rows = len(df)
            if total_rows == 0:
                raise ValueError("Excel 文件没有可拆分的数据行。")

            out_dir = src.parent / "Excel拆分结果"
            out_dir.mkdir(exist_ok=True)

            part_index = 1
            for start in range(0, total_rows, rows_per_file):
                end = min(start + rows_per_file, total_rows)
                chunk = df.iloc[start:end]
                out_name = f"part_{part_index}_{end}.xlsx"
                chunk.to_excel(out_dir / out_name, index=False, engine="openpyxl")
                part_index += 1

        self._run_in_thread(task, "Excel 拆分完成，结果已保存到 Excel拆分结果 文件夹。")

    def pdf_merge(self) -> None:
        folder = filedialog.askdirectory(title="请选择包含 PDF 文件的文件夹")
        if not folder:
            messagebox.showinfo("已取消", "未选择文件夹，操作已取消。")
            return

        folder_path = Path(folder)

        def task() -> None:
            pdf_files = sorted(
                [p for p in folder_path.glob("*.pdf") if p.is_file()],
                key=lambda x: x.name.lower(),
            )
            if not pdf_files:
                raise ValueError("所选文件夹中未找到 .pdf 文件。")

            out_file = self._desktop_path() / f"合并PDF_{self._timestamp()}.pdf"

            merger = PdfMerger()
            try:
                for pdf in pdf_files:
                    merger.append(str(pdf))
                with open(out_file, "wb") as f:
                    merger.write(f)
            finally:
                merger.close()

        self._run_in_thread(task, "PDF 合并完成，结果已输出到桌面。")

    def pdf_split(self) -> None:
        file_path = filedialog.askopenfilename(
            title="请选择 PDF 文件",
            filetypes=[("PDF 文件", "*.pdf")],
        )
        if not file_path:
            messagebox.showinfo("已取消", "未选择文件，操作已取消。")
            return

        pages_per_file = simpledialog.askinteger(
            "输入拆分页数",
            "请输入每个文件包含的页数（默认10，最小1）：",
            minvalue=1,
            initialvalue=10,
        )
        if pages_per_file is None:
            messagebox.showinfo("已取消", "未输入拆分页数，操作已取消。")
            return

        src = Path(file_path)

        def task() -> None:
            reader = PdfReader(str(src))
            total_pages = len(reader.pages)
            if total_pages == 0:
                raise ValueError("PDF 文件没有可拆分的页面。")

            out_dir = src.parent / "PDF拆分结果"
            out_dir.mkdir(exist_ok=True)
            base_name = src.stem

            for start in range(0, total_pages, pages_per_file):
                end = min(start + pages_per_file, total_pages)
                writer = PdfWriter()
                for page_idx in range(start, end):
                    writer.add_page(reader.pages[page_idx])

                out_name = f"{base_name}_第{start + 1}-{end}页.pdf"
                with open(out_dir / out_name, "wb") as f:
                    writer.write(f)

        self._run_in_thread(task, "PDF 拆分完成，结果已保存到 PDF拆分结果 文件夹。")


def main() -> None:
    root = tk.Tk()
    OfficeToolApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
