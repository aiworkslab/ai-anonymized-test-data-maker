from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import pandas as pd


APP_TITLE = "AI Anonymized Test Data Maker"
SUPPORTED_EXTENSIONS = {".csv", ".xlsx"}


def validate_input_file(file_path):
    path = Path(file_path)
    extension = path.suffix.lower()
    if extension not in SUPPORTED_EXTENSIONS:
        raise ValueError(".xlsx または .csv ファイルを選択してください。")
    return path


def read_csv_columns(file_path):
    last_error = None
    for encoding in ("utf-8", "cp932"):
        try:
            dataframe = pd.read_csv(file_path, nrows=0, encoding=encoding)
            return list(dataframe.columns), encoding
        except UnicodeDecodeError as error:
            last_error = error

    raise UnicodeDecodeError(
        last_error.encoding,
        last_error.object,
        last_error.start,
        last_error.end,
        "UTF-8 と CP932 のどちらでも読み込めませんでした。",
    )


def get_excel_sheets(file_path):
    excel_file = pd.ExcelFile(file_path, engine="openpyxl")
    return excel_file.sheet_names


def read_excel_columns(file_path, sheet_name):
    dataframe = pd.read_excel(
        file_path,
        sheet_name=sheet_name,
        nrows=0,
        engine="openpyxl",
    )
    return list(dataframe.columns)


class AnonymizedTestDataMakerApp:
    def __init__(self, root):
        self.root = root
        self.root.title(APP_TITLE)
        self.root.geometry("720x520")
        self.root.minsize(560, 420)

        self.selected_file_path = None
        self.sheet_names = []

        self.file_name_var = tk.StringVar(value="未選択")
        self.sheet_var = tk.StringVar()
        self.status_var = tk.StringVar(value="ファイルを選択してください。")

        self._build_widgets()

    def _build_widgets(self):
        main_frame = ttk.Frame(self.root, padding=16)
        main_frame.pack(fill=tk.BOTH, expand=True)
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(4, weight=1)

        title_label = ttk.Label(
            main_frame,
            text="AI Anonymized Test Data Maker",
            font=("", 16, "bold"),
        )
        title_label.grid(row=0, column=0, sticky=tk.W, pady=(0, 14))

        select_button = ttk.Button(
            main_frame,
            text="ファイルを選択",
            command=self.select_file,
        )
        select_button.grid(row=1, column=0, sticky=tk.W, pady=(0, 10))

        file_frame = ttk.Frame(main_frame)
        file_frame.grid(row=2, column=0, sticky=tk.EW, pady=(0, 10))
        file_frame.columnconfigure(1, weight=1)

        ttk.Label(file_frame, text="選択したファイル名:").grid(
            row=0,
            column=0,
            sticky=tk.W,
            padx=(0, 8),
        )
        ttk.Label(file_frame, textvariable=self.file_name_var).grid(
            row=0,
            column=1,
            sticky=tk.EW,
        )

        sheet_frame = ttk.Frame(main_frame)
        sheet_frame.grid(row=3, column=0, sticky=tk.EW, pady=(0, 12))
        sheet_frame.columnconfigure(1, weight=1)

        ttk.Label(sheet_frame, text="Excelシート選択:").grid(
            row=0,
            column=0,
            sticky=tk.W,
            padx=(0, 8),
        )
        self.sheet_combo = ttk.Combobox(
            sheet_frame,
            textvariable=self.sheet_var,
            state="disabled",
            values=[],
        )
        self.sheet_combo.grid(row=0, column=1, sticky=tk.EW)
        self.sheet_combo.bind("<<ComboboxSelected>>", self.on_sheet_selected)

        columns_frame = ttk.LabelFrame(main_frame, text="列名一覧", padding=10)
        columns_frame.grid(row=4, column=0, sticky=tk.NSEW, pady=(0, 12))
        columns_frame.columnconfigure(0, weight=1)
        columns_frame.rowconfigure(0, weight=1)

        self.columns_listbox = tk.Listbox(columns_frame, activestyle="none")
        self.columns_listbox.grid(row=0, column=0, sticky=tk.NSEW)

        scrollbar = ttk.Scrollbar(
            columns_frame,
            orient=tk.VERTICAL,
            command=self.columns_listbox.yview,
        )
        scrollbar.grid(row=0, column=1, sticky=tk.NS)
        self.columns_listbox.configure(yscrollcommand=scrollbar.set)

        status_frame = ttk.LabelFrame(main_frame, text="状態メッセージ", padding=10)
        status_frame.grid(row=5, column=0, sticky=tk.EW)
        status_frame.columnconfigure(0, weight=1)

        ttk.Label(status_frame, textvariable=self.status_var).grid(
            row=0,
            column=0,
            sticky=tk.EW,
        )

    def select_file(self):
        file_path = filedialog.askopenfilename(
            title="入力ファイルを選択",
            filetypes=[
                ("Excel または CSV", "*.xlsx *.csv"),
                ("Excel ファイル", "*.xlsx"),
                ("CSV ファイル", "*.csv"),
            ],
        )
        if not file_path:
            return

        try:
            path = validate_input_file(file_path)
            self.selected_file_path = path
            self.file_name_var.set(path.name)
            self.clear_columns()

            if path.suffix.lower() == ".csv":
                self.disable_sheet_selection()
                columns, encoding = read_csv_columns(path)
                self.show_columns(columns)
                self.status_var.set(
                    f"CSVファイルを読み込みました。文字コード: {encoding}"
                )
                return

            self.sheet_names = get_excel_sheets(path)
            if not self.sheet_names:
                raise ValueError("Excelファイルにシートが見つかりません。")

            self.sheet_combo.configure(state="readonly", values=self.sheet_names)
            self.sheet_var.set(self.sheet_names[0])
            self.load_excel_sheet_columns(self.sheet_names[0])
            self.status_var.set("Excelファイルを読み込みました。シートを選択できます。")

        except Exception as error:
            self.reset_selection_after_error()
            self.show_error(error)

    def on_sheet_selected(self, _event):
        sheet_name = self.sheet_var.get()
        if not self.selected_file_path or not sheet_name:
            return

        try:
            self.load_excel_sheet_columns(sheet_name)
            self.status_var.set(f"シート「{sheet_name}」の列名を表示しました。")
        except Exception as error:
            self.clear_columns()
            self.show_error(error)

    def load_excel_sheet_columns(self, sheet_name):
        columns = read_excel_columns(self.selected_file_path, sheet_name)
        self.show_columns(columns)

    def show_columns(self, columns):
        self.clear_columns()
        for column in columns:
            self.columns_listbox.insert(tk.END, str(column))

        if columns:
            self.status_var.set(f"{len(columns)}件の列名を表示しました。")
        else:
            self.status_var.set("列名が見つかりませんでした。")

    def clear_columns(self):
        self.columns_listbox.delete(0, tk.END)

    def disable_sheet_selection(self):
        self.sheet_names = []
        self.sheet_var.set("")
        self.sheet_combo.configure(state="disabled", values=[])

    def reset_selection_after_error(self):
        self.selected_file_path = None
        self.file_name_var.set("未選択")
        self.disable_sheet_selection()
        self.clear_columns()

    def show_error(self, error):
        message = f"読み込みに失敗しました: {error}"
        self.status_var.set(message)
        messagebox.showerror("読み込みエラー", message)


def main():
    root = tk.Tk()
    AnonymizedTestDataMakerApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
