"""Operator UI for the fictional fixed-schema vehicle pipeline.

No web server, network request, external API or save dialog is used.
"""
from __future__ import annotations

import sys
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, ttk

sys.dont_write_bytecode = True
import pipeline

HERE = Path(__file__).resolve().parent
DEMO = HERE


class OperatorApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        root.title("車両照合・操作画面｜架空データ専用")
        root.geometry("1000x790")
        root.minsize(850, 680)
        self.values = {
            "vehicles": tk.StringVar(value=str(DEMO / "vehicles.xlsx")),
            "contracts": tk.StringVar(value=str(DEMO / "contracts.pdf")),
            "installments": tk.StringVar(value=str(DEMO / "installments.pdf")),
            "month": tk.StringVar(value="2027-01"),
            "out": tk.StringVar(value=str(self.next_output())),
        }
        frame = ttk.Frame(root, padding=18)
        frame.pack(fill="both", expand=True)
        frame.columnconfigure(1, weight=1)
        frame.rowconfigure(10, weight=1)
        ttk.Label(frame, text="車両Excel → 契約・割賦PDF → 指定月CSV", font=("Yu Gothic UI", 16, "bold")).grid(row=0, column=0, columnspan=3, sticky="w")
        ttk.Label(frame, text="固定架空形式のデモ。固定形式のみ。実資料・個人情報は入力しないでください。", foreground="#9a3412").grid(row=1, column=0, columnspan=3, sticky="w", pady=(8, 16))
        for row, (key, label, suffix) in enumerate([
            ("vehicles", "車両Excel", ".xlsx"),
            ("contracts", "契約PDF", ".pdf"),
            ("installments", "割賦PDF", ".pdf"),
        ], 2):
            ttk.Label(frame, text=label).grid(row=row, column=0, sticky="w", padx=(0, 10), pady=7)
            ttk.Entry(frame, textvariable=self.values[key]).grid(row=row, column=1, sticky="ew", pady=7)
            ttk.Button(frame, text="選択…", command=lambda k=key, s=suffix: self.choose(k, s)).grid(row=row, column=2, padx=(8, 0))
        ttk.Label(frame, text="対象月 YYYY-MM").grid(row=5, column=0, sticky="w", pady=7)
        ttk.Entry(frame, textvariable=self.values["month"]).grid(row=5, column=1, sticky="ew", pady=7)
        ttk.Label(frame, text="新規出力フォルダー").grid(row=6, column=0, sticky="w", pady=7)
        ttk.Entry(frame, textvariable=self.values["out"]).grid(row=6, column=1, sticky="ew", pady=7)
        ttk.Button(frame, text="新しい出力先", command=self.new_output).grid(row=6, column=2, padx=(8, 0))
        ttk.Label(frame, text="まだ存在しないフォルダーへ自動保存します。既存先は上書きしません。", foreground="#475569").grid(row=7, column=1, columnspan=2, sticky="w")
        self.run_button = ttk.Button(frame, text="照合して保存", command=self.execute)
        self.run_button.grid(row=8, column=1, sticky="w", pady=14)
        self.status = tk.StringVar(value="入力3点と対象月を確認して実行してください。保存ダイアログは開きません。")
        ttk.Label(frame, textvariable=self.status, wraplength=920).grid(row=9, column=0, columnspan=3, sticky="w", pady=(0, 8))
        result_frame = ttk.Frame(frame)
        result_frame.grid(row=10, column=0, columnspan=3, sticky="nsew")
        self.result = tk.Text(result_frame, height=16, wrap="word", font=("Yu Gothic UI", 11), state="disabled")
        scrollbar = ttk.Scrollbar(result_frame, orient="vertical", command=self.result.yview)
        self.result.configure(yscrollcommand=scrollbar.set)
        self.result.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        ttk.Label(frame, text="出典付き表示は登録番号・型式の既知不一致に対応。ほかの失敗は理由のみの場合があります。", foreground="#475569").grid(row=11, column=0, columnspan=3, sticky="w", pady=(12, 0))
        ttk.Label(frame, text="任意帳票/OCR/実会計CSV/現場導入は未検証。人が元資料を修正して新しい出力先で再実行します。", foreground="#475569").grid(row=12, column=0, columnspan=3, sticky="w", pady=(5, 0))

    @staticmethod
    def next_output() -> Path:
        return HERE / ("output-" + datetime.now().strftime("%Y%m%d-%H%M%S-%f"))

    def new_output(self):
        self.values["out"].set(str(self.next_output()))

    def choose(self, key, suffix):
        selected = filedialog.askopenfilename(parent=self.root, title="架空入力を選択", filetypes=[("入力ファイル", "*" + suffix)])
        if selected:
            self.values[key].set(selected)

    def show_result(self, text):
        self.result.configure(state="normal")
        self.result.delete("1.0", "end")
        self.result.insert("1.0", text)
        self.result.configure(state="disabled")

    def execute(self):
        self.run_button.state(["disabled"])
        self.status.set("照合中です。処理が終わるまで再実行しないでください。")
        self.root.update_idletasks()
        try:
            args = {key: value.get().strip() for key, value in self.values.items()}
            if any(not value for value in args.values()):
                raise ValueError("入力3点、対象月、新規出力先を指定してください")
            if not Path(args["out"]).is_absolute():
                raise ValueError("新規出力先は絶対パスで指定してください")
            result = pipeline.run(**args)
            output = Path(args["out"]).resolve()
            self.status.set("保存完了：架空入力の照合を通過しました。実帳票への適合を示すものではありません。")
            self.show_result(
                f"対象月：{result['month']}\n対象月の合計：{result['selected_total_jpy']:,}円\n"
                f"車両数：{result['vehicles']}台 / 全支払明細：{result['payment_rows']}件\n\n"
                + "\n".join(str(output / name) for name in ("payment-schedule.xlsx", "finance.csv", "source-report.json"))
                + "\n\n上の3ファイルを自動保存しました。入力は変更していません。"
            )
        except pipeline.SourceMismatch as error:
            data = error.details
            self.status.set("照合で停止：不一致の出典を確認してください。CSVは作成していません。")
            rows = [f"車両：{data['vehicle_id']}", f"Excel出典：{data['xlsx_source']}",
                    f"契約PDF出典：{data['contract_source']}", ""]
            for difference in data["differences"]:
                label = {"registration": "登録番号", "model": "型式"}[difference["field"]]
                rows.extend([label, f"  Excel：{difference['xlsx_value']}", f"  PDF：{difference['pdf_value']}"])
            rows.append("\n入力を確認・修正した後、新しい出力先を指定して再実行してください。")
            self.show_result("\n".join(rows))
        except Exception as error:
            self.status.set("処理を停止しました。理由と出力先を確認し、内容不明のまま再実行しないでください。")
            self.show_result(f"{type(error).__name__}: {error}\n\n出力先が残っている場合は一部保存の可能性があります。自動再試行や上書きはしません。")
        finally:
            self.run_button.state(["!disabled"])


if __name__ == "__main__":
    window = tk.Tk()
    OperatorApp(window)
    window.mainloop()
