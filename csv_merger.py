"""CSV選考データマージツール

3つのD&D枠（2次選考 / 最終選考 / 内定出し）にCSVをドロップし、
各CSVから「日時」と「氏名」の2列を抜き出して、種別を付けた1つのCSVに集約する。
"""
from __future__ import annotations

import os
import re
import subprocess
import sys
import tkinter as tk
from datetime import datetime
from pathlib import Path
from tkinter import messagebox, ttk

import pandas as pd

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
except ImportError:
    print("tkinterdnd2 が未インストールです。`pip install tkinterdnd2` を実行してください。")
    sys.exit(1)


SECTIONS = [
    {
        "name": "2次選考",
        "date_col_name": "2次選考",
        "date_col_index": 2,   # C列
        "name_col_name": "グループ選考 【総合評価】（ユーザ）",
        "name_col_index": 7,   # H列
    },
    {
        "name": "最終選考",
        "date_col_name": "最終選考会",
        "date_col_index": 3,   # D列
        "name_col_name": "最終 【総合評価】（ユーザ）",
        "name_col_index": 9,   # J列
    },
    {
        "name": "内定出し",
        "date_col_name": "内定だし（＝面談）",
        "date_col_index": 4,   # E列
        "name_col_name": "内定通知 1.入社させたいか（ユーザ）",
        "name_col_index": 11,  # L列
    },
]

ENCODING_CANDIDATES = ["utf-8-sig", "cp932", "utf-8", "latin-1"]

# 元データのセル例: "合格 2026-03-04 13:00 WEB【1】" / "予約 2026-03-24 13:00 WEB【1】"
# 文字列中の "YYYY-M-D HH:MM"（区切りは - / 年月日いずれも許容）を抽出する。
DATETIME_PATTERN = re.compile(
    r"(\d{4})[-/年](\d{1,2})[-/月](\d{1,2})日?\s*[T\s]\s*(\d{1,2})[:時](\d{2})"
)

# 氏名の姓・名区切り: 半角または全角スペース（連続は1区切りとみなす）
NAME_SPLIT_PATTERN = re.compile(r"[ 　]+")


def read_csv_auto(path: Path) -> pd.DataFrame:
    """エンコーディング候補を順に試してCSVを読み込む。"""
    last_err: Exception | None = None
    for enc in ENCODING_CANDIDATES:
        try:
            return pd.read_csv(path, encoding=enc, dtype=str, keep_default_na=False)
        except (UnicodeDecodeError, UnicodeError) as e:
            last_err = e
            continue
        except Exception as e:
            last_err = e
            continue
    raise RuntimeError(f"CSVの読み込みに失敗しました: {path.name} ({last_err})")


def format_datetime(value) -> str:
    """文字列から日時部分を抽出し `yyyy/M/d HH:mm` 形式に整形。

    元データのセルは "合格 2026-03-04 13:00 WEB【1】" のようにステータス語や場所名と
    一緒に格納されているため、まず正規表現で日時部分のみを抜き出す。
    日時が含まれない場合は空文字を返す。時刻の値（年月日・時分）は絶対に変更しない。
    """
    if value is None:
        return ""
    s = str(value).strip()
    if s == "" or s.lower() == "nan":
        return ""
    m = DATETIME_PATTERN.search(s)
    if m:
        year, month, day, hour, minute = m.groups()
        return f"{int(year)}/{int(month)}/{int(day)} {int(hour):02d}:{minute}"
    # フォールバック: 文字列全体を pandas でパース（純粋な日時文字列向け）
    parsed = pd.to_datetime(s, errors="coerce")
    if pd.isna(parsed):
        return ""
    return f"{parsed.year}/{parsed.month}/{parsed.day} {parsed.strftime('%H:%M')}"


def split_name(value) -> tuple[str, str]:
    """氏名を姓と名に分割する。

    - 区切り（半角/全角スペース）の最初の出現位置のみで分割
    - 区切りが見つからなければ全文字を姓に格納し、名は空（データ脱落防止）
    - 空文字や None は ("", "") を返す
    """
    if value is None:
        return ("", "")
    s = str(value).strip()
    if s == "" or s.lower() == "nan":
        return ("", "")
    parts = NAME_SPLIT_PATTERN.split(s, maxsplit=1)
    if len(parts) == 1:
        return (parts[0], "")
    return (parts[0].strip(), parts[1].strip())


def pick_column(df: pd.DataFrame, name: str, idx: int) -> pd.Series | None:
    """列名優先、見つからなければ位置（インデックス）でフォールバック。"""
    if name in df.columns:
        return df[name]
    if 0 <= idx < df.shape[1]:
        return df.iloc[:, idx]
    return None


def extract_section(files: list[Path], section: dict) -> tuple[pd.DataFrame, list[str]]:
    """指定セクションのCSV群から日時・氏名を抽出して長形式DataFrameを返す。

    戻り値: (DataFrame[日時, 氏名, 種別], 警告メッセージ一覧)
    """
    warnings: list[str] = []
    rows: list[pd.DataFrame] = []

    for path in files:
        try:
            df = read_csv_auto(path)
        except Exception as e:
            warnings.append(f"[{section['name']}] 読み込み失敗: {path.name} — {e}")
            continue

        if df.empty:
            warnings.append(f"[{section['name']}] 空のCSV: {path.name}")
            continue

        date_series = pick_column(df, section["date_col_name"], section["date_col_index"])
        name_series = pick_column(df, section["name_col_name"], section["name_col_index"])

        if date_series is None and name_series is None:
            warnings.append(
                f"[{section['name']}] 対象列が見つかりません（列名・位置とも不一致）: {path.name}"
            )
            continue

        if date_series is None:
            date_series = pd.Series([""] * len(df))
        if name_series is None:
            name_series = pd.Series([""] * len(df))

        # 氏名を姓・名に分割
        name_pairs = name_series.astype(str).map(split_name)
        sei_series = name_pairs.map(lambda t: t[0])
        mei_series = name_pairs.map(lambda t: t[1])

        sub = pd.DataFrame({
            "日時": date_series.astype(str).map(format_datetime),
            "姓": sei_series,
            "名": mei_series,
        })
        # 姓・名がともに空の行は除外（=元の氏名が空白）
        sub = sub[~((sub["姓"] == "") & (sub["名"] == ""))]
        sub["種別"] = section["name"]
        rows.append(sub)

    if not rows:
        return pd.DataFrame(columns=["日時", "姓", "名", "種別"]), warnings
    return pd.concat(rows, ignore_index=True), warnings


def get_downloads_dir() -> Path:
    """Windowsのダウンロードフォルダのパスを返す（無ければホーム直下にフォールバック）。"""
    home = Path.home()
    downloads = home / "Downloads"
    downloads.mkdir(parents=True, exist_ok=True)
    return downloads


def open_in_explorer(path: Path) -> None:
    """エクスプローラーで指定ファイルを選択状態で開く（Windows専用）。"""
    try:
        subprocess.Popen(["explorer", "/select,", str(path)])
    except Exception:
        # フォールバック: フォルダだけでも開く
        try:
            os.startfile(path.parent)
        except Exception:
            pass


class DropZone(ttk.LabelFrame):
    """D&D対応のCSV受領フレーム。"""

    def __init__(self, parent: tk.Widget, title: str):
        super().__init__(parent, text=f"＜{title}＞", padding=8)
        self.title_text = title
        self.files: list[Path] = []

        hint = ttk.Label(
            self,
            text="ここにCSVをドラッグ&ドロップ\n（複数可）",
            anchor="center",
            justify="center",
            foreground="#666",
        )
        hint.pack(fill="x", pady=(0, 4))

        self.listbox = tk.Listbox(self, height=10, activestyle="none")
        self.listbox.pack(fill="both", expand=True)

        btn_row = ttk.Frame(self)
        btn_row.pack(fill="x", pady=(6, 0))
        ttk.Button(btn_row, text="選択を削除", command=self._remove_selected).pack(side="left")
        ttk.Button(btn_row, text="全クリア", command=self._clear_all).pack(side="right")

        # D&D登録（フレーム本体とListbox両方）
        for widget in (self, hint, self.listbox):
            widget.drop_target_register(DND_FILES)
            widget.dnd_bind("<<Drop>>", self._on_drop)

    def _on_drop(self, event) -> str:
        # event.data はスペース区切り。空白を含むパスは{}で囲まれる。
        raw = event.data
        paths = self._parse_dnd_paths(raw)
        added = 0
        skipped: list[str] = []
        for p in paths:
            path = Path(p)
            if not path.exists() or path.suffix.lower() != ".csv":
                skipped.append(path.name)
                continue
            if path in self.files:
                continue
            self.files.append(path)
            self.listbox.insert("end", path.name)
            added += 1
        if skipped:
            messagebox.showinfo(
                "スキップ",
                f"CSV以外のファイルはスキップしました:\n" + "\n".join(skipped),
            )
        return "break"

    @staticmethod
    def _parse_dnd_paths(data: str) -> list[str]:
        """tkinterdnd2の<<Drop>>イベントdata文字列をパス一覧に変換。"""
        result: list[str] = []
        i = 0
        n = len(data)
        while i < n:
            if data[i] == " ":
                i += 1
                continue
            if data[i] == "{":
                j = data.find("}", i + 1)
                if j == -1:
                    result.append(data[i + 1:])
                    break
                result.append(data[i + 1:j])
                i = j + 1
            else:
                j = data.find(" ", i)
                if j == -1:
                    result.append(data[i:])
                    break
                result.append(data[i:j])
                i = j + 1
        return result

    def _remove_selected(self) -> None:
        for idx in reversed(self.listbox.curselection()):
            del self.files[idx]
            self.listbox.delete(idx)

    def _clear_all(self) -> None:
        self.files.clear()
        self.listbox.delete(0, "end")


class App:
    def __init__(self, root: TkinterDnD.Tk):
        self.root = root
        root.title("CSV選考データマージ")
        root.minsize(720, 400)

        container = ttk.Frame(root, padding=10)
        container.pack(fill="both", expand=True)

        zones_frame = ttk.Frame(container)
        zones_frame.pack(fill="both", expand=True)

        self.zones: list[DropZone] = []
        for i, sec in enumerate(SECTIONS):
            zone = DropZone(zones_frame, sec["name"])
            zone.grid(row=0, column=i, sticky="nsew", padx=4)
            zones_frame.columnconfigure(i, weight=1)
            self.zones.append(zone)
        zones_frame.rowconfigure(0, weight=1)

        btn_row = ttk.Frame(container)
        btn_row.pack(fill="x", pady=(10, 0))
        ttk.Button(btn_row, text="マージしてCSV出力", command=self.on_merge_click).pack()

        self.status = ttk.Label(container, text="準備完了", foreground="#444")
        self.status.pack(fill="x", pady=(8, 0))

    def on_merge_click(self) -> None:
        all_files = [zone.files for zone in self.zones]
        if all(len(f) == 0 for f in all_files):
            messagebox.showwarning("ファイルなし", "少なくとも1つの枠にCSVをドロップしてください。")
            return

        all_warnings: list[str] = []
        frames: list[pd.DataFrame] = []
        for zone, section in zip(self.zones, SECTIONS):
            if not zone.files:
                continue
            df, warnings = extract_section(zone.files, section)
            all_warnings.extend(warnings)
            if not df.empty:
                frames.append(df)

        if not frames:
            messagebox.showerror(
                "抽出失敗",
                "対象列が抽出できませんでした。\n\n" + "\n".join(all_warnings),
            )
            return

        merged = pd.concat(frames, ignore_index=True)

        out_dir = get_downloads_dir()
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        out_path = out_dir / f"merged_{timestamp}.csv"

        try:
            merged.to_csv(out_path, index=False, encoding="utf-8-sig")
        except Exception as e:
            messagebox.showerror("書き出し失敗", f"CSVの書き出しに失敗しました:\n{e}")
            return

        self.status.config(text=f"出力: {out_path}")

        warn_text = "\n\n警告:\n" + "\n".join(all_warnings) if all_warnings else ""
        messagebox.showinfo(
            "完了",
            f"マージしました（{len(merged)}行）。\n保存先:\n{out_path}{warn_text}",
        )
        open_in_explorer(out_path)


def main() -> None:
    root = TkinterDnD.Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
