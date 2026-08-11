import os
import sys
import threading
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional, Sequence, Set, Tuple

import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

APP_TITLE = "명성공업 서열정보&소요자재 자동 취합 프로그램"
OUTPUT_PREFIX = "서열_자재_통합"
MONTH_KEY = "생산번호"
MATERIAL_KEY = "물류번호"
EXCEL_FILETYPES = [("Excel files", "*.xlsx *.xlsm *.xls"), ("All files", "*.*")]
HI_SRM_MONTH_HEADER_ROW = 1
HI_SRM_MATERIAL_HEADER_ROW = 0
DEFAULT_MONTH_COLUMNS = [
    "생산번호",
    "영업모델",
    "차대호기",
    "착수일",
    "국가",
    "후방",
    "Radar",
    "COWL",
    "습식크리너",
    "G_RAIL",
    "특이도장",
]
DEFAULT_MATERIAL_COLUMNS = ["착수일자", "물류번호", "자재번호", "품명"]


def excel_read_path(path: str) -> str:
    resolved = str(Path(path).resolve())
    if os.name == "nt" and not resolved.startswith("\\\\?\\"):
        return "\\\\?\\" + resolved
    return resolved


def normalize_text(value: object) -> str:
    if pd.isna(value):
        return ""
    text = str(value).strip()
    if text.endswith(".0"):
        text = text[:-2]
    return text


START_DATE_KEYWORD = "착수일"
START_DATE_OUTPUT = "착수일자"
MONTH_START_DATE_INTERNAL = "__month_start_date"
MATERIAL_START_DATE_INTERNAL = "__material_start_date"


def normalize_date_value(value: object) -> str:
    if pd.isna(value):
        return ""
    if isinstance(value, datetime):
        return value.strftime("%Y-%m-%d")
    text = str(value).strip()
    if not text:
        return ""
    if text.endswith(".0"):
        text = text[:-2]
    digits = "".join(ch for ch in text if ch.isdigit())
    if len(digits) == 8:
        parsed = pd.to_datetime(digits, format="%Y%m%d", errors="coerce")
    else:
        parsed = pd.to_datetime(text, errors="coerce")
    if pd.isna(parsed):
        return text[:10]
    return parsed.strftime("%Y-%m-%d")


def find_start_date_column(columns: Sequence[str]) -> Optional[str]:
    for preferred in (START_DATE_OUTPUT, START_DATE_KEYWORD):
        if preferred in columns:
            return preferred
    for col in columns:
        if START_DATE_KEYWORD in col:
            return col
    return None


def normalize_start_date_columns(frame: pd.DataFrame) -> pd.DataFrame:
    for col in list(frame.columns):
        if START_DATE_KEYWORD in col:
            frame[col] = frame[col].map(normalize_date_value)
    return frame


def columns_with_required(columns: Sequence[str], selected: Set[str], required: Sequence[str]) -> List[str]:
    wanted = [col for col in columns if col in selected]
    for col in reversed(list(required)):
        if col in columns and col not in wanted:
            wanted.insert(0, col)
    return wanted


def collapse_key_and_start_date_columns(frame: pd.DataFrame) -> pd.DataFrame:
    result = frame.copy()
    if MATERIAL_KEY in result.columns and MONTH_KEY in result.columns:
        result = result.drop(columns=[MATERIAL_KEY])

    material_date = result[MATERIAL_START_DATE_INTERNAL] if MATERIAL_START_DATE_INTERNAL in result.columns else pd.Series([""] * len(result), index=result.index)
    month_date = result[MONTH_START_DATE_INTERNAL] if MONTH_START_DATE_INTERNAL in result.columns else pd.Series([""] * len(result), index=result.index)
    output_date = material_date.where(material_date.astype(str) != "", month_date)

    drop_cols = [
        col for col in result.columns
        if col in (MONTH_START_DATE_INTERNAL, MATERIAL_START_DATE_INTERNAL)
        or (START_DATE_KEYWORD in col and col != START_DATE_OUTPUT)
    ]
    result = result.drop(columns=[col for col in drop_cols if col in result.columns])
    key_col = MONTH_KEY if MONTH_KEY in result.columns else MATERIAL_KEY if MATERIAL_KEY in result.columns else None
    if START_DATE_OUTPUT in result.columns:
        result[START_DATE_OUTPUT] = output_date
    else:
        insert_at = list(result.columns).index(key_col) + 1 if key_col else 0
        result.insert(insert_at, START_DATE_OUTPUT, output_date)
    leading = [col for col in (key_col, START_DATE_OUTPUT) if col and col in result.columns]
    trailing = [col for col in result.columns if col not in leading]
    return result[leading + trailing]


def normalize_columns(columns: Sequence[object]) -> List[str]:
    seen: Dict[str, int] = {}
    result: List[str] = []
    for idx, col in enumerate(columns, start=1):
        name = normalize_text(col)
        if not name or name.lower().startswith("unnamed"):
            name = f"빈컬럼_{idx}"
        count = seen.get(name, 0)
        seen[name] = count + 1
        result.append(name if count == 0 else f"{name}_{count + 1}")
    return result


def find_header_row(path: str, required_key: str) -> int:
    preview = pd.read_excel(excel_read_path(path), sheet_name=0, header=None, nrows=30, dtype=object)
    best_row = 0
    best_score = -1
    for idx, row in preview.iterrows():
        values = [normalize_text(value) for value in row.tolist()]
        non_empty = [value for value in values if value]
        if required_key in non_empty:
            return int(idx)
        if len(non_empty) > best_score:
            best_row = int(idx)
            best_score = len(non_empty)
    return best_row


def read_excel_frame(
    path: str,
    required_key: str,
    preferred_header_row: Optional[int] = None,
    expected_column_count: Optional[int] = None,
) -> Tuple[pd.DataFrame, int]:
    header_row = preferred_header_row if preferred_header_row is not None else find_header_row(path, required_key)
    frame = pd.read_excel(excel_read_path(path), sheet_name=0, header=header_row, dtype=object)
    frame.columns = normalize_columns(frame.columns)
    if required_key not in frame.columns and preferred_header_row is not None:
        header_row = find_header_row(path, required_key)
        frame = pd.read_excel(excel_read_path(path), sheet_name=0, header=header_row, dtype=object)
        frame.columns = normalize_columns(frame.columns)
    frame = frame.dropna(how="all").reset_index(drop=True)
    return frame, header_row


def selected_columns_with_key(columns: Sequence[str], selected: Set[str], key: str) -> List[str]:
    return columns_with_required(columns, selected, [key])


def default_selected_columns(columns: Sequence[str], defaults: Sequence[str]) -> Set[str]:
    by_exact = {col: col for col in columns}
    by_lower = {col.lower(): col for col in columns}
    selected: Set[str] = set()
    for default in defaults:
        if default in by_exact:
            selected.add(by_exact[default])
        else:
            match = by_lower.get(default.lower())
            if match:
                selected.add(match)
    return selected


def canonicalize_default_columns(frame: pd.DataFrame, defaults: Sequence[str]) -> pd.DataFrame:
    rename_map: Dict[str, str] = {}
    default_by_lower = {default.lower(): default for default in defaults}
    for col in frame.columns:
        canonical = default_by_lower.get(str(col).lower())
        if canonical and col != canonical:
            rename_map[col] = canonical
    if rename_map:
        frame = frame.rename(columns=rename_map)
    return frame


def safe_sheet_name(name: str) -> str:
    for char in ["\\", "/", "*", "[", "]", ":", "?"]:
        name = name.replace(char, "_")
    return name[:31]


def autosize_excel(writer: pd.ExcelWriter, sheet_name: str, frame: pd.DataFrame) -> None:
    worksheet = writer.sheets.get(sheet_name)
    if worksheet is None:
        return
    for idx, column in enumerate(frame.columns, start=1):
        values = frame.iloc[:, idx - 1].astype(str).head(200).tolist() if len(frame.columns) >= idx else []
        width = max([len(str(column))] + [len(str(value)) for value in values]) + 2
        worksheet.column_dimensions[worksheet.cell(row=1, column=idx).column_letter].width = min(width, 45)


@dataclass
class InputSlot:
    slot_id: str
    title: str
    kind: str
    color: str
    accent: str
    rule: str
    required_key: str
    preferred_header_row: Optional[int] = None
    expected_column_count: Optional[int] = None
    path: Optional[str] = None
    columns: List[str] = field(default_factory=list)
    header_row: Optional[int] = None
    row_count: int = 0
    selected_columns: Set[str] = field(default_factory=set)


class ColumnSelector(ttk.Frame):
    def __init__(self, master: tk.Widget, slot: InputSlot, on_change) -> None:
        super().__init__(master, style="Panel.TFrame")
        self.slot = slot
        self.on_change = on_change
        self.filtered_columns: List[str] = []
        self.syncing = False

        header = ttk.Frame(self, style="Panel.TFrame")
        header.pack(fill="x", padx=10, pady=(10, 4))
        ttk.Label(header, text=slot.title, foreground=slot.accent, font=("Malgun Gothic", 10, "bold"), style="Panel.TLabel").pack(side="left")
        self.count_label = ttk.Label(header, text="선택 0개", foreground=slot.accent, style="Panel.TLabel")
        self.count_label.pack(side="right")

        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", lambda *_: self.refresh())
        ttk.Entry(self, textvariable=self.search_var).pack(fill="x", padx=10, pady=(4, 6))

        actions = ttk.Frame(self, style="Panel.TFrame")
        actions.pack(fill="x", padx=10, pady=(0, 8))
        ttk.Button(actions, text="전체 선택", command=self.select_all_visible).pack(side="left", fill="x", expand=True)
        ttk.Button(actions, text="전체 해제", command=self.clear_all).pack(side="left", fill="x", expand=True, padx=(4, 0))

        list_frame = ttk.Frame(self, style="Panel.TFrame")
        list_frame.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        self.listbox = tk.Listbox(list_frame, selectmode=tk.MULTIPLE, activestyle="none", exportselection=False, font=("Malgun Gothic", 9), borderwidth=0, highlightthickness=1, highlightbackground=slot.color)
        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=self.listbox.yview)
        self.listbox.configure(yscrollcommand=scrollbar.set)
        self.listbox.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        self.listbox.bind("<<ListboxSelect>>", self.capture_selection)
        self.refresh()

    def capture_selection(self, _event=None) -> None:
        if self.syncing:
            return
        selected_indices = set(self.listbox.curselection())
        self.slot.selected_columns.difference_update(set(self.filtered_columns))
        for idx in selected_indices:
            if 0 <= idx < len(self.filtered_columns):
                self.slot.selected_columns.add(self.filtered_columns[idx])
        self.update_count()
        self.on_change()

    def refresh(self) -> None:
        query = self.search_var.get().strip().lower()
        self.filtered_columns = [col for col in self.slot.columns if not query or query in col.lower()]
        self.syncing = True
        self.listbox.delete(0, tk.END)
        for col in self.filtered_columns:
            self.listbox.insert(tk.END, col)
        for idx, col in enumerate(self.filtered_columns):
            if col in self.slot.selected_columns:
                self.listbox.selection_set(idx)
        self.syncing = False
        self.update_count()

    def select_all_visible(self) -> None:
        self.slot.selected_columns.update(self.filtered_columns)
        self.refresh()
        self.on_change()

    def clear_all(self) -> None:
        self.slot.selected_columns.clear()
        self.refresh()
        self.on_change()

    def update_count(self) -> None:
        self.count_label.configure(text=f"선택 {len(self.slot.selected_columns)}개")


class MergePlannerApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("1360x820")
        self.minsize(1180, 720)
        self.output_dir = tk.StringVar(value=str(Path.home() / "Desktop"))
        self.status_var = tk.StringVar(value="월확정서열 파일 1개 이상과 자재소요현황 파일을 선택하세요.")
        self.running = False
        self.slots: List[InputSlot] = [
            InputSlot("month_1", "월확정서열 통합1라인", "month", "#b7d3ff", "#2563eb", "선택 사항 · 생산번호 기준 · HI-SRM 헤더 2행/178컬럼", MONTH_KEY, HI_SRM_MONTH_HEADER_ROW, 178),
            InputSlot("month_2", "월확정서열 통합2라인", "month", "#f6b5b5", "#dc2626", "선택 사항 · 생산번호 기준 · HI-SRM 헤더 2행/163컬럼", MONTH_KEY, HI_SRM_MONTH_HEADER_ROW, 163),
            InputSlot("month_3", "월확정서열 선진정공", "month", "#f3cf70", "#d97706", "선택 사항 · 생산번호 기준 · HI-SRM 헤더 2행/119컬럼", MONTH_KEY, HI_SRM_MONTH_HEADER_ROW, 119),
            InputSlot("month_4", "월확정서열 초대형", "month", "#a8d9b5", "#16a34a", "선택 사항 · 생산번호 기준 · HI-SRM 헤더 2행/93컬럼", MONTH_KEY, HI_SRM_MONTH_HEADER_ROW, 93),
            InputSlot("material", "물류번호별 자재소요현황", "material", "#d7b7ff", "#9333ea", "필수 · 물류번호 기준 · HI-SRM 헤더 1행/33컬럼", MATERIAL_KEY, HI_SRM_MATERIAL_HEADER_ROW, 33),
        ]
        self.card_labels: Dict[str, Dict[str, ttk.Label]] = {}
        self.column_selectors: Dict[str, ColumnSelector] = {}
        self.summary_labels: Dict[str, ttk.Label] = {}
        self.configure_styles()
        self.build_ui()
        self.update_state()

    def configure_styles(self) -> None:
        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure(".", font=("Malgun Gothic", 10))
        style.configure("Root.TFrame", background="#f6f9ff")
        style.configure("Panel.TFrame", background="#ffffff")
        style.configure("Panel.TLabel", background="#ffffff")
        style.configure("Title.TLabel", background="#f6f9ff", foreground="#0f172a", font=("Malgun Gothic", 13, "bold"))
        style.configure("Metric.TLabel", background="#ffffff", foreground="#0f3ea5", font=("Malgun Gothic", 19, "bold"))
        style.configure("Subtle.TLabel", background="#ffffff", foreground="#475569")
        style.configure("Accent.TButton", font=("Malgun Gothic", 10, "bold"))

    def build_ui(self) -> None:
        container = ttk.Frame(self, style="Root.TFrame")
        container.pack(fill="both", expand=True)
        canvas = tk.Canvas(container, bg="#f6f9ff", highlightthickness=0)
        scrollbar = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
        self.content = ttk.Frame(canvas, style="Root.TFrame")
        self.content.bind("<Configure>", lambda _event: canvas.configure(scrollregion=canvas.bbox("all")))
        window = canvas.create_window((0, 0), window=self.content, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.bind("<Configure>", lambda event: canvas.itemconfigure(window, width=event.width))
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        ttk.Label(self.content, text="1. 월확정서열 조립라인 및 자재 파일 선택", style="Title.TLabel").pack(anchor="w", padx=22, pady=(16, 8))
        cards = ttk.Frame(self.content, style="Root.TFrame")
        cards.pack(fill="x", padx=22)
        for idx, slot in enumerate(self.slots, start=1):
            card = tk.Frame(cards, bg="#ffffff", highlightbackground=slot.color, highlightthickness=1)
            card.grid(row=0, column=idx - 1, sticky="nsew", padx=(0 if idx == 1 else 8, 0), ipadx=8, ipady=8)
            cards.columnconfigure(idx - 1, weight=1, uniform="cards")
            self.build_input_card(card, slot, idx)

        metrics = ttk.Frame(self.content, style="Root.TFrame")
        metrics.pack(fill="x", padx=22, pady=(10, 20))
        for idx, (key, title, value) in enumerate([
            ("upload", "업로드 현황", "월확정서열 0/4 ·\n자재 대기"),
            ("columns", "선택 컬럼", "0개"),
            ("merged", "정상병합", "생성 후 확인"),
            ("unmatched", "예외 수", "생성 후 확인"),
            ("result", "생성 결과", "대기"),
        ]):
            box = ttk.Frame(metrics, style="Panel.TFrame", padding=(18, 14))
            box.grid(row=0, column=idx, sticky="nsew", padx=(0 if idx == 0 else 8, 0))
            metrics.columnconfigure(idx, weight=1, uniform="metrics")
            ttk.Label(box, text=title, style="Subtle.TLabel").pack(anchor="w")
            label = ttk.Label(box, text=value, style="Metric.TLabel")
            label.pack(anchor="w", pady=(14, 0))
            self.summary_labels[key] = label

        header = ttk.Frame(self.content, style="Root.TFrame")
        header.pack(fill="x", padx=22, pady=(0, 8))
        ttk.Label(header, text="2. 출력 컬럼 선택", style="Title.TLabel").pack(side="left")
        ttk.Label(header, text="검색 후 필요한 컬럼을 복수 선택하세요. 병합 키는 내부적으로 자동 포함됩니다.", background="#f6f9ff", foreground="#64748b").pack(side="right")

        selectors = ttk.Frame(self.content, style="Root.TFrame")
        selectors.pack(fill="both", expand=True, padx=22)
        for idx, slot in enumerate(self.slots):
            panel = tk.Frame(selectors, bg="#ffffff", highlightbackground=slot.color, highlightthickness=1)
            panel.grid(row=0, column=idx, sticky="nsew", padx=(0 if idx == 0 else 8, 0))
            selectors.columnconfigure(idx, weight=1, uniform="selectors")
            selector = ColumnSelector(panel, slot, self.update_state)
            selector.pack(fill="both", expand=True)
            self.column_selectors[slot.slot_id] = selector

        footer = ttk.Frame(self.content, style="Root.TFrame")
        footer.pack(fill="x", padx=22, pady=22)
        ttk.Button(footer, text="1  폴더 지정", command=self.choose_output_dir, style="Accent.TButton").pack(side="left", padx=(0, 10), ipadx=16, ipady=8)
        ttk.Button(footer, text="2  병합 엑셀 생성", command=self.start_merge, style="Accent.TButton").pack(side="left", padx=(0, 10), ipadx=16, ipady=8)
        ttk.Button(footer, text="3  폴더 열기", command=self.open_output_dir).pack(side="left", padx=(0, 10), ipadx=14, ipady=8)
        ttk.Button(footer, text="초기화", command=self.reset_all).pack(side="left", ipadx=12, ipady=8)
        right = ttk.Frame(footer, style="Root.TFrame")
        right.pack(side="right", fill="x", expand=True)
        ttk.Label(right, textvariable=self.status_var, background="#f6f9ff", foreground="#0f172a", font=("Malgun Gothic", 10, "bold")).pack(anchor="e")
        ttk.Label(right, textvariable=self.output_dir, background="#f6f9ff", foreground="#64748b").pack(anchor="e", pady=(6, 0))

    def build_input_card(self, parent: tk.Frame, slot: InputSlot, index: int) -> None:
        header = tk.Frame(parent, bg="#ffffff")
        header.pack(fill="x", padx=8, pady=(8, 6))
        tk.Label(header, text=str(index), bg=slot.accent, fg="#ffffff", width=3, font=("Malgun Gothic", 10, "bold")).pack(side="left")
        tk.Label(header, text=slot.title, bg=slot.color, fg=slot.accent, anchor="w", font=("Malgun Gothic", 10, "bold")).pack(side="left", fill="x", expand=True, padx=(8, 0))
        ttk.Label(parent, text=slot.rule, style="Panel.TLabel", foreground="#475569").pack(anchor="w", padx=8, pady=(4, 0))
        path_label = ttk.Label(parent, text="파일을 선택하세요", style="Panel.TLabel", foreground=slot.accent)
        path_label.pack(anchor="w", padx=8, pady=(8, 0))
        count_label = ttk.Label(parent, text="헤더/데이터 건수 대기", style="Panel.TLabel", foreground="#475569")
        count_label.pack(anchor="w", padx=8, pady=(8, 0))
        bottom = ttk.Frame(parent, style="Panel.TFrame")
        bottom.pack(fill="x", padx=8, pady=(12, 8))
        status_label = ttk.Label(bottom, text="대기", style="Panel.TLabel", foreground="#334155")
        status_label.pack(side="left", ipadx=12, ipady=8)
        ttk.Button(bottom, text="파일 선택", command=lambda: self.choose_file(slot)).pack(side="right", ipadx=10, ipady=5)
        self.card_labels[slot.slot_id] = {"path": path_label, "count": count_label, "status": status_label}

    def choose_file(self, slot: InputSlot) -> None:
        path = filedialog.askopenfilename(title=f"{slot.title} 파일 선택", filetypes=EXCEL_FILETYPES)
        if not path:
            return
        if Path(path).name.startswith("~$"):
            messagebox.showwarning("임시 파일", "Excel 임시 파일(~$)은 선택할 수 없습니다. 원본 파일을 선택해 주세요.")
            return
        normalized = os.path.normcase(os.path.abspath(path))
        for other in self.slots:
            if other is not slot and other.path and os.path.normcase(os.path.abspath(other.path)) == normalized:
                messagebox.showerror("중복 파일", "동일 파일을 두 입력 영역에 중복 선택할 수 없습니다.")
                return
        try:
            frame, header_row = read_excel_frame(path, slot.required_key, slot.preferred_header_row, slot.expected_column_count)
            defaults = DEFAULT_MATERIAL_COLUMNS if slot.kind == "material" else DEFAULT_MONTH_COLUMNS
            frame = canonicalize_default_columns(frame, defaults)
        except Exception as exc:
            messagebox.showerror("파일 읽기 실패", f"{slot.title} 파일을 읽지 못했습니다.\n\n{exc}")
            return
        if slot.required_key not in frame.columns:
            messagebox.showwarning("병합 키 확인 필요", f"'{slot.required_key}' 컬럼을 찾지 못했습니다.\n파일 구조를 확인해 주세요.")
        slot.path = path
        slot.columns = list(frame.columns)
        slot.header_row = header_row + 1
        slot.row_count = len(frame)
        defaults = DEFAULT_MATERIAL_COLUMNS if slot.kind == "material" else DEFAULT_MONTH_COLUMNS
        slot.selected_columns = default_selected_columns(slot.columns, defaults)
        self.refresh_slot(slot)
        self.update_state()

    def refresh_slot(self, slot: InputSlot) -> None:
        labels = self.card_labels[slot.slot_id]
        if slot.path:
            labels["path"].configure(text=Path(slot.path).name)
            expected = f" · 기준 {slot.expected_column_count}컬럼" if slot.expected_column_count else ""
            labels["count"].configure(text=f"헤더 {slot.header_row}행 · 데이터 {slot.row_count:,}건{expected}")
            labels["status"].configure(text="선택 완료")
        else:
            labels["path"].configure(text="파일을 선택하세요")
            labels["count"].configure(text="헤더/데이터 건수 대기")
            labels["status"].configure(text="대기")
        self.column_selectors[slot.slot_id].refresh()

    def update_state(self) -> None:
        month_count = sum(1 for slot in self.slots if slot.kind == "month" and slot.path)
        material_ready = any(slot.kind == "material" and slot.path for slot in self.slots)
        total_selected = sum(len(slot.selected_columns) for slot in self.slots)
        self.summary_labels["upload"].configure(text=f"월확정서열 {month_count}/4 ·\n자재 {'완료' if material_ready else '대기'}")
        self.summary_labels["columns"].configure(text=f"{total_selected}개")
        if month_count >= 1 and material_ready:
            self.status_var.set("출력할 컬럼을 선택하세요." if total_selected == 0 else "병합 엑셀을 생성할 수 있습니다.")
        else:
            self.status_var.set("월확정서열 파일 1개 이상과 자재소요현황 파일을 선택하세요.")

    def choose_output_dir(self) -> None:
        path = filedialog.askdirectory(title="저장할 폴더 선택", initialdir=self.output_dir.get())
        if path:
            self.output_dir.set(path)

    def open_output_dir(self) -> None:
        path = self.output_dir.get()
        if not path or not Path(path).exists():
            messagebox.showwarning("폴더 없음", "먼저 저장 폴더를 지정해 주세요.")
            return
        os.startfile(path)

    def reset_all(self) -> None:
        for slot in self.slots:
            slot.path = None
            slot.columns.clear()
            slot.selected_columns.clear()
            slot.header_row = None
            slot.row_count = 0
            self.refresh_slot(slot)
        self.summary_labels["merged"].configure(text="생성 후 확인")
        self.summary_labels["unmatched"].configure(text="생성 후 확인")
        self.summary_labels["result"].configure(text="대기")
        self.update_state()

    def validate_before_merge(self) -> Optional[str]:
        if self.running:
            return "이미 생성 작업이 진행 중입니다."
        selected_months = [slot for slot in self.slots if slot.kind == "month" and slot.path]
        material = next(slot for slot in self.slots if slot.kind == "material")
        if not selected_months:
            return "월확정서열 파일을 1개 이상 선택해 주세요."
        if not material.path:
            return "물류번호별 자재소요현황 파일은 필수입니다."
        if not Path(self.output_dir.get()).exists():
            return "저장 폴더를 먼저 지정해 주세요."
        for slot in selected_months + [material]:
            if len(slot.selected_columns) == 0:
                return f"{slot.title} 영역에서 출력할 컬럼을 1개 이상 선택해 주세요."
        for slot in selected_months:
            if MONTH_KEY not in slot.columns:
                return f"{slot.title} 파일에서 '{MONTH_KEY}' 컬럼을 찾지 못했습니다."
        if MATERIAL_KEY not in material.columns:
            return f"{material.title} 파일에서 '{MATERIAL_KEY}' 컬럼을 찾지 못했습니다."
        return None

    def start_merge(self) -> None:
        error = self.validate_before_merge()
        if error:
            messagebox.showwarning("생성 불가", error)
            return
        self.running = True
        self.status_var.set("병합 엑셀을 생성하는 중입니다...")
        self.summary_labels["result"].configure(text="생성 중")
        threading.Thread(target=self.merge_worker, daemon=True).start()

    def merge_worker(self) -> None:
        try:
            result_path, stats = self.create_output_file()
        except Exception as exc:
            self.after(0, lambda: self.merge_failed(exc))
            return
        self.after(0, lambda: self.merge_finished(result_path, stats))

    def merge_failed(self, exc: Exception) -> None:
        self.running = False
        self.summary_labels["result"].configure(text="실패")
        self.status_var.set("생성 중 오류가 발생했습니다.")
        messagebox.showerror("생성 실패", str(exc))

    def merge_finished(self, result_path: str, stats: Dict[str, int]) -> None:
        self.running = False
        self.summary_labels["merged"].configure(text=f"{stats['merged']:,}건")
        self.summary_labels["unmatched"].configure(text=f"{stats['month_unmatched'] + stats['material_unmatched'] + stats.get('start_date_mismatch', 0):,}건")
        self.summary_labels["result"].configure(text="완료")
        self.status_var.set(f"생성 완료: {Path(result_path).name}")
        messagebox.showinfo("생성 완료", f"통합 엑셀 파일을 생성했습니다.\n\n{result_path}")

    def create_output_file(self) -> Tuple[str, Dict[str, int]]:
        month_slots = [slot for slot in self.slots if slot.kind == "month" and slot.path]
        material_slot = next(slot for slot in self.slots if slot.kind == "material")
        month_frames = []

        for slot in month_slots:
            frame, _ = read_excel_frame(slot.path or "", MONTH_KEY, slot.preferred_header_row, slot.expected_column_count)
            frame = canonicalize_default_columns(frame, DEFAULT_MONTH_COLUMNS)
            frame = normalize_start_date_columns(frame)
            frame[MONTH_KEY] = frame[MONTH_KEY].map(normalize_text)
            frame = frame[frame[MONTH_KEY] != ""].copy()
            month_start_col = find_start_date_column(frame.columns)
            frame[MONTH_START_DATE_INTERNAL] = frame[month_start_col] if month_start_col else ""
            frame["월확정_구분"] = slot.title
            columns = columns_with_required(frame.columns, slot.selected_columns, [MONTH_KEY, MONTH_START_DATE_INTERNAL])
            if "월확정_구분" not in columns:
                columns.append("월확정_구분")
            month_frames.append(frame[columns])
        month_all = pd.concat(month_frames, ignore_index=True)

        material_frame, _ = read_excel_frame(material_slot.path or "", MATERIAL_KEY, material_slot.preferred_header_row, material_slot.expected_column_count)
        material_frame = canonicalize_default_columns(material_frame, DEFAULT_MATERIAL_COLUMNS)
        material_frame = normalize_start_date_columns(material_frame)
        material_frame[MATERIAL_KEY] = material_frame[MATERIAL_KEY].map(normalize_text)
        material_frame = material_frame[material_frame[MATERIAL_KEY] != ""].copy()
        material_start_col = find_start_date_column(material_frame.columns)
        material_frame[MATERIAL_START_DATE_INTERNAL] = material_frame[material_start_col] if material_start_col else ""
        material_columns = columns_with_required(material_frame.columns, material_slot.selected_columns, [MATERIAL_KEY, MATERIAL_START_DATE_INTERNAL])
        material_output = material_frame[material_columns]

        merged_raw = material_output.merge(
            month_all,
            left_on=MATERIAL_KEY,
            right_on=MONTH_KEY,
            how="inner",
            suffixes=("_자재", "_월확정"),
        )
        month_keys = set(month_all[MONTH_KEY])
        material_keys = set(material_frame[MATERIAL_KEY])
        material_unmatched = material_output[~material_output[MATERIAL_KEY].isin(month_keys)].copy()
        month_unmatched = month_all[~month_all[MONTH_KEY].isin(material_keys)].copy()

        material_dates = merged_raw[MATERIAL_START_DATE_INTERNAL].astype(str) if MATERIAL_START_DATE_INTERNAL in merged_raw.columns else pd.Series([""] * len(merged_raw), index=merged_raw.index)
        month_dates = merged_raw[MONTH_START_DATE_INTERNAL].astype(str) if MONTH_START_DATE_INTERNAL in merged_raw.columns else pd.Series([""] * len(merged_raw), index=merged_raw.index)
        date_mismatch_mask = (material_dates != "") & (month_dates != "") & (material_dates != month_dates)
        start_date_mismatch = merged_raw[date_mismatch_mask].copy()
        if not start_date_mismatch.empty:
            start_date_mismatch["자재_착수일자"] = start_date_mismatch[MATERIAL_START_DATE_INTERNAL]
            start_date_mismatch["월확정_착수일자"] = start_date_mismatch[MONTH_START_DATE_INTERNAL]
            if MATERIAL_KEY in start_date_mismatch.columns:
                start_date_mismatch = start_date_mismatch.drop(columns=[MATERIAL_KEY])
            start_date_mismatch = start_date_mismatch.drop(
                columns=[col for col in (MONTH_START_DATE_INTERNAL, MATERIAL_START_DATE_INTERNAL) if col in start_date_mismatch.columns]
            )

        merged = collapse_key_and_start_date_columns(merged_raw[~date_mismatch_mask].copy())
        month_unmatched = collapse_key_and_start_date_columns(month_unmatched.copy())
        material_unmatched = collapse_key_and_start_date_columns(material_unmatched.copy())

        summary = pd.DataFrame([
            {"항목": "선택 월확정 파일 수", "값": len(month_slots)},
            {"항목": "월확정 전체 행 수", "값": len(month_all)},
            {"항목": "자재소요현황 전체 행 수", "값": len(material_frame)},
            {"항목": "정상병합 행 수", "값": len(merged)},
            {"항목": "착수일자 불일치 행 수", "값": len(start_date_mismatch)},
            {"항목": "월확정 미매칭 행 수", "값": len(month_unmatched)},
            {"항목": "자재 미매칭 행 수", "값": len(material_unmatched)},
            {"항목": "생성일시", "값": datetime.now().strftime("%Y-%m-%d %H:%M:%S")},
        ])
        output_path = str(Path(self.output_dir.get()) / f"{OUTPUT_PREFIX}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
        sheets = {
            "정상병합": merged,
            "착수일자_불일치": start_date_mismatch,
            "월확정_미매칭": month_unmatched,
            "자재_미매칭": material_unmatched,
            "실행요약": summary,
        }
        with pd.ExcelWriter(excel_read_path(output_path), engine="openpyxl") as writer:
            for sheet_name, frame in sheets.items():
                safe_name = safe_sheet_name(sheet_name)
                frame.to_excel(writer, sheet_name=safe_name, index=False)
                autosize_excel(writer, safe_name, frame)
        return output_path, {
            "merged": len(merged),
            "month_unmatched": len(month_unmatched),
            "material_unmatched": len(material_unmatched),
            "start_date_mismatch": len(start_date_mismatch),
        }


def main() -> None:
    app = MergePlannerApp()
    app.mainloop()


if __name__ == "__main__":
    main()
