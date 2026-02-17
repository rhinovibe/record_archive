import json
import os
import shutil
import re
import sqlite3
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from tkinter import (
    BOTH,
    END,
    LEFT,
    RIGHT,
    VERTICAL,
    X,
    Y,
    BooleanVar,
    Button,
    Checkbutton,
    Entry,
    Frame,
    Label,
    LabelFrame,
    Listbox,
    EXTENDED,
    Menu,
    Message,
    Scrollbar,
    StringVar,
    Text,
    Tk,
    filedialog,
    messagebox,
)

from importlib import import_module
from importlib.util import find_spec

if find_spec("docx") is not None:
    Document = import_module("docx").Document
else:
    Document = None


APP_DIR = Path(__file__).parent
DEFAULT_EXPORT_PATH = r"D:\기록 프로젝트"
CONFIG_PATH = APP_DIR / "config.json"


def resolve_db_path() -> Path:
    preferred = Path(DEFAULT_EXPORT_PATH) / "records.db"
    legacy = APP_DIR / "records.db"
    home_fallback = Path.home() / "record_archive" / "records.db"

    # Windows에서는 D: 우선, 실패 시 사용자 홈 경로로 폴백
    if os.name == "nt":
        for candidate in (preferred, home_fallback, legacy):
            try:
                candidate.parent.mkdir(parents=True, exist_ok=True)
                if candidate == preferred and not candidate.exists() and legacy.exists():
                    shutil.copy2(legacy, candidate)
                return candidate
            except Exception:
                continue
        return legacy

    # 비-Windows 개발 환경에서는 저장소 로컬 DB 사용
    return legacy


DB_PATH = resolve_db_path()


def normalize_filename(name: str) -> str:
    return re.sub(r'[\\/:*?"<>|]', "_", name).strip()[:120] or "untitled"


@dataclass
class RecordSummary:
    record_id: int
    title: str
    created_at: str


class RecordDB:
    def __init__(self, db_path: Path):
        self.db_path = db_path
        db_path.parent.mkdir(parents=True, exist_ok=True)
        self.conn = sqlite3.connect(db_path, timeout=10)
        self.conn.row_factory = sqlite3.Row
        self.conn.execute("PRAGMA journal_mode=WAL")
        self.conn.execute("PRAGMA synchronous=FULL")
        self.conn.execute("PRAGMA busy_timeout=10000")
        self._init_schema()

    def _init_schema(self):
        cur = self.conn.cursor()
        cur.execute(
            """
            CREATE TABLE IF NOT EXISTS records (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                title TEXT NOT NULL,
                body TEXT NOT NULL,
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL
            )
            """
        )
        cur.execute(
            """
            CREATE TABLE IF NOT EXISTS attributes (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL,
                level INTEGER NOT NULL,
                parent_id INTEGER,
                UNIQUE(name, level, parent_id),
                FOREIGN KEY(parent_id) REFERENCES attributes(id)
            )
            """
        )
        cur.execute(
            """
            CREATE TABLE IF NOT EXISTS record_attributes (
                record_id INTEGER NOT NULL,
                attribute_id INTEGER NOT NULL,
                PRIMARY KEY(record_id, attribute_id),
                FOREIGN KEY(record_id) REFERENCES records(id),
                FOREIGN KEY(attribute_id) REFERENCES attributes(id)
            )
            """
        )
        self.conn.commit()

    def get_attributes_by_parent(self, parent_id=None):
        if parent_id is None:
            return self.conn.execute(
                "SELECT * FROM attributes WHERE level = 1 ORDER BY name"
            ).fetchall()
        parent = self.conn.execute("SELECT * FROM attributes WHERE id = ?", (parent_id,)).fetchone()
        if not parent:
            return []
        return self.conn.execute(
            "SELECT * FROM attributes WHERE parent_id = ? ORDER BY name", (parent_id,)
        ).fetchall()

    def find_or_create_attribute(self, name: str, level: int, parent_id=None):
        name = name.strip()
        row = self.conn.execute(
            "SELECT * FROM attributes WHERE name = ? AND level = ? AND parent_id IS ?",
            (name, level, parent_id),
        ).fetchone()
        if row:
            return row["id"]
        cur = self.conn.cursor()
        cur.execute(
            "INSERT INTO attributes(name, level, parent_id) VALUES (?, ?, ?)",
            (name, level, parent_id),
        )
        self.conn.commit()
        return cur.lastrowid

    def save_record(self, title: str, body: str, created_at: str, attribute_ids: list[int], record_id=None):
        now = datetime.now().isoformat(timespec="seconds")
        with self.conn:
            cur = self.conn.cursor()
            if record_id:
                cur.execute(
                    "UPDATE records SET title=?, body=?, updated_at=? WHERE id=?",
                    (title, body, now, record_id),
                )
                rid = record_id
                cur.execute("DELETE FROM record_attributes WHERE record_id = ?", (rid,))
            else:
                cur.execute(
                    "INSERT INTO records(title, body, created_at, updated_at) VALUES (?, ?, ?, ?)",
                    (title, body, created_at, now),
                )
                rid = cur.lastrowid
            for attr_id in attribute_ids:
                cur.execute(
                    "INSERT OR IGNORE INTO record_attributes(record_id, attribute_id) VALUES (?, ?)",
                    (rid, attr_id),
                )

        self.conn.execute("PRAGMA wal_checkpoint(FULL)")
        return rid

    def list_records(self):
        rows = self.conn.execute(
            "SELECT id, title, created_at FROM records ORDER BY created_at DESC"
        ).fetchall()
        return [RecordSummary(r["id"], r["title"], r["created_at"]) for r in rows]

    def get_record(self, record_id: int):
        row = self.conn.execute("SELECT * FROM records WHERE id = ?", (record_id,)).fetchone()
        attrs = self.conn.execute(
            """
            SELECT a.* FROM attributes a
            JOIN record_attributes ra ON a.id = ra.attribute_id
            WHERE ra.record_id = ?
            ORDER BY a.level, a.name
            """,
            (record_id,),
        ).fetchall()
        return row, attrs

    def search_records(self, query: str, use_attr: bool, use_date: bool, use_title: bool, use_body: bool):
        normalized = query.strip()
        if not normalized:
            return self.list_records()

        tokens = [token for token in normalized.split() if token]
        if not tokens:
            return self.list_records()

        ids = set()

        if use_title or use_body or use_date:
            field_clauses = []
            if use_title:
                field_clauses.append("title LIKE ? COLLATE NOCASE")
            if use_body:
                field_clauses.append("body LIKE ? COLLATE NOCASE")
            if use_date:
                field_clauses.append("created_at LIKE ?")

            if field_clauses:
                token_clauses = []
                params = []
                for token in tokens:
                    token_clauses.append("(" + " OR ".join(field_clauses) + ")")
                    q = f"%{token}%"
                    params.extend([q] * len(field_clauses))
                where = " AND ".join(token_clauses)
                rows = self.conn.execute(
                    f"SELECT id FROM records WHERE {where}",
                    params,
                ).fetchall()
                ids.update(r["id"] for r in rows)

        if use_attr:
            attr_params = []
            attr_where = []
            for token in tokens:
                attr_where.append("a.name LIKE ? COLLATE NOCASE")
                attr_params.append(f"%{token}%")
            rows = self.conn.execute(
                f"""
                SELECT DISTINCT ra.record_id as id
                FROM record_attributes ra
                JOIN attributes a ON ra.attribute_id = a.id
                WHERE {' AND '.join(attr_where)}
                """,
                tuple(attr_params),
            ).fetchall()
            ids.update(r["id"] for r in rows)

        if not ids:
            return []

        placeholders = ",".join(["?"] * len(ids))
        rows = self.conn.execute(
            f"SELECT id, title, created_at FROM records WHERE id IN ({placeholders}) ORDER BY created_at DESC",
            tuple(ids),
        ).fetchall()
        return [RecordSummary(r["id"], r["title"], r["created_at"]) for r in rows]


class RecordApp:
    def __init__(self, root: Tk):
        self.root = root
        self.root.title("기록 관리 프로그램")
        self.root.geometry("1100x700")
        self.db = RecordDB(DB_PATH)
        self.export_path = self._load_export_path()
        self.current_record_id = None

        self.search_query = StringVar()
        self.search_attr = BooleanVar(value=True)
        self.search_date = BooleanVar(value=True)
        self.search_title = BooleanVar(value=True)
        self.search_body = BooleanVar(value=True)

        self._build_ui()
        self.refresh_records()

    def _load_export_path(self) -> str:
        if CONFIG_PATH.exists():
            try:
                data = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
                return data.get("export_path", DEFAULT_EXPORT_PATH)
            except Exception:
                pass
        return DEFAULT_EXPORT_PATH

    def _save_export_path(self):
        CONFIG_PATH.write_text(
            json.dumps({"export_path": self.export_path}, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )

    def _build_ui(self):
        menu = Menu(self.root)
        settings = Menu(menu, tearoff=False)
        settings.add_command(label="워드 저장 경로 변경", command=self.change_export_path)
        menu.add_cascade(label="설정", menu=settings)
        self.root.config(menu=menu)

        left = Frame(self.root)
        left.pack(side=LEFT, fill=BOTH, expand=False, padx=10, pady=10)

        search_box = LabelFrame(left, text="검색")
        search_box.pack(fill=X)
        Entry(search_box, textvariable=self.search_query, width=35).pack(side=LEFT, padx=5, pady=5)
        Button(search_box, text="검색", command=self.perform_search).pack(side=LEFT, padx=5)
        Button(search_box, text="초기화", command=self.refresh_records).pack(side=LEFT, padx=5)

        options = Frame(left)
        options.pack(fill=X, pady=(5, 0))
        Checkbutton(options, text="Attribute", variable=self.search_attr).pack(side=LEFT)
        Checkbutton(options, text="날짜", variable=self.search_date).pack(side=LEFT)
        Checkbutton(options, text="제목", variable=self.search_title).pack(side=LEFT)
        Checkbutton(options, text="본문", variable=self.search_body).pack(side=LEFT)

        self.record_list = Listbox(left, width=45, height=32)
        self.record_list.pack(side=LEFT, fill=Y, pady=10)
        self.record_list.bind("<<ListboxSelect>>", self.on_select_record)

        scroll = Scrollbar(left, orient=VERTICAL, command=self.record_list.yview)
        scroll.pack(side=RIGHT, fill=Y)
        self.record_list.configure(yscrollcommand=scroll.set)

        right = Frame(self.root)
        right.pack(side=LEFT, fill=BOTH, expand=True, padx=10, pady=10)

        header = Frame(right)
        header.pack(fill=X)
        Button(header, text="새 기록", command=self.new_record).pack(side=LEFT, padx=4)
        Button(header, text="저장", command=self.save_record).pack(side=LEFT, padx=4)
        Button(header, text="선택 기록 편집", command=self.load_selected_for_edit).pack(side=LEFT, padx=4)

        Label(right, text="제목").pack(anchor="w")
        self.title_entry = Entry(right)
        self.title_entry.pack(fill=X, pady=(0, 6))

        Label(right, text="본문").pack(anchor="w")
        self.body_text = Text(right, height=18)
        self.body_text.pack(fill=BOTH, expand=True)

        attr = LabelFrame(right, text="Attribute 지정")
        attr.pack(fill=X, pady=8)

        Label(attr, text="1차 Attribute (필수)").grid(row=0, column=0, sticky="w", padx=5)
        self.primary_list = Listbox(attr, height=6, exportselection=False, selectmode=EXTENDED)
        self.primary_list.grid(row=1, column=0, padx=5)
        self.primary_list.bind("<<ListboxSelect>>", self.on_primary_change)
        self.primary_new = Entry(attr)
        self.primary_new.grid(row=2, column=0, padx=5, pady=4)

        Label(attr, text="2차 Attribute (선택)").grid(row=0, column=1, sticky="w", padx=5)
        self.secondary_list = Listbox(attr, height=6, exportselection=False, selectmode=EXTENDED)
        self.secondary_list.grid(row=1, column=1, padx=5)
        self.secondary_list.bind("<<ListboxSelect>>", self.on_secondary_change)
        self.secondary_new = Entry(attr)
        self.secondary_new.grid(row=2, column=1, padx=5, pady=4)

        Label(attr, text="3차 Attribute (선택)").grid(row=0, column=2, sticky="w", padx=5)
        self.tertiary_list = Listbox(attr, height=6, exportselection=False, selectmode=EXTENDED)
        self.tertiary_list.grid(row=1, column=2, padx=5)
        self.tertiary_new = Entry(attr)
        self.tertiary_new.grid(row=2, column=2, padx=5, pady=4)

        self.status = Message(
            right,
            width=700,
            text=f"DB 경로: {DB_PATH} | 워드 저장 경로: {self.export_path}",
        )
        self.status.pack(fill=X)

        self.primary_map = {}
        self.secondary_map = {}
        self.tertiary_map = {}
        self.populate_primary_attributes()

    def change_export_path(self):
        selected = filedialog.askdirectory(title="워드 저장 경로 선택")
        if selected:
            self.export_path = selected
            self._save_export_path()
            self.status.configure(text=f"DB 경로: {DB_PATH} | 워드 저장 경로: {self.export_path}")

    def populate_primary_attributes(self):
        self.primary_list.delete(0, END)
        self.primary_map.clear()
        for idx, row in enumerate(self.db.get_attributes_by_parent(None)):
            self.primary_list.insert(END, row["name"])
            self.primary_map[idx] = row["id"]

    def on_primary_change(self, _event=None):
        self.secondary_list.delete(0, END)
        self.secondary_map.clear()
        self.tertiary_list.delete(0, END)
        self.tertiary_map.clear()

        selections = self.primary_list.curselection()
        if not selections:
            return

        idx = 0
        seen = set()
        for selected_idx in selections:
            pid = self.primary_map.get(selected_idx)
            primary_name = self.primary_list.get(selected_idx)
            for row in self.db.get_attributes_by_parent(pid):
                if row["id"] in seen:
                    continue
                seen.add(row["id"])
                self.secondary_list.insert(END, f"{primary_name} > {row['name']}")
                self.secondary_map[idx] = row["id"]
                idx += 1

    def on_secondary_change(self, _event=None):
        self.tertiary_list.delete(0, END)
        self.tertiary_map.clear()

        selections = self.secondary_list.curselection()
        if not selections:
            return

        idx = 0
        seen = set()
        for selected_idx in selections:
            sid = self.secondary_map.get(selected_idx)
            secondary_label = self.secondary_list.get(selected_idx)
            for row in self.db.get_attributes_by_parent(sid):
                if row["id"] in seen:
                    continue
                seen.add(row["id"])
                self.tertiary_list.insert(END, f"{secondary_label} > {row['name']}")
                self.tertiary_map[idx] = row["id"]
                idx += 1

    def refresh_records(self):
        self._render_record_list(self.db.list_records())

    def _render_record_list(self, records: list[RecordSummary]):
        self.record_list.delete(0, END)
        self._record_index = {}
        for i, rec in enumerate(records):
            self.record_list.insert(END, f"[{rec.created_at[:10]}] {rec.title}")
            self._record_index[i] = rec.record_id

    def perform_search(self):
        any_selected = any(
            [self.search_attr.get(), self.search_date.get(), self.search_title.get(), self.search_body.get()]
        )
        use_all = not any_selected
        records = self.db.search_records(
            self.search_query.get(),
            use_attr=self.search_attr.get() or use_all,
            use_date=self.search_date.get() or use_all,
            use_title=self.search_title.get() or use_all,
            use_body=self.search_body.get() or use_all,
        )
        self._render_record_list(records)

    def on_select_record(self, _event=None):
        sel = self.record_list.curselection()
        if not sel:
            return
        rid = self._record_index.get(sel[0])
        row, attrs = self.db.get_record(rid)
        if not row:
            return
        attr_txt = ", ".join([f"{a['level']}차:{a['name']}" for a in attrs]) if attrs else "없음"
        self.status.configure(
            text=f"선택 기록 ID {rid} | 생성일: {row['created_at']} | Attribute: {attr_txt}"
        )

    def load_selected_for_edit(self):
        sel = self.record_list.curselection()
        if not sel:
            messagebox.showwarning("안내", "편집할 기록을 선택하세요.")
            return
        rid = self._record_index.get(sel[0])
        row, attrs = self.db.get_record(rid)
        if not row:
            return
        self.current_record_id = rid
        self.title_entry.delete(0, END)
        self.title_entry.insert(0, row["title"])
        self.body_text.delete("1.0", END)
        self.body_text.insert("1.0", row["body"])

        self.primary_list.selection_clear(0, END)
        self.secondary_list.selection_clear(0, END)
        self.tertiary_list.selection_clear(0, END)

        level1_ids = {a["id"] for a in attrs if a["level"] == 1}
        level2_ids = {a["id"] for a in attrs if a["level"] == 2}
        level3_ids = {a["id"] for a in attrs if a["level"] == 3}

        if level1_ids:
            for idx, aid in self.primary_map.items():
                if aid in level1_ids:
                    self.primary_list.selection_set(idx)
            self.on_primary_change()
        if level2_ids:
            for idx, aid in self.secondary_map.items():
                if aid in level2_ids:
                    self.secondary_list.selection_set(idx)
            self.on_secondary_change()
        if level3_ids:
            for idx, aid in self.tertiary_map.items():
                if aid in level3_ids:
                    self.tertiary_list.selection_set(idx)

    def new_record(self):
        self.current_record_id = None
        self.title_entry.delete(0, END)
        self.body_text.delete("1.0", END)
        self.primary_list.selection_clear(0, END)
        self.secondary_list.delete(0, END)
        self.secondary_list.selection_clear(0, END)
        self.tertiary_list.delete(0, END)
        self.tertiary_list.selection_clear(0, END)
        self.primary_new.delete(0, END)
        self.secondary_new.delete(0, END)
        self.tertiary_new.delete(0, END)

    @staticmethod
    def _parse_new_attribute_names(raw: str):
        return [name.strip() for name in raw.split(",") if name.strip()]

    def _resolve_attribute_ids(self):
        selected_primary = self.primary_list.curselection()
        primary_names = self._parse_new_attribute_names(self.primary_new.get().strip())
        if not selected_primary and not primary_names:
            raise ValueError("1차 Attribute는 반드시 지정해야 합니다.")

        primary_ids = {self.primary_map[idx] for idx in selected_primary}
        for primary_name in primary_names:
            primary_ids.add(self.db.find_or_create_attribute(primary_name, 1, None))

        if primary_names:
            self.populate_primary_attributes()

        selected_secondary = self.secondary_list.curselection()
        secondary_names = self._parse_new_attribute_names(self.secondary_new.get().strip())
        secondary_ids = {self.secondary_map[idx] for idx in selected_secondary}

        for secondary_name in secondary_names:
            for p_id in primary_ids:
                secondary_ids.add(self.db.find_or_create_attribute(secondary_name, 2, p_id))

        if secondary_names:
            self.on_primary_change()

        selected_tertiary = self.tertiary_list.curselection()
        tertiary_names = self._parse_new_attribute_names(self.tertiary_new.get().strip())
        tertiary_ids = {self.tertiary_map[idx] for idx in selected_tertiary}

        for tertiary_name in tertiary_names:
            for s_id in secondary_ids:
                tertiary_ids.add(self.db.find_or_create_attribute(tertiary_name, 3, s_id))

        if tertiary_names:
            self.on_secondary_change()

        return list(primary_ids | secondary_ids | tertiary_ids)

    def save_record(self):
        title = self.title_entry.get().strip()
        body = self.body_text.get("1.0", END).strip()
        if not title or not body:
            messagebox.showwarning("안내", "제목과 본문을 입력하세요.")
            return
        try:
            attribute_ids = self._resolve_attribute_ids()
        except ValueError as e:
            messagebox.showwarning("안내", str(e))
            return

        created_at = datetime.now().isoformat(timespec="seconds")
        if self.current_record_id:
            existing, _ = self.db.get_record(self.current_record_id)
            if existing:
                created_at = existing["created_at"]

        try:
            rid = self.db.save_record(title, body, created_at, attribute_ids, self.current_record_id)
            persisted, _ = self.db.get_record(rid)
            if not persisted:
                raise RuntimeError("저장 직후 DB에서 레코드를 다시 찾지 못했습니다.")
        except Exception as exc:
            messagebox.showerror(
                "저장 실패",
                f"DB 저장 실패: {exc}\n\n현재 DB 경로: {self.db.db_path}",
            )
            self.status.configure(text=f"저장 실패 | DB 경로: {self.db.db_path} | 오류: {exc}")
            return

        self.current_record_id = rid
        self.refresh_records()
        self._export_to_word(title, body, created_at)
        messagebox.showinfo("완료", f"기록이 저장되었습니다.\nDB: {self.db.db_path}")

    def _export_to_word(self, title: str, body: str, created_at: str):
        if Document is None:
            self.status.configure(text="python-docx가 없어 워드 저장을 건너뜀")
            return
        try:
            date_folder = created_at[:10]
            target_dir = Path(self.export_path) / date_folder
            target_dir.mkdir(parents=True, exist_ok=True)
            filename = normalize_filename(f"{date_folder}: {title}") + ".docx"
            doc = Document()
            doc.add_heading(title, level=1)
            doc.add_paragraph(f"기록일시: {created_at}")
            doc.add_paragraph(body)
            doc.save(target_dir / filename)
            self.status.configure(text=f"워드 저장 완료: {target_dir / filename}")
        except Exception as exc:
            self.status.configure(text=f"워드 저장 실패(기록 본문 저장은 완료): {exc}")


def main():
    root = Tk()
    app = RecordApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
