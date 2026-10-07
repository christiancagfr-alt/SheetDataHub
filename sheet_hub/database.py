from __future__ import annotations

import hashlib
import json
import re
import sqlite3
from datetime import datetime
from pathlib import Path

from .models import Record


class AggregateDatabase:
    def __init__(self, directory: str | Path, max_rows: int = 500000):
        self.directory = Path(directory)
        self.directory.mkdir(parents=True, exist_ok=True)
        self.max_rows = max(1000, int(max_rows))

    def paths(self) -> list[Path]:
        return sorted(
            path for path in self.directory.glob("aggregate_*.sqlite")
            if path.stem.removeprefix("aggregate_").isdigit()
        )

    def _path(self, number: int) -> Path:
        return self.directory / f"aggregate_{number:03d}.sqlite"

    @staticmethod
    def _init(conn: sqlite3.Connection) -> None:
        conn.executescript(
            """
            CREATE TABLE IF NOT EXISTS records (
                row_hash TEXT PRIMARY KEY,
                source_id TEXT NOT NULL,
                source_name TEXT NOT NULL,
                spreadsheet_id TEXT NOT NULL,
                sheet_name TEXT NOT NULL,
                row_number INTEGER NOT NULL,
                payload TEXT NOT NULL,
                updated_at TEXT NOT NULL DEFAULT (datetime('now','localtime'))
            );
            CREATE INDEX IF NOT EXISTS ix_records_source ON records(source_id, sheet_name);
            """
        )

    def replace_all(self, records: list[Record]) -> tuple[int, int]:
        temp_paths: list[Path] = []
        try:
            chunks = [records[i:i + self.max_rows] for i in range(0, len(records), self.max_rows)] or [[]]
            for number, chunk in enumerate(chunks, start=1):
                temp = self.directory / f"aggregate_{number:03d}.tmp.sqlite"
                if temp.exists():
                    temp.unlink()
                conn = sqlite3.connect(temp)
                try:
                    self._init(conn)
                    conn.executemany(
                        "INSERT OR REPLACE INTO records(row_hash,source_id,source_name,spreadsheet_id,sheet_name,row_number,payload) VALUES(?,?,?,?,?,?,?)",
                        [
                            (r.row_hash, r.source_id, r.source_name, r.spreadsheet_id, r.sheet_name, r.row_number,
                             json.dumps(r.values, ensure_ascii=False))
                            for r in chunk
                        ],
                    )
                    conn.commit()
                finally:
                    conn.close()
                temp_paths.append(temp)
            for old in self.paths():
                old.unlink()
            for number, temp in enumerate(temp_paths, start=1):
                temp.replace(self._path(number))
            return len(records), len(temp_paths)
        except Exception:
            for temp in temp_paths:
                if temp.exists():
                    temp.unlink()
            raise

    def all_records(self) -> list[Record]:
        result: list[Record] = []
        for path in self.paths():
            conn = sqlite3.connect(path)
            try:
                rows = conn.execute(
                    "SELECT source_id,source_name,spreadsheet_id,sheet_name,row_number,payload,row_hash FROM records"
                ).fetchall()
            finally:
                conn.close()
            result.extend(Record(*row[:5], json.loads(row[5]), row[6]) for row in rows)
        return result

    def sample_records(self, limit: int = 20) -> list[Record]:
        remaining = max(1, int(limit))
        result: list[Record] = []
        for path in self.paths():
            conn = sqlite3.connect(path)
            try:
                rows = conn.execute(
                    "SELECT source_id,source_name,spreadsheet_id,sheet_name,row_number,payload,row_hash FROM records LIMIT ?",
                    (remaining,),
                ).fetchall()
            finally:
                conn.close()
            result.extend(Record(*row[:5], json.loads(row[5]), row[6]) for row in rows)
            remaining -= len(rows)
            if remaining <= 0:
                break
        return result

    def query(self, field: str, value: str, exact: bool = True) -> list[Record]:
        needle = value.strip().casefold()
        result: list[Record] = []
        for record in self.all_records():
            haystack = record.values.get(field, "").strip().casefold()
            if (haystack == needle) if exact else (needle in haystack):
                result.append(record)
        return result


class SourceCache:
    def __init__(self, directory: str | Path):
        self.directory = Path(directory)
        self.directory.mkdir(parents=True, exist_ok=True)

    @staticmethod
    def _safe_id(source_id: str) -> str:
        return re.sub(r"[^a-zA-Z0-9_-]", "_", str(source_id or "source"))[:80] or "source"

    def path(self, source_id: str) -> Path:
        return self.directory / f"{self._safe_id(source_id)}.sqlite"

    @staticmethod
    def _prepare(conn: sqlite3.Connection) -> None:
        conn.execute("PRAGMA journal_mode=WAL")
        conn.execute("PRAGMA synchronous=NORMAL")
        conn.execute("PRAGMA temp_store=MEMORY")

    def has(self, source_id: str) -> bool:
        path = self.path(source_id)
        return path.exists() and path.stat().st_size > 0

    def replace(self, source_id: str, records: list[Record], headers: list[str] | None = None, target: str = "") -> int:
        path = self.path(source_id)
        temp = path.with_suffix(".tmp.sqlite")
        if temp.exists():
            temp.unlink()
        conn = sqlite3.connect(temp)
        try:
            self._prepare(conn)
            conn.executescript(
                """
                CREATE TABLE records (
                    row_hash TEXT PRIMARY KEY,
                    source_id TEXT NOT NULL,
                    source_name TEXT NOT NULL,
                    spreadsheet_id TEXT NOT NULL,
                    sheet_name TEXT NOT NULL,
                    row_number INTEGER NOT NULL,
                    payload TEXT NOT NULL
                );
                CREATE TABLE meta (
                    key TEXT PRIMARY KEY,
                    value TEXT NOT NULL
                );
                CREATE INDEX ix_records_sheet ON records(sheet_name);
                """
            )
            conn.executemany(
                "INSERT OR REPLACE INTO records(row_hash,source_id,source_name,spreadsheet_id,sheet_name,row_number,payload) VALUES(?,?,?,?,?,?,?)",
                [
                    (r.row_hash, r.source_id, r.source_name, r.spreadsheet_id, r.sheet_name, r.row_number,
                     json.dumps(r.values, ensure_ascii=False))
                    for r in records
                ],
            )
            keys = list(headers or [])
            if not keys:
                seen: list[str] = []
                for record in records[:50]:
                    for key in record.values:
                        if key not in seen:
                            seen.append(key)
                keys = seen
            conn.execute("INSERT OR REPLACE INTO meta(key,value) VALUES(?,?)", ("headers", json.dumps(keys, ensure_ascii=False)))
            conn.execute("INSERT OR REPLACE INTO meta(key,value) VALUES(?,?)", ("updated_at", datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
            conn.execute("INSERT OR REPLACE INTO meta(key,value) VALUES(?,?)", ("rows", str(len(records))))
            conn.execute("INSERT OR REPLACE INTO meta(key,value) VALUES(?,?)", ("fingerprint", self.content_fingerprint(records)))
            if target:
                conn.execute("INSERT OR REPLACE INTO meta(key,value) VALUES(?,?)", ("target", str(target)))
            conn.commit()
        finally:
            conn.close()
        if path.exists():
            path.unlink()
        temp.replace(path)
        return len(records)

    def load(self, source_id: str) -> list[Record]:
        if not self.has(source_id):
            return []
        conn = sqlite3.connect(self.path(source_id))
        try:
            self._prepare(conn)
            rows = conn.execute(
                "SELECT source_id,source_name,spreadsheet_id,sheet_name,row_number,payload,row_hash FROM records"
            ).fetchall()
        finally:
            conn.close()
        return [Record(*row[:5], json.loads(row[5]), row[6]) for row in rows]

    def headers(self, source_id: str) -> list[str]:
        return [str(item) for item in json.loads(self._meta(source_id, "headers") or "[]")]

    @staticmethod
    def content_fingerprint(records: list[Record]) -> str:
        digest = hashlib.sha256()
        for record in sorted(records, key=lambda item: (item.sheet_name, item.row_number, item.row_hash)):
            digest.update(f"{record.sheet_name}\0{record.row_number}\0{record.row_hash}\n".encode("utf-8"))
        return digest.hexdigest()

    def merge(
        self,
        source_id: str,
        records: list[Record],
        headers: list[str] | None = None,
        target: str = "",
    ) -> dict[str, int | bool]:
        new_fp = self.content_fingerprint(records)
        existed = self.has(source_id)
        if existed and self._meta(source_id, "fingerprint") == new_fp:
            return {
                "rows": len(records),
                "inserted": 0,
                "updated": 0,
                "deleted": 0,
                "skipped": len(records),
                "unchanged": True,
            }
        self.replace(source_id, records, headers, target)
        return {
            "rows": len(records),
            "inserted": 0 if existed else len(records),
            "updated": len(records) if existed else 0,
            "deleted": 0,
            "skipped": 0,
            "unchanged": False,
        }

    def updated_at(self, source_id: str) -> str:
        return self._meta(source_id, "updated_at")

    def row_count(self, source_id: str) -> int:
        try:
            return int(self._meta(source_id, "rows") or 0)
        except ValueError:
            return 0

    def _meta(self, source_id: str, key: str) -> str:
        if not self.has(source_id):
            return ""
        conn = sqlite3.connect(self.path(source_id))
        try:
            row = conn.execute("SELECT value FROM meta WHERE key=?", (key,)).fetchone()
        except sqlite3.Error:
            return ""
        finally:
            conn.close()
        return str(row[0]) if row else ""
