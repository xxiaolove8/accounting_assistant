"""Bounded-memory accounting checks and summaries for filtered Excel rows."""

from __future__ import annotations

import json
import sqlite3
import tempfile
from collections import OrderedDict
from dataclasses import dataclass
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Iterator, Sequence


_ISSUE_LIMIT = 10_000
_GROUP_CACHE_LIMIT = 2_048
_COMMIT_INTERVAL = 2_000


@dataclass(frozen=True)
class AccountingConfig:
    """One-based columns to total, group, require, or check for duplicates."""

    amount_columns: tuple[int, ...] = ()
    group_column: int | None = None
    required_columns: tuple[int, ...] = ()
    duplicate_columns: tuple[int, ...] = ()

    def __post_init__(self) -> None:
        for name in ("amount_columns", "required_columns", "duplicate_columns"):
            columns = getattr(self, name)
            if any(not isinstance(col, int) or isinstance(col, bool) or col < 1 for col in columns):
                raise ValueError(f"{name} 中的列号必须是正整数")
            if len(set(columns)) != len(columns):
                raise ValueError(f"{name} 不能包含重复列号")
        if self.group_column is not None and (
            not isinstance(self.group_column, int)
            or isinstance(self.group_column, bool)
            or self.group_column < 1
        ):
            raise ValueError("group_column 必须是正整数或 None")


@dataclass
class _GroupState:
    count: int
    amounts: list[Decimal]


class AccountingAccumulator:
    """Accumulate exact totals and keep duplicate keys in a temporary SQLite file."""

    def __init__(self, config: AccountingConfig) -> None:
        self.config = config
        self.row_count = 0
        self.issue_count = 0
        self._issues: list[list[object]] = []
        self._totals = [Decimal(0) for _ in config.amount_columns]
        self._groups: OrderedDict[str, _GroupState] = OrderedDict()
        self._temp_dir: tempfile.TemporaryDirectory[str] | None = None
        self._db: sqlite3.Connection | None = None
        self._pending_writes = 0

        if config.group_column is not None or config.duplicate_columns:
            self._temp_dir = tempfile.TemporaryDirectory(prefix="excel_accounting_")
            db_path = Path(self._temp_dir.name) / "accounting.sqlite3"
            self._db = sqlite3.connect(db_path)
            if config.duplicate_columns:
                self._db.execute(
                    "CREATE TABLE duplicate_keys (key TEXT PRIMARY KEY, first_row INTEGER NOT NULL)"
                )
            if config.group_column is not None:
                self._db.execute(
                    "CREATE TABLE groups (name TEXT PRIMARY KEY, row_count INTEGER NOT NULL, "
                    "amounts TEXT NOT NULL)"
                )
            self._db.commit()

    def add_row(self, values: Sequence[object], source_row: int) -> None:
        """Add one matched source row; report invalid amounts instead of treating them as zero."""

        if not isinstance(source_row, int) or isinstance(source_row, bool) or source_row < 1:
            raise ValueError("source_row 必须是正整数")

        self.row_count += 1
        parsed_amounts: list[Decimal | None] = []
        for col in self.config.amount_columns:
            raw = self._cell(values, col)
            if self._blank(raw):
                self._issue(source_row, "金额为空", f"第 {col} 列金额为空")
                parsed_amounts.append(None)
                continue
            try:
                amount = Decimal(str(raw))
                if not amount.is_finite():
                    raise InvalidOperation
            except (InvalidOperation, ValueError):
                shown = str(raw)
                if len(shown) > 80:
                    shown = shown[:77] + "..."
                self._issue(source_row, "金额无效", f"第 {col} 列的值“{shown}”不是有效数字")
                parsed_amounts.append(None)
                continue
            parsed_amounts.append(amount)

        for index, amount in enumerate(parsed_amounts):
            if amount is not None:
                self._totals[index] += amount

        for col in self.config.required_columns:
            if self._blank(self._cell(values, col)):
                self._issue(source_row, "必填缺失", f"第 {col} 列不能为空")

        if self.config.duplicate_columns:
            assert self._db is not None
            key_values = [self._as_text(self._cell(values, col)) for col in self.config.duplicate_columns]
            key = json.dumps(key_values, ensure_ascii=False, separators=(",", ":"))
            try:
                self._db.execute(
                    "INSERT INTO duplicate_keys (key, first_row) VALUES (?, ?)",
                    (key, source_row),
                )
            except sqlite3.IntegrityError:
                first = self._db.execute(
                    "SELECT first_row FROM duplicate_keys WHERE key = ?", (key,)
                ).fetchone()
                columns = "、".join(str(col) for col in self.config.duplicate_columns)
                self._issue(
                    source_row,
                    "重复键",
                    f"第 {columns} 列的组合值与原始第 {first[0]} 行重复",
                )
            self._record_write()

        if self.config.group_column is not None:
            name = self._as_text(self._cell(values, self.config.group_column))
            state = self._get_group(name)
            state.count += 1
            for index, amount in enumerate(parsed_amounts):
                if amount is not None:
                    state.amounts[index] += amount
            self._groups.move_to_end(name)
            if len(self._groups) > _GROUP_CACHE_LIMIT:
                old_name, old_state = self._groups.popitem(last=False)
                self._write_group(old_name, old_state)

    def iter_summary_rows(self, headers: list[str]) -> Iterator[list[object]]:
        """Stream labeled totals and optional group rows from the temporary database."""

        amount_labels = [
            f"{headers[col - 1] if col <= len(headers) and headers[col - 1] else f'第 {col} 列'}合计"
            for col in self.config.amount_columns
        ]
        yield ["范围", "分组", "命中行数", *amount_labels]
        yield ["总计", "", self.row_count, *self._totals]
        if self.config.group_column is not None:
            self._flush_groups()
            assert self._db is not None
            for name, count, amounts_json in self._db.execute(
                "SELECT name, row_count, amounts FROM groups ORDER BY name"
            ):
                amounts = [Decimal(text) for text in json.loads(amounts_json)]
                yield ["分组", name, count, *amounts]

    def summary_rows(self, headers: list[str]) -> list[list[object]]:
        """Return summary rows as a list for callers handling small group sets."""

        return list(self.iter_summary_rows(headers))

    def issue_rows(self) -> list[list[object]]:
        """Return up to 10,000 issue rows: source row, type, description."""

        return [row.copy() for row in self._issues]

    def close(self) -> None:
        """Release the temporary database after summaries and issues are exported."""

        if self._db is not None:
            self._db.close()
            self._db = None
        if self._temp_dir is not None:
            self._temp_dir.cleanup()
            self._temp_dir = None

    def __enter__(self) -> AccountingAccumulator:
        return self

    def __exit__(self, *_exc: object) -> None:
        self.close()

    def __del__(self) -> None:
        try:
            self.close()
        except Exception:
            pass

    @staticmethod
    def _cell(values: Sequence[object], column: int) -> object:
        return values[column - 1] if column <= len(values) else None

    @staticmethod
    def _blank(value: object) -> bool:
        return value is None or isinstance(value, str) and not value.strip()

    @staticmethod
    def _as_text(value: object) -> str:
        return "" if value is None else str(value)

    def _issue(self, source_row: int, kind: str, description: str) -> None:
        self.issue_count += 1
        if len(self._issues) < _ISSUE_LIMIT:
            self._issues.append([source_row, kind, description])

    def _get_group(self, name: str) -> _GroupState:
        state = self._groups.get(name)
        if state is not None:
            return state
        assert self._db is not None
        record = self._db.execute(
            "SELECT row_count, amounts FROM groups WHERE name = ?", (name,)
        ).fetchone()
        if record is None:
            state = _GroupState(0, [Decimal(0) for _ in self.config.amount_columns])
        else:
            state = _GroupState(record[0], [Decimal(text) for text in json.loads(record[1])])
        self._groups[name] = state
        return state

    def _write_group(self, name: str, state: _GroupState) -> None:
        assert self._db is not None
        amounts_json = json.dumps([str(amount) for amount in state.amounts])
        self._db.execute(
            "INSERT INTO groups (name, row_count, amounts) VALUES (?, ?, ?) "
            "ON CONFLICT(name) DO UPDATE SET row_count = excluded.row_count, "
            "amounts = excluded.amounts",
            (name, state.count, amounts_json),
        )
        self._record_write()

    def _flush_groups(self) -> None:
        for name, state in self._groups.items():
            self._write_group(name, state)
        self._groups.clear()
        if self._db is not None:
            self._db.commit()
            self._pending_writes = 0

    def _record_write(self) -> None:
        self._pending_writes += 1
        if self._pending_writes >= _COMMIT_INTERVAL:
            assert self._db is not None
            self._db.commit()
            self._pending_writes = 0
