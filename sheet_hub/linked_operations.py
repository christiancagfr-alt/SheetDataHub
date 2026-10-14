from __future__ import annotations

import re
import unicodedata
from datetime import date, datetime

from .models import SourceConfig
from .source_reader import SourceReader, google_retry, spreadsheet_id


def _parse_linked_date(value: str):
    from .engine import parse_date
    return parse_date(value)


class LinkedOperations:
    def _linked_mappings(self, setting: str) -> list[dict[str, str]]:
        mappings: list[dict[str, str]] = []
        for item in self.store.get(setting, []) or []:
            if not isinstance(item, dict):
                continue
            source = str(item.get("source") or "").strip()
            target = str(item.get("target") or "").strip()
            if source and target:
                mappings.append({"source": source, "target": target})
        return mappings

    def _linked_source(self, setting: str) -> SourceConfig:
        source_id = str(self.store.get(setting, "") or "").strip()
        source = next(
            (item for item in self.store.load_sources() if item.enabled and item.id == source_id),
            None,
        )
        if source is None:
            raise ValueError("请先选择一个已启用的数据源")
        source = SourceConfig.from_dict(source.to_dict())
        if not source.credential_path:
            source.credential_path = str(self.store.get("credential_path", "") or "").strip()
        return source

    def _linked_credential(self, source: SourceConfig | None = None) -> str:
        credential = str((source.credential_path if source else "") or "").strip()
        if not credential:
            credential = str(self.store.get("credential_path", "") or "").strip()
        if not credential:
            credential = next(
                (item.credential_path for item in self.store.load_sources() if item.credential_path),
                "",
            )
        if not credential:
            raise ValueError("写入 Google 表格需要配置服务账号 JSON")
        return credential

    def _linked_repository_credential(self) -> str:
        credential = str(self.store.get("credential_path", "") or "").strip()
        if credential:
            return credential
        submit_source_id = str(self.store.get("linked_submit_source_id", "") or "").strip()
        submit_source = next(
            (item for item in self.store.load_sources() if item.id == submit_source_id),
            None,
        )
        if submit_source and submit_source.credential_path:
            return submit_source.credential_path
        return self._linked_credential()

    @staticmethod
    def _linked_key(values: dict[str, object], fields: list[str]) -> tuple[str, ...]:
        folded = {str(key).strip().casefold(): value for key, value in values.items()}
        return tuple(
            str(folded.get(str(field).strip().casefold(), "") or "").strip().casefold()
            for field in fields
        )

    @staticmethod
    def _google_sheet_values(
        http,
        spreadsheet_id_value: str,
        sheet_name: str,
        logger,
    ) -> list[list[str]]:
        escaped = sheet_name.replace("'", "''")
        response = google_retry(
            lambda: http.values_get(
                spreadsheet_id_value,
                # Normal matching searches date headers from DJ to the right,
                # so the read range must not stop at DJ itself.
                f"'{escaped}'!A:ZZZ",
                params={"majorDimension": "ROWS", "valueRenderOption": "FORMATTED_VALUE"},
            ),
            logger,
        )
        return [list(row) for row in (response.get("values") or [])]

    @staticmethod
    def _google_range_values(http, spreadsheet_id_value: str, range_name: str, logger) -> list[list[str]]:
        response = google_retry(
            lambda: http.values_get(
                spreadsheet_id_value,
                range_name,
                params={"majorDimension": "ROWS", "valueRenderOption": "FORMATTED_VALUE"},
            ),
            logger,
        )
        return [list(row) for row in (response.get("values") or [])]

    @staticmethod
    def _linked_id_key(value: object) -> str:
        # The original Apps Script uses String(value).trim() on display values.
        return unicodedata.normalize("NFKC", str(value or "")).strip()

    def _linked_google_notes(
        self,
        http,
        spreadsheet_id_value: str,
        sheet_names: list[str],
        column_letter: str = "E",
    ) -> dict[tuple[str, int], str]:
        """Read Google Sheets inserted notes, keyed by (sheet title, 1-based row)."""
        if not sheet_names:
            return {}
        operation = "提交数据"
        logger = lambda level, message: self._log(operation, level, message)
        ranges = [
            f"'{name.replace(chr(39), chr(39) * 2)}'!{column_letter}:{column_letter}"
            for name in dict.fromkeys(sheet_names)
        ]
        response = google_retry(
            lambda: http.fetch_sheet_metadata(
                spreadsheet_id_value,
                params={
                    "includeGridData": True,
                    "ranges": ranges,
                    "fields": "sheets(properties(title),data(startRow,rowData(values(note))))",
                },
            ),
            logger,
        )
        notes: dict[tuple[str, int], str] = {}
        for sheet in response.get("sheets", []):
            title = str(sheet.get("properties", {}).get("title") or "")
            if not title:
                continue
            for grid in sheet.get("data", []):
                start_row = int(grid.get("startRow", 0) or 0)
                for offset, row_data in enumerate(grid.get("rowData", [])):
                    cells = row_data.get("values", [])
                    note = str(cells[0].get("note") or "").strip() if cells else ""
                    if note:
                        notes[(title, start_row + offset + 1)] = note
        self._log(operation, "INFO", f"已读取源表 E 列插入备注 {len(notes)} 条")
        return notes

    @staticmethod
    def _format_row_numbers(rows: list[int]) -> str:
        ordered = sorted(set(int(row) for row in rows))
        if not ordered:
            return ""
        groups: list[str] = []
        start = previous = ordered[0]
        for row in ordered[1:]:
            if row == previous + 1:
                previous = row
                continue
            groups.append(str(start) if start == previous else f"{start}-{previous}")
            start = previous = row
        groups.append(str(start) if start == previous else f"{start}-{previous}")
        return ",".join(groups)

    def _write_linked_run_log(self, values: list[object]) -> None:
        operation = str(values[1] if len(values) > 1 else "联动任务")
        data_url = str(self.store.get("linked_data_url", "") or "").strip()
        sid = spreadsheet_id(data_url)
        if not sid:
            raise ValueError("联动表链接无效，无法写入运行日志")
        credential = self._linked_repository_credential()
        http = SourceReader._gspread_client(credential).http_client
        logger = lambda level, message: self._log(operation, level, message)
        log_sheet = "运行日志"
        headers = [
            "执行时间", "操作", "状态", "来源/目标", "日期范围", "模式",
            "读取/目标数", "成功数", "未匹配/跳过数", "写入行号", "详情", "执行人",
        ]
        metadata = google_retry(
            lambda: http.fetch_sheet_metadata(
                sid,
                params={"fields": "sheets.properties(sheetId,title,gridProperties(columnCount))"},
            ), logger,
        )
        properties = next((
            item.get("properties", {}) for item in metadata.get("sheets", [])
            if item.get("properties", {}).get("title") == log_sheet
        ), None)
        if properties is None:
            google_retry(
                lambda: http.batch_update(sid, {"requests": [{"addSheet": {"properties": {
                    "title": log_sheet,
                    "gridProperties": {"rowCount": 2000, "columnCount": len(headers)},
                }}}]}), logger,
            )
        else:
            sheet_id = properties.get("sheetId")
            requests = []
            column_count = int(properties.get("gridProperties", {}).get("columnCount", len(headers)))
            if column_count < len(headers):
                requests.append({"appendDimension": {
                    "sheetId": sheet_id,
                    "dimension": "COLUMNS",
                    "length": len(headers) - column_count,
                }})
            requests.append({"insertDimension": {
                "range": {"sheetId": sheet_id, "dimension": "ROWS", "startIndex": 1, "endIndex": 2},
                "inheritFromBefore": False,
            }})
            google_retry(
                lambda: http.batch_update(sid, {"requests": requests}), logger,
            )
        row = [str(value if value is not None else "") for value in values[:len(headers)]]
        row.extend([""] * (len(headers) - len(row)))
        google_retry(
            lambda: http.values_update(
                sid, f"'{log_sheet}'!A1",
                params={"valueInputOption": "RAW"},
                body={"majorDimension": "ROWS", "values": [headers, row]},
            ), logger,
        )

    def _safe_write_linked_run_log(self, operation: str, values: list[object]) -> bool:
        try:
            self._write_linked_run_log(values)
            self._log(operation, "INFO", "联动表运行日志写入成功")
            return True
        except Exception as exc:
            self._log(operation, "WARNING", "联动表运行日志写入失败", repr(exc))
            return False

    def submit_linked_data(self) -> dict[str, int]:
        operation = "提交数据"
        timestamp = datetime.now().astimezone().strftime("%Y-%m-%d %H:%M:%S %z")
        submitter = str(self.store.get("linked_submitter", "") or "").strip()
        try:
            result = self._submit_linked_data_impl()
        except Exception as exc:
            self._safe_write_linked_run_log(operation, [
                timestamp, operation, "失败",
                str(self.store.get("linked_submit_sheet", "全部可用子 Sheet") or "全部可用子 Sheet"),
                "", "提交", 0, 0, 0, "", str(exc), submitter,
            ])
            raise
        written_rows = list(result.pop("_written_rows", []))
        source_detail = str(result.pop("_source_detail", ""))
        result["log_written"] = int(self._safe_write_linked_run_log(operation, [
            timestamp, operation, "成功",
            source_detail or str(self.store.get("linked_submit_sheet", "全部可用子 Sheet") or "全部可用子 Sheet"),
            "", "提交", result["read"], result["written"], result["duplicates"],
            self._format_row_numbers(written_rows),
            "", submitter,
        ]))
        return result

    def _submit_linked_data_impl(self) -> dict[str, object]:
        operation = "提交数据"
        self._log(operation, "INFO", "提交任务开始")
        source = self._linked_source("linked_submit_source_id")
        sheet_name = str(self.store.get("linked_submit_sheet", "") or "").strip()
        if sheet_name:
            source.include_sheets = [sheet_name]
        id_field = str(self.store.get("linked_submit_id_field", "ID") or "ID").strip()
        name_field = str(self.store.get("linked_submit_name_field", "名字") or "名字").strip()
        date_field = str(self.store.get("linked_submit_date_field", "日期") or "日期").strip()
        submitter = str(self.store.get("linked_submitter", "") or "").strip()
        if not all((id_field, name_field, date_field)):
            raise ValueError("请完整选择 ID、名字和日期来源字段")
        if not submitter:
            raise ValueError("请填写提交人")
        data_url = str(self.store.get("linked_data_url", "") or "").strip()
        sid = spreadsheet_id(data_url)
        if not sid:
            raise ValueError("请填写有效的数据表 Google 表格链接")
        target_sheet = str(self.store.get("linked_data_sheet", "data") or "data").strip()
        target_headers = ["ID", "名字", "日期", "备注", "提交人", "来源表格"]
        dedup_fields = ["ID", "日期", "来源表格"]

        source_sid = spreadsheet_id(source.url)
        if not source_sid:
            raise ValueError("读取 E 列插入备注需要使用 Google 表格数据源")
        records = self._reader(operation, source=source).read(source)
        credential = self._linked_credential(source)
        source_http = SourceReader._gspread_client(credential).http_client
        source_notes = self._linked_google_notes(
            source_http,
            source_sid,
            [record.sheet_name for record in records],
            "E",
        )
        mapped_rows = [{
            "ID": self._field_value(record, id_field),
            "名字": self._field_value(record, name_field),
            "日期": self._field_value(record, date_field),
            "备注": source_notes.get((record.sheet_name, record.row_number), ""),
            "提交人": submitter,
            "来源表格": f"{source.name}/{record.sheet_name}" if source.name else record.sheet_name,
        } for record in records]

        logger = lambda level, message: self._log(operation, level, message)
        http = SourceReader._gspread_client(credential).http_client
        metadata = google_retry(
            lambda: http.fetch_sheet_metadata(
                sid,
                params={
                    "fields": (
                        "sheets.properties(sheetId,title,"
                        "gridProperties(rowCount,columnCount))"
                    )
                },
            ),
            logger,
        )
        target_properties = next(
            (
                item["properties"]
                for item in metadata.get("sheets", [])
                if item["properties"]["title"] == target_sheet
            ),
            None,
        )
        names = {item["properties"]["title"] for item in metadata.get("sheets", [])}
        target_sheet_id = target_properties.get("sheetId") if target_properties else None
        target_row_count = int(
            (target_properties or {}).get("gridProperties", {}).get("rowCount", 1000)
        )
        if target_sheet not in names:
            create_response = google_retry(
                lambda: http.batch_update(
                    sid,
                    {"requests": [{"addSheet": {"properties": {
                        "title": target_sheet,
                        "gridProperties": {
                            "rowCount": max(1000, len(mapped_rows) + 10),
                            "columnCount": max(20, len(target_headers) + 2),
                        },
                    }}}]},
                ),
                logger,
            )
            if isinstance(create_response, dict):
                replies = create_response.get("replies") or []
                if replies:
                    created_properties = replies[0].get("addSheet", {}).get("properties", {})
                    target_sheet_id = created_properties.get("sheetId", target_sheet_id)
                    target_row_count = int(
                        created_properties.get("gridProperties", {}).get(
                            "rowCount", target_row_count
                        )
                    )
            values: list[list[str]] = []
        else:
            values = self._google_sheet_values(http, sid, target_sheet, logger)

        existing_headers = [str(value).strip() for value in (values[0] if values else [])]
        while existing_headers and not existing_headers[-1]:
            existing_headers.pop()
        if existing_headers and [item.casefold() for item in existing_headers[:6]] != [
            item.casefold() for item in target_headers
        ]:
            raise ValueError("联动数据表表头不正确；A-F 必须依次为：ID、名字、日期、备注、提交人、来源表格")
        headers = list(target_headers)

        normalized_rows = []
        for row in mapped_rows:
            folded_row = {str(key).casefold(): value for key, value in row.items()}
            normalized_rows.append({
                header: str(folded_row.get(header.casefold(), "") or "")
                for header in headers
            })
        existing_keys: set[tuple[str, ...]] = set()
        if dedup_fields and values:
            old_lookup = {header.casefold(): index for index, header in enumerate(existing_headers)}
            for old_row in values[1:]:
                item = {
                    field: (
                        str(old_row[old_lookup[field.casefold()]]).strip()
                        if field.casefold() in old_lookup and old_lookup[field.casefold()] < len(old_row)
                        else ""
                    )
                    for field in dedup_fields
                }
                key = self._linked_key(item, dedup_fields)
                if any(key):
                    existing_keys.add(key)

        new_rows: list[dict[str, str]] = []
        duplicates = 0
        for row in normalized_rows:
            if not str(row.get("ID", "")).strip():
                duplicates += 1
                continue
            if dedup_fields:
                key = self._linked_key(row, dedup_fields)
                if not any(key):
                    duplicates += 1
                    continue
                if key in existing_keys:
                    duplicates += 1
                    continue
                existing_keys.add(key)
            new_rows.append(row)

        escaped = target_sheet.replace("'", "''")
        header_changed = not existing_headers
        if header_changed:
            google_retry(
                lambda: http.values_update(
                    sid,
                    f"'{escaped}'!A1",
                    params={"valueInputOption": "RAW"},
                    body={"majorDimension": "ROWS", "values": [headers]},
                ),
                logger,
            )
        start_row = max(2, len(values) + 1)
        rows = [[row.get(header, "") for header in headers] for row in new_rows]
        for start in range(0, len(rows), 5000):
            chunk = rows[start:start + 5000]
            google_retry(
                lambda chunk=chunk, start=start: http.values_update(
                    sid,
                    f"'{escaped}'!A{start_row + start}",
                    params={"valueInputOption": "RAW"},
                    body={"majorDimension": "ROWS", "values": chunk},
                ),
                logger,
            )
        self._compact_google_sheet(
            http,
            sid,
            target_sheet_id,
            max(target_row_count, start_row + len(rows)),
            len(headers),
            logger,
        )
        self._log(
            operation,
            "INFO",
            f"提交结束：读取 {len(records)} 行，写入 {len(new_rows)} 行，跳过重复或空键 {duplicates} 行",
        )
        return {
            "read": len(records), "written": len(new_rows), "duplicates": duplicates,
            "_written_rows": list(range(start_row, start_row + len(new_rows))),
            "_source_detail": f"{source.name}/" + ",".join(dict.fromkeys(record.sheet_name for record in records)),
        }

    def match_linked_data(self) -> dict[str, int]:
        operation = "匹配写入"
        timestamp = datetime.now().astimezone().strftime("%Y-%m-%d %H:%M:%S %z")
        all_dates = bool(self.store.get("linked_match_all_dates", False))
        selected = str(self.store.get("linked_match_date", "") or "")
        scope = "全部日期" if all_dates else selected
        target = str(self.store.get("linked_match_target_sheet", "") or "")
        mode = str(self.store.get("linked_match_mode", "normal") or "normal")
        operator = str(self.store.get("linked_match_operator", "") or "").strip()
        try:
            result = self._match_linked_data_impl()
        except Exception as exc:
            self._safe_write_linked_run_log(operation, [
                timestamp, operation, "失败", target, scope, mode,
                0, 0, 0, "", str(exc), operator,
            ])
            raise
        written_detail = str(result.pop("_written_detail", ""))
        processed_dates = int(result.pop("_processed_dates", 1))
        result["log_written"] = int(self._safe_write_linked_run_log(operation, [
            timestamp, operation, "成功", target, scope, mode,
            result["target_rows"], result["matched"], result["unmatched"],
            written_detail,
            f"处理日期列 {processed_dates} 个；写入单元格 {result['updated']} 个",
            operator,
        ]))
        return result

    def _match_linked_data_impl(self) -> dict[str, object]:
        operation = "匹配写入"
        self._log(operation, "INFO", "匹配任务开始")
        target_source = self._linked_source("linked_match_target_source_id")
        target_sheet = str(self.store.get("linked_match_target_sheet", "") or "").strip()
        if not target_sheet:
            raise ValueError("请选择要写回的目标子 Sheet")
        data_url = str(self.store.get("linked_data_url", "") or "").strip()
        data_sheet = str(self.store.get("linked_data_sheet", "data") or "data").strip()
        mode = str(self.store.get("linked_match_mode", "normal") or "normal")
        all_dates = bool(self.store.get("linked_match_all_dates", False))
        selected_date = _parse_linked_date(str(self.store.get("linked_match_date", "") or ""))
        operator = str(self.store.get("linked_match_operator", "") or "").strip()
        if not operator:
            raise ValueError("请填写匹配执行人")
        if selected_date is None:
            raise ValueError("请选择有效的匹配日期")
        if all_dates and mode == "makeup":
            raise ValueError("补课模式固定写入 DI，只支持选择一个日期；请取消“匹配所有日期”")
        sid = spreadsheet_id(data_url)
        target_sid = spreadsheet_id(target_source.url)
        if not sid:
            raise ValueError("请填写有效的数据表 Google 表格链接")
        if not target_sid:
            raise ValueError("匹配写回目前只支持 Google 表格数据源")

        repository_credential = self._linked_repository_credential()
        target_credential = self._linked_credential(target_source)
        logger = lambda level, message: self._log(operation, level, message)
        repository_http = SourceReader._gspread_client(repository_credential).http_client
        escaped_data_sheet = data_sheet.replace("'", "''")
        repository_values = self._google_range_values(
            repository_http, sid, f"'{escaped_data_sheet}'!A:D", logger
        )
        if not repository_values:
            raise ValueError("联动数据表没有可匹配的数据，已停止写回")

        # Faithful to the original script: A is the ID, C is the value/date,
        # and D is the inserted note to attach to the target cell.
        repository_entries: list[tuple[str, str, str]] = []
        for row in repository_values:
            source_id = str(row[0]).strip() if len(row) > 0 else ""
            content = str(row[2]).strip() if len(row) > 2 else ""
            note = str(row[3]).strip() if len(row) > 3 else ""
            if not source_id or not content:
                continue
            key = self._linked_id_key(source_id)
            if key:
                repository_entries.append((key, content, note))
        if not repository_entries:
            raise ValueError("联动数据表 A/C 列没有可匹配的数据，已停止写回")

        http = SourceReader._gspread_client(target_credential).http_client
        values = self._google_sheet_values(http, target_sid, target_sheet, logger)
        if not values:
            raise ValueError("目标子 Sheet 没有可用数据")
        if all_dates:
            tasks: list[tuple[date, int]] = []
            seen_dates: set[tuple[int, int]] = set()
            for column in range(113, len(values[0])):
                month_day = self._linked_extract_month_day(values[0][column], selected_date.year)
                if month_day is None or month_day in seen_dates:
                    continue
                month, day = month_day
                try:
                    task_date = date(selected_date.year, month, day)
                except ValueError:
                    continue
                tasks.append((task_date, column))
                seen_dates.add(month_day)
            if not tasks:
                raise ValueError("从 DJ 列开始没有找到可识别的日期表头")
        else:
            tasks = [(selected_date, self._linked_target_column(values[0], selected_date, mode))]
        unmatched_value = str(self.store.get("linked_match_unmatched_value", "") or "").strip() or "❌absent"
        matched_rows = 0
        unmatched_rows = 0
        update_requests: list[dict[str, object]] = []
        written_details: list[str] = []
        key_columns, lookup_mode = self._linked_sheet_key_config(target_sheet)
        for task_date, target_column in tasks:
            index: dict[str, list[tuple[str, str]]] = {}
            for key, content, note in repository_entries:
                if self._linked_repository_date_matches(content, task_date):
                    index.setdefault(key, []).append((content, note))
            if not all_dates and not index:
                raise ValueError("联动数据表在所选日期没有可匹配的数据，已停止写回")
            output_rows: list[dict[str, object]] = []
            matched_for_date: list[int] = []
            for matrix_row in range(2, len(values)):
                row = values[matrix_row]
                keys: list[str] = []
                for column in key_columns:
                    key = self._linked_id_key(row[column] if len(row) > column else "")
                    if key and key not in keys:
                        keys.append(key)
                matched: list[tuple[str, str]] = []
                for key in keys:
                    found = index.get(key, [])
                    if lookup_mode == "fallback":
                        if found:
                            matched = found
                            break
                    else:
                        matched.extend(found)
                if matched:
                    matched_rows += 1
                    matched_for_date.append(matrix_row + 1)
                    contents: list[str] = []
                    notes: list[str] = []
                    for content, note in matched:
                        if content in contents:
                            continue
                        contents.append(content)
                        if note:
                            notes.append(note)
                    output = "\n".join(contents)
                    note_text = "\n".join(notes)
                elif keys:
                    unmatched_rows += 1
                    output = unmatched_value
                    note_text = ""
                else:
                    output = ""
                    note_text = ""
                cell: dict[str, object] = {}
                if output:
                    cell["userEnteredValue"] = {"stringValue": output}
                if note_text:
                    cell["note"] = note_text
                output_rows.append({"values": [cell]})
            label = str(values[0][target_column] if target_column < len(values[0]) else task_date.isoformat())
            written_details.append(f"{label}：{self._format_row_numbers(matched_for_date) or '无'}")
            update_requests.append({"updateCells": {
                "range": {
                    "startRowIndex": 2,
                    "endRowIndex": 2 + len(output_rows),
                    "startColumnIndex": target_column,
                    "endColumnIndex": target_column + 1,
                },
                "rows": output_rows,
                "fields": "userEnteredValue,note",
            }})

        metadata = google_retry(
            lambda: http.fetch_sheet_metadata(
                target_sid, params={"fields": "sheets.properties(sheetId,title)"}
            ), logger,
        )
        sheet_id = next((
            item["properties"]["sheetId"] for item in metadata.get("sheets", [])
            if item.get("properties", {}).get("title") == target_sheet
        ), None)
        if sheet_id is None:
            raise ValueError("目标子 Sheet 不存在")
        for request in update_requests:
            request["updateCells"]["range"]["sheetId"] = sheet_id
        if update_requests:
            google_retry(
                lambda: http.batch_update(target_sid, {"requests": update_requests}), logger,
            )
        target_rows = max(0, len(values) - 2)
        updated_cells = target_rows * len(tasks)
        self._log(
            operation,
            "INFO",
            f"匹配结束：处理 {len(tasks)} 个日期列，目标 {target_rows} 行，匹配 {matched_rows} 行，未匹配 {unmatched_rows} 行，重写 {updated_cells} 个单元格",
        )
        return {
            "target_rows": target_rows,
            "matched": matched_rows,
            "unmatched": unmatched_rows,
            "updated": updated_cells,
            "skipped_filled": 0,
            "_written_detail": "；".join(written_details)[:45000],
            "_processed_dates": len(tasks),
        }

    @staticmethod
    def _linked_sheet_key_config(sheet_name: str) -> tuple[tuple[int, int], str]:
        if "🍓" in sheet_name:
            return (27, 29), "merge"  # AB / AD, merge both IDs.
        if "🏠" in sheet_name:
            return (14, 16), "fallback"  # O first; only then Q.
        raise ValueError("目标子 Sheet 名称必须包含 🍓 或 🏠，与原匹配代码规则一致")

    @staticmethod
    def _linked_target_column(headers: list[object], selected: date, mode: str) -> int:
        if mode == "makeup":
            return 112  # DI
        return LinkedOperations._linked_normal_target_column(headers, selected)

    @staticmethod
    def _linked_repository_date_matches(value: str, selected: date) -> bool:
        return LinkedOperations._linked_date_text_matches(value, selected)

    @staticmethod
    def _linked_extract_month_day(value: object, reference_year: int) -> tuple[int, int] | None:
        text = unicodedata.normalize("NFKC", str(value or "")).replace("\u00a0", " ").strip()
        if not text:
            return None
        parsed = _parse_linked_date(text)
        if parsed is not None:
            return parsed.month, parsed.day
        full_date = re.search(r"(?<!\d)(\d{4})[./-](\d{1,2})[./-](\d{1,2})(?!\d)", text)
        if full_date:
            try:
                parsed = date(int(full_date.group(1)), int(full_date.group(2)), int(full_date.group(3)))
                return parsed.month, parsed.day
            except ValueError:
                return None
        chinese = re.search(r"(?<!\d)(\d{1,2})\s*月\s*(\d{1,2})\s*日?", text)
        if chinese:
            month, day = int(chinese.group(1)), int(chinese.group(2))
            try:
                date(reference_year, month, day)
                return month, day
            except ValueError:
                return None
        month_names = (
            ("jan", "january", "janeiro"), ("feb", "february", "fev", "fevereiro"),
            ("mar", "march", "março"), ("apr", "april", "abr", "abril"),
            ("may", "mai", "maio"), ("jun", "june", "junho"),
            ("jul", "july", "julho"), ("aug", "august", "ago", "agosto"),
            ("sep", "sept", "september", "set", "setembro"),
            ("oct", "october", "out", "outubro"), ("nov", "november", "novembro"),
            ("dec", "december", "dez", "dezembro"),
        )
        for month, names in enumerate(month_names, start=1):
            tokens = "|".join(re.escape(name) for name in sorted(names, key=len, reverse=True))
            matched = re.search(rf"(?<![A-Za-zÀ-ÿ])(?:{tokens})\.?\s*(\d{{1,2}})(?!\d)", text, re.IGNORECASE)
            if matched:
                day = int(matched.group(1))
                try:
                    date(reference_year, month, day)
                    return month, day
                except ValueError:
                    return None
        numeric = re.search(r"(?<!\d)(\d{1,2})\s*[./-]\s*(\d{1,2})(?!\d)", text)
        if numeric:
            first, second = int(numeric.group(1)), int(numeric.group(2))
            month, day = (second, first) if first > 12 else (first, second)
            try:
                date(reference_year, month, day)
                return month, day
            except ValueError:
                return None
        return None

    @staticmethod
    def _linked_date_text_matches(value: object, selected: date) -> bool:
        """Match the selected month/day across common Sheets display formats."""
        text = unicodedata.normalize("NFKC", str(value or "")).replace("\u00a0", " ").strip()
        if not text:
            return False
        parsed = _parse_linked_date(text)
        if parsed is not None and parsed.month == selected.month and parsed.day == selected.day:
            return True
        month = selected.month
        day = selected.day
        patterns = [
            rf"(?<!\d)0?{month}\s*[./-]\s*0?{day}(?!\d)",
            rf"(?<!\d)0?{day}\s*[./-]\s*0?{month}(?!\d)",
            rf"(?<!\d)0?{month}\s*月\s*0?{day}\s*日?(?!\d)",
        ]
        english = (
            ("Jan", "January"), ("Feb", "February"), ("Mar", "March"),
            ("Apr", "April"), ("May",), ("Jun", "June"),
            ("Jul", "July"), ("Aug", "August"), ("Sep", "Sept", "September"),
            ("Oct", "October"), ("Nov", "November"), ("Dec", "December"),
        )[month - 1]
        portuguese = (
            ("Jan", "Janeiro"), ("Fev", "Fevereiro"), ("Mar", "Março"),
            ("Abr", "Abril"), ("Mai", "Maio"), ("Jun", "Junho"),
            ("Jul", "Julho"), ("Ago", "Agosto"), ("Set", "Setembro"),
            ("Out", "Outubro"), ("Nov", "Novembro"), ("Dez", "Dezembro"),
        )[month - 1]
        month_tokens = sorted({*english, *portuguese}, key=len, reverse=True)
        token_pattern = "|".join(re.escape(token) for token in month_tokens)
        patterns.append(rf"(?<![A-Za-zÀ-ÿ])(?:{token_pattern})\.?\s*0?{day}(?!\d)")
        return any(re.search(pattern, text, re.IGNORECASE) for pattern in patterns)

    @staticmethod
    def _linked_normal_target_column(headers: list[object], selected: date) -> int:
        for column in range(113, len(headers)):
            if LinkedOperations._linked_date_text_matches(headers[column], selected):
                return column
        samples = [str(value).strip() for value in headers[113:123] if str(value).strip()]
        detail = f"；DJ 起已读取到：{', '.join(samples)}" if samples else "；DJ 起没有读取到日期表头"
        raise ValueError(f"从 DJ 列开始找不到所选日期对应的表头{detail}")

    @staticmethod
    def _compact_google_sheet(
        http,
        spreadsheet_key: str,
        sheet_id: int | None,
        row_count: int,
        column_count: int,
        logger,
    ) -> None:
        """Use normal cell dimensions while preserving complete cell values."""
        if sheet_id is None or column_count <= 0:
            return
        end_row = max(2, row_count)
        google_retry(
            lambda: http.batch_update(
                spreadsheet_key,
                {
                    "requests": [
                        {
                            "repeatCell": {
                                "range": {
                                    "sheetId": sheet_id,
                                    "startRowIndex": 1,
                                    "endRowIndex": end_row,
                                    "startColumnIndex": 0,
                                    "endColumnIndex": column_count,
                                },
                                "cell": {
                                    "userEnteredFormat": {
                                        "wrapStrategy": "CLIP",
                                    }
                                },
                                "fields": "userEnteredFormat.wrapStrategy",
                            }
                        },
                        {
                            "updateDimensionProperties": {
                                "range": {
                                    "sheetId": sheet_id,
                                    "dimension": "ROWS",
                                    "startIndex": 1,
                                    "endIndex": end_row,
                                },
                                "properties": {"pixelSize": 21},
                                "fields": "pixelSize",
                            }
                        },
                        {
                            "updateDimensionProperties": {
                                "range": {
                                    "sheetId": sheet_id,
                                    "dimension": "COLUMNS",
                                    "startIndex": 0,
                                    "endIndex": column_count,
                                },
                                "properties": {"pixelSize": 100},
                                "fields": "pixelSize",
                            }
                        },
                    ]
                },
            ),
            logger,
        )

