from __future__ import annotations

import argparse
import copy
import json
import os
import re
import shutil
import sys
import time
import urllib.error
import urllib.request
from datetime import date, datetime, timedelta, timezone
from pathlib import Path
from typing import Any
from zoneinfo import ZoneInfo

from openpyxl import load_workbook
from openpyxl.styles import PatternFill

API_BASE_URL = "http://127.0.0.1:8000/v1"
API_KEY = "your-api-key"
MODEL_NAME = "your-model-name"

SEARCH_START_DATE = None
SEARCH_END_DATE = None
SPECIFIC_REPORT_FOLDER = None

ASSET_DIR = Path(r"C:\Scripts\海外逆向到货")
ATTACHMENT_ROOT = ASSET_DIR / "Attachments" / "海外逆向到货"
OUTPUT_FOLDER = "報告生成"
DRAFT_REPORT_NAME = "Draft_Report_REVIEW_REQUIRED.xlsx"
FINAL_REPORT_NAME = "Final_Report.xlsx"
API_CHAT_PATH = "/chat/completions"
API_TIMEOUT_SECONDS = 240
API_MAX_RETRIES = 3
API_RETRY_SECONDS = 3
MODEL_TEMPERATURE = 0.0
MODEL_MAX_TOKENS = 1800
MIN_ACCEPTED_CONFIDENCE = 0.70
MAX_EVIDENCE_ITEMS = 12
MAX_EVIDENCE_CHARACTERS = 14000
KEEP_AI_RESPONSE_JSON = True
OVERWRITE_FINAL_REPORT = True
REPROCESS_COMPLETED_REPORTS = False
HK_TIMEZONE = ZoneInfo("Asia/Hong_Kong")
MANUAL_VALUE = "需自行填寫"
LIGHT_YELLOW = "FFF2CC"

HEADER_TO_FIELD = {
    "req number*": "req_number",
    "country*": "country",
    "im or ex*": "im_or_ex",
    "type*": "type_",
    "sub type": "sub_type",
    "trade country*": "trade_country",
    "transport mode*": "transport_mode",
    "trade term*": "trade_term",
    "declaration company*": "declaration_company",
    "seller name*": "seller_name",
    "seller address*": "seller_address",
    "buyer name*": "buyer_name",
    "buyer address*": "buyer_address",
    "shipper name*": "shipper_name",
    "shipper address*": "shipper_address",
    "consignee name*": "consignee_name",
    "consignee address*": "consignee_address",
    "trans via hk": "trans_via_hk",
    "total packages*": "total_packages",
    "departure country*": "departure_country",
    "departure port": "departure_port",
    "final destination*": "final_destination",
    "truckno/vesselname": "vessel_name",
    "flightno/ voyageno": "voyage_no",
    "flightno/voyageno": "voyage_no",
    "mawb": "mbl",
    "hawb": "hbl",
    "import/export date": "import_export_date",
    "item code*": "item_code",
    "item quantity*": "quantity",
    "unit*": "unit",
    "unit net weight*": "unit_net_weight",
    "unit price*": "unit_price",
    "currency*": "currency",
    "c.o.o*": "coo",
    "box material": "box_material",
    "business type": "bu",
    "bu": "bu",
}

LOCKED_FIELDS = {
    "country", "im_or_ex", "type_", "sub_type", "trade_country", "trans_via_hk",
    "final_destination", "coo", "bu", "consignee_name", "consignee_address",
    "import_export_date", "unit_net_weight", "item_code", "quantity", "unit", "currency",
    "box_material", "total_packages",
}

ALLOWED_DECLARATION_COMPANIES = {
    "0021-Huawei Technologies Co., Ltd.",
    "1421-Huawei International Co. Limited",
    "1241-Huawei International Pte. Ltd.",
    "0311-Huawei Tech. Investment Co., Limited",
    "1321-Huawei Device (Hong Kong) Co., Limited",
    "5531-Sparkoo Technologies Hong Kong Co., Limited",
}

ALLOWED_DOCUMENT_TYPES = {
    "invoice", "packing_list", "air_waybill_master", "air_waybill_house",
    "bill_of_lading_master", "bill_of_lading_house", "combined_invoice_packing_list",
    "arrival_notice", "other_shipping_document", "unknown",
}

AIR_MASTER_PATTERN = re.compile(r"^\d{3}-\d{8}$")
COUNTRY_PATTERN = re.compile(r"^[A-Z]{2}-.+$")
TRANSPORT_REFERENCE_NOISE = (
    "INVOICE", "PACKING", "CONTRACT", "BOOKING", "CONTAINER", "PURCHASEORDER", "PONUMBER"
)


def progress(message: str, level: str = "資訊") -> None:
    print(f"[{datetime.now(HK_TIMEZONE):%H:%M:%S}] [{level}] {message}", flush=True)


def utc_now() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


def norm(value: Any) -> str:
    return " ".join(str(value or "").replace("\r", " ").replace("\n", " ").strip().split())


def norm_header(value: Any) -> str:
    return norm(value).casefold().replace("（", "(").replace("）", ")")


def load_json(path: Path) -> dict[str, Any]:
    return json.loads(path.read_text(encoding="utf-8"))


def save_json(path: Path, data: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_suffix(path.suffix + ".tmp")
    temporary.write_text(json.dumps(data, ensure_ascii=False, indent=2, default=str), encoding="utf-8")
    temporary.replace(path)


def parse_date(value: Any, name: str) -> date | None:
    if value is None:
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, str) and value.strip():
        try:
            return datetime.strptime(value.strip(), "%Y-%m-%d").date()
        except ValueError as exc:
            raise ValueError(f"{name} 必須使用 YYYY-MM-DD 格式") from exc
    return None


def selected_period(start_value: Any = None, end_value: Any = None) -> tuple[date, date]:
    today = datetime.now(HK_TIMEZONE).date()
    start = parse_date(SEARCH_START_DATE if start_value is None else start_value, "SEARCH_START_DATE")
    end = parse_date(SEARCH_END_DATE if end_value is None else end_value, "SEARCH_END_DATE")
    if start is None and end is None:
        return today, today
    if start is None:
        start = end
    if end is None:
        end = start
    if start is None or end is None or start > end:
        raise ValueError("日期範圍無效")
    return start, end


def date_sequence(start: date, end: date):
    current = start
    while current <= end:
        yield current
        current += timedelta(days=1)


def discover_report_folders(start: date, end: date) -> tuple[list[Path], list[Path]]:
    selected = []
    missing = []
    for day in date_sequence(start, end):
        date_folder = ATTACHMENT_ROOT / day.isoformat()
        progress(f"掃描日期資料夾：{date_folder}")
        if not date_folder.is_dir():
            missing.append(date_folder)
            progress(f"日期資料夾不存在：{date_folder}", "警告")
            continue
        for shipment in sorted((path for path in date_folder.iterdir() if path.is_dir() and not path.name.startswith(".")), key=lambda path: path.name.casefold()):
            report_dir = shipment / OUTPUT_FOLDER
            request = report_dir / "ai_review_request.json"
            draft = report_dir / DRAFT_REPORT_NAME
            final = report_dir / FINAL_REPORT_NAME
            if not request.is_file() or not draft.is_file():
                continue
            if final.is_file() and not REPROCESS_COMPLETED_REPORTS:
                progress(f"略過已完成報表：{shipment.name}")
                continue
            selected.append(report_dir)
    return selected, missing


def validate_configuration() -> None:
    if not API_BASE_URL.strip() or API_BASE_URL.startswith("<"):
        raise ValueError("API_BASE_URL 尚未設定")
    if not MODEL_NAME.strip() or MODEL_NAME == "your-model-name":
        raise ValueError("MODEL_NAME 尚未設定")
    if not API_KEY.strip() or API_KEY == "your-api-key":
        progress("API_KEY 仍為預設值；如本機端點不需要驗證可以繼續，否則請先設定", "警告")


def endpoint_url() -> str:
    return API_BASE_URL.rstrip("/") + "/" + API_CHAT_PATH.strip("/")


def extract_json_object(text: str) -> dict[str, Any]:
    value = text.strip()
    if value.startswith("```"):
        value = re.sub(r"^```(?:json)?\s*", "", value, flags=re.IGNORECASE)
        value = re.sub(r"\s*```$", "", value)
    try:
        parsed = json.loads(value)
        if isinstance(parsed, dict):
            return parsed
    except Exception:
        pass
    start = value.find("{")
    if start < 0:
        raise ValueError("MODEL_RESPONSE_HAS_NO_JSON")
    depth = 0
    in_string = False
    escaped = False
    for index in range(start, len(value)):
        character = value[index]
        if in_string:
            if escaped:
                escaped = False
            elif character == "\\":
                escaped = True
            elif character == '"':
                in_string = False
        else:
            if character == '"':
                in_string = True
            elif character == "{":
                depth += 1
            elif character == "}":
                depth -= 1
                if depth == 0:
                    parsed = json.loads(value[start:index + 1])
                    if not isinstance(parsed, dict):
                        raise ValueError("MODEL_JSON_NOT_OBJECT")
                    return parsed
    raise ValueError("MODEL_JSON_INCOMPLETE")


def call_api(messages: list[dict[str, str]]) -> tuple[dict[str, Any], dict[str, Any]]:
    payload = {
        "model": MODEL_NAME,
        "messages": messages,
        "temperature": MODEL_TEMPERATURE,
        "max_tokens": MODEL_MAX_TOKENS,
        "stream": False,
    }
    headers = {"Content-Type": "application/json", "Accept": "application/json"}
    if API_KEY.strip():
        headers["Authorization"] = f"Bearer {API_KEY.strip()}"
    last_error = None
    for attempt in range(1, API_MAX_RETRIES + 1):
        try:
            request = urllib.request.Request(endpoint_url(), data=json.dumps(payload).encode("utf-8"), headers=headers, method="POST")
            with urllib.request.urlopen(request, timeout=API_TIMEOUT_SECONDS) as response:
                response_data = json.loads(response.read().decode("utf-8"))
            choices = response_data.get("choices") or []
            if not choices:
                raise ValueError("API_RESPONSE_HAS_NO_CHOICES")
            message = choices[0].get("message") or {}
            content = message.get("content")
            if isinstance(content, list):
                content = "".join(str(item.get("text", "")) if isinstance(item, dict) else str(item) for item in content)
            if not isinstance(content, str):
                raise ValueError("API_RESPONSE_HAS_NO_CONTENT")
            return extract_json_object(content), response_data.get("usage") or {}
        except urllib.error.HTTPError as exc:
            body = exc.read().decode("utf-8", errors="replace")[:2000]
            last_error = RuntimeError(f"HTTP {exc.code}: {body}")
        except Exception as exc:
            last_error = exc
        progress(f"API 呼叫失敗，嘗試 {attempt}/{API_MAX_RETRIES}：{type(last_error).__name__}: {last_error}", "警告")
        if attempt < API_MAX_RETRIES:
            time.sleep(API_RETRY_SECONDS * attempt)
    raise RuntimeError(f"API_CALL_FAILED: {last_error}")


def referenced_document_ids(task: dict[str, Any]) -> set[str]:
    ids = set()
    if task.get("document_id"):
        ids.add(str(task["document_id"]))
    for collection in (task.get("evidence", []), task.get("candidates", [])):
        for item in collection:
            if isinstance(item, dict) and item.get("document_id"):
                ids.add(str(item["document_id"]))
    return ids


def compact_task_evidence(task: dict[str, Any], report_dir: Path, manifest: dict[str, Any]) -> dict[str, Any]:
    evidence = []
    used = 0
    for item in task.get("evidence", [])[:MAX_EVIDENCE_ITEMS]:
        text = json.dumps(item, ensure_ascii=False, default=str)
        if used + len(text) > MAX_EVIDENCE_CHARACTERS:
            break
        evidence.append(item)
        used += len(text)
    document_ids = referenced_document_ids(task)
    normalized = []
    manifest_map = {str(item.get("document_id")): item for item in manifest.get("documents", [])}
    for document_id in sorted(document_ids):
        if used >= MAX_EVIDENCE_CHARACTERS:
            break
        manifest_item = manifest_map.get(document_id)
        normalized_path = report_dir / "normalized" / document_id / "document.json"
        if not normalized_path.is_file() and manifest_item and manifest_item.get("normalized_json"):
            candidate = Path(manifest_item["normalized_json"])
            if candidate.is_file() and report_dir in candidate.resolve().parents:
                normalized_path = candidate
        if not normalized_path.is_file():
            continue
        document = load_json(normalized_path)
        content = document.get("content", {})
        excerpt = {"document_id": document_id, "filename": document.get("source", {}).get("filename"), "classification": document.get("classification"), "content_kind": content.get("kind")}
        if content.get("kind") == "pdf":
            excerpt["pages"] = []
            for page in content.get("pages", []):
                page_item = {"page": page.get("page"), "text": str(page.get("text", ""))[:3000], "blocks": page.get("blocks", [])[:20], "tables": page.get("tables", [])[:5], "ocr": content.get("ocr")}
                serialized = json.dumps(page_item, ensure_ascii=False, default=str)
                if used + len(serialized) > MAX_EVIDENCE_CHARACTERS:
                    break
                excerpt["pages"].append(page_item)
                used += len(serialized)
        elif content.get("kind") == "workbook":
            excerpt["sheets"] = []
            for sheet in content.get("sheets", []):
                sheet_item = {"name": sheet.get("name"), "visibility": sheet.get("visibility"), "classification": sheet.get("classification"), "merged_ranges": sheet.get("merged_ranges", []), "rows": sheet.get("rows", [])[:150]}
                serialized = json.dumps(sheet_item, ensure_ascii=False, default=str)
                if used + len(serialized) > MAX_EVIDENCE_CHARACTERS:
                    sheet_item["rows"] = sheet_item["rows"][:30]
                    serialized = json.dumps(sheet_item, ensure_ascii=False, default=str)
                if used + len(serialized) > MAX_EVIDENCE_CHARACTERS:
                    break
                excerpt["sheets"].append(sheet_item)
                used += len(serialized)
        normalized.append(excerpt)
    return {"task_evidence": evidence, "normalized_json_excerpts": normalized}


def system_prompt() -> str:
    return (
        "You are a conservative shipment-data reviewer. Return exactly one JSON object and no Markdown. "
        "Use only the supplied JSON evidence. Never assume access to PDF, XLSX, MSG, images, workbook cells, internet, or external knowledge. "
        "Resolve only the requested task. If evidence is missing, ambiguous, conflicting, or confidence is below 0.70, return unresolved. "
        "Never change a field not requested. Never use Invoice, Packing List, Contract, Booking, Container, or PO number as MAWB or HAWB. "
        "For air mbl, require explicit MAWB or Master Air Waybill evidence and format NNN-NNNNNNNN. "
        "For sea mbl, require explicit MBL, Master B/L, Bill of Lading Number, or Sea Waybill evidence. "
        "For hbl, require explicit HAWB, HBL, House Air Waybill, or House Bill of Lading evidence. "
        "Transport mode must be BY AIR or BY SEA. Trade term must be CIP or DAP. Departure country must be XX-Country Name. "
        "For item tasks, identify the row by line_id, document, worksheet, and source row, not by item code alone. "
        "Output keys: task_id, entity_type, field, line_id, status, value, confidence, reason, provenance. "
        "status must be resolved or unresolved. provenance must be a JSON array."
    )


def task_prompt(task: dict[str, Any], evidence: dict[str, Any], shipment_context: dict[str, Any]) -> str:
    payload = {
        "instruction": "Resolve this one authorized task from JSON evidence only.",
        "task": task,
        "shipment_context": shipment_context,
        "evidence": evidence,
        "required_output_example": {
            "task_id": task.get("task_id"),
            "entity_type": task.get("entity_type"),
            "field": task.get("field"),
            "line_id": task.get("line_id"),
            "status": "resolved_or_unresolved",
            "value": None,
            "confidence": 0.0,
            "reason": "short evidence-based reason",
            "provenance": [],
        },
    }
    return json.dumps(payload, ensure_ascii=False, default=str)


def task_key(item: dict[str, Any]) -> tuple[Any, ...]:
    return item.get("task_id"), item.get("entity_type"), item.get("field"), item.get("line_id")


def validate_model_result(task: dict[str, Any], result: dict[str, Any], report: dict[str, Any]) -> dict[str, Any]:
    for key in ("task_id", "entity_type", "field"):
        if result.get(key) != task.get(key):
            raise ValueError(f"TASK_MISMATCH_{key.upper()}")
    if result.get("line_id") != task.get("line_id"):
        raise ValueError("TASK_MISMATCH_LINE_ID")
    if task.get("resolution_owner") != "ai_agent" or task.get("ai_editable") is not True or task.get("raw_file_access_allowed") is not False:
        raise ValueError("TASK_NOT_AUTHORIZED")
    if task.get("field") in LOCKED_FIELDS or task.get("current_value") == MANUAL_VALUE:
        raise ValueError("FIELD_LOCKED")
    status = result.get("status")
    if status not in {"resolved", "unresolved"}:
        raise ValueError("INVALID_STATUS")
    try:
        confidence = float(result.get("confidence", 0))
    except Exception as exc:
        raise ValueError("INVALID_CONFIDENCE") from exc
    if confidence < 0 or confidence > 1:
        raise ValueError("INVALID_CONFIDENCE")
    if status == "resolved" and confidence < MIN_ACCEPTED_CONFIDENCE:
        status = "unresolved"
        result["value"] = None
        result["reason"] = "Confidence below acceptance threshold"
    result["status"] = status
    result["confidence"] = confidence
    result["provenance"] = result.get("provenance") if isinstance(result.get("provenance"), list) else []
    result["reason"] = norm(result.get("reason")) or "No reason supplied"
    if status == "unresolved":
        result["value"] = None
        return result
    result["value"] = normalize_field_value(task["field"], result.get("value"), report)
    return result


def normalize_field_value(field: str, value: Any, report: dict[str, Any]) -> Any:
    text = norm(value)
    if not text:
        raise ValueError("EMPTY_VALUE")
    if field == "transport_mode":
        text = text.upper()
        if text not in {"BY AIR", "BY SEA"}:
            raise ValueError("INVALID_TRANSPORT_MODE")
    elif field == "trade_term":
        text = text.upper()
        if text not in {"CIP", "DAP"}:
            raise ValueError("INVALID_TRADE_TERM")
    elif field == "declaration_company":
        if text not in ALLOWED_DECLARATION_COMPANIES:
            raise ValueError("INVALID_DECLARATION_COMPANY")
    elif field == "departure_country":
        if not COUNTRY_PATTERN.fullmatch(text):
            raise ValueError("INVALID_DEPARTURE_COUNTRY")
    elif field == "unit_price":
        try:
            numeric = float(str(value).replace(",", ""))
        except Exception as exc:
            raise ValueError("INVALID_UNIT_PRICE") from exc
        if numeric < 0:
            raise ValueError("INVALID_UNIT_PRICE")
        return numeric
    elif field == "document_type":
        if text not in ALLOWED_DOCUMENT_TYPES:
            raise ValueError("INVALID_DOCUMENT_TYPE")
    elif field in {"mbl", "hbl", "req_number"}:
        compact = text.upper().replace(" ", "")
        noise_test = compact.replace("/", "").replace("-", "").replace("_", "").replace(".", "")
        if any(noise in noise_test for noise in TRANSPORT_REFERENCE_NOISE):
            raise ValueError("INVALID_TRANSPORT_REFERENCE")
        mode = report["shipments"][0]["header"].get("transport_mode")
        if field == "mbl" and mode == "BY AIR":
            if re.fullmatch(r"\d{11}", compact):
                compact = compact[:3] + "-" + compact[3:]
            if not AIR_MASTER_PATTERN.fullmatch(compact):
                raise ValueError("INVALID_MAWB_FORMAT")
            text = compact
    return text


def apply_result(report: dict[str, Any], result: dict[str, Any]) -> None:
    shipment = report["shipments"][0]
    field = result["field"]
    value = result["value"]
    if result["entity_type"] == "shipment":
        shipment["header"][field] = value
        metadata = shipment.setdefault("field_metadata", {}).setdefault(field, {})
        metadata.update({"value": value, "status": "resolved_by_ai", "confidence": result["confidence"], "provenance": result["provenance"]})
    elif result["entity_type"] == "item":
        item = next((item for item in shipment["items"] if item.get("line_id") == result.get("line_id")), None)
        if item is None:
            raise ValueError("ITEM_LINE_NOT_FOUND")
        item[field] = value
        item.setdefault("ai_provenance", {})[field] = {"confidence": result["confidence"], "provenance": result["provenance"]}
    elif result["entity_type"] == "document":
        document = next((item for item in shipment.get("documents", []) if item.get("document_id") == result.get("document_id") or item.get("document_id") == result.get("task_id")), None)
        if field != "document_type":
            raise ValueError("UNSUPPORTED_DOCUMENT_FIELD")
        if document is not None:
            document["classification"]["document_type"] = value
            document["classification"]["needs_review"] = False


def recalculate_dependencies(report: dict[str, Any]) -> None:
    shipment = report["shipments"][0]
    header = shipment["header"]
    header["shipper_name"] = header.get("seller_name")
    header["shipper_address"] = header.get("seller_address")
    candidates = shipment.get("req_number_resolution", {}).get("candidates", [])
    values = {item.get("type"): item.get("value") for item in candidates if item.get("value")}
    if header.get("hbl"):
        values["hbl"] = header["hbl"]
    header["req_number"] = values.get("hbl") or values.get("invoice_no") or values.get("pl_no") or values.get("filename") or header.get("req_number")
    shipment.setdefault("req_number_resolution", {})["selected_value"] = header.get("req_number")
    if header.get("transport_mode") == "BY AIR":
        header["vessel_name"] = None
    elif header.get("transport_mode") == "BY SEA" and header.get("voyage_no") and not norm(header["voyage_no"]).upper().startswith("V."):
        header["voyage_no"] = "V." + norm(header["voyage_no"])


def workbook_columns(sheet) -> dict[str, int]:
    return {HEADER_TO_FIELD[norm_header(cell.value)]: cell.column for cell in sheet[1] if norm_header(cell.value) in HEADER_TO_FIELD}


def item_rows(report: dict[str, Any]) -> dict[str, int]:
    return {item["line_id"]: index + 2 for index, item in enumerate(report["shipments"][0]["items"])}


def clear_fill(cell) -> None:
    cell.fill = PatternFill(fill_type=None)


def write_final_workbook(report_dir: Path, report: dict[str, Any], resolved: list[dict[str, Any]], unresolved: list[dict[str, Any]]) -> Path:
    draft = report_dir / DRAFT_REPORT_NAME
    final = report_dir / FINAL_REPORT_NAME
    if not draft.is_file():
        raise FileNotFoundError(draft)
    if final.exists() and not OVERWRITE_FINAL_REPORT:
        raise FileExistsError(final)
    temporary = report_dir / (FINAL_REPORT_NAME + ".tmp.xlsx")
    shutil.copy2(draft, temporary)
    workbook = load_workbook(temporary)
    original_names = list(workbook.sheetnames)
    original_validations = {name: len(workbook[name].data_validations.dataValidation) for name in original_names}
    sheet = workbook["模板"] if "模板" in workbook.sheetnames else workbook.active
    columns = workbook_columns(sheet)
    rows = item_rows(report)
    yellow = PatternFill(start_color=LIGHT_YELLOW, end_color=LIGHT_YELLOW, fill_type="solid")
    header = report["shipments"][0]["header"]
    items_by_line = {item["line_id"]: item for item in report["shipments"][0]["items"]}
    for result in resolved:
        field = result["field"]
        if field not in columns:
            continue
        if result["entity_type"] == "shipment":
            for row in rows.values():
                sheet.cell(row, columns[field]).value = header.get(field)
                clear_fill(sheet.cell(row, columns[field]))
        elif result["entity_type"] == "item":
            row = rows.get(result.get("line_id"))
            item = items_by_line.get(result.get("line_id"))
            if row and item:
                sheet.cell(row, columns[field]).value = item.get(field)
                clear_fill(sheet.cell(row, columns[field]))
    for field in ("req_number", "shipper_name", "shipper_address", "vessel_name", "voyage_no"):
        if field in columns:
            for row in rows.values():
                sheet.cell(row, columns[field]).value = header.get(field)
    for result in unresolved:
        field = result["field"]
        if field not in columns:
            continue
        if result["entity_type"] == "shipment":
            for row in rows.values():
                sheet.cell(row, columns[field]).value = MANUAL_VALUE
                sheet.cell(row, columns[field]).fill = copy.copy(yellow)
        elif result["entity_type"] == "item":
            row = rows.get(result.get("line_id"))
            if row:
                sheet.cell(row, columns[field]).value = MANUAL_VALUE
                sheet.cell(row, columns[field]).fill = copy.copy(yellow)
    if "import_export_date" in columns:
        for row in rows.values():
            sheet.cell(row, columns["import_export_date"]).value = MANUAL_VALUE
            sheet.cell(row, columns["import_export_date"]).fill = copy.copy(yellow)
    workbook.save(temporary)
    workbook.close()
    check = load_workbook(temporary)
    preserved = list(check.sheetnames) == original_names and all(len(check[name].data_validations.dataValidation) == original_validations[name] for name in original_names)
    check.close()
    if not preserved:
        temporary.unlink(missing_ok=True)
        raise RuntimeError("WORKBOOK_STRUCTURE_NOT_PRESERVED")
    if final.exists():
        final.unlink()
    temporary.replace(final)
    return final


def process_report_folder(report_dir: Path) -> dict[str, Any]:
    started = time.time()
    report_dir = report_dir.resolve()
    progress(f"開始 Step 3：{report_dir.parent.name}")
    required = ["ai_review_request.json", "extracted_report.json", "validation.json", "document_manifest.json", DRAFT_REPORT_NAME]
    missing = [name for name in required if not (report_dir / name).is_file()]
    if missing:
        raise FileNotFoundError("缺少必要檔案：" + ", ".join(missing))
    request = load_json(report_dir / "ai_review_request.json")
    report = load_json(report_dir / "extracted_report.json")
    manifest = load_json(report_dir / "document_manifest.json")
    if request.get("source_policy") != "JSON_ONLY" or request.get("raw_file_access_allowed") is not False:
        raise ValueError("Step 2 request 不是 JSON_ONLY")
    tasks = [task for task in request.get("review_tasks", []) if task.get("resolution_owner") == "ai_agent" and task.get("ai_editable") is True and task.get("raw_file_access_allowed") is False and task.get("field") not in LOCKED_FIELDS and task.get("current_value") != MANUAL_VALUE]
    if not tasks:
        progress("沒有可由 AI 修改的任務")
        final = report_dir / FINAL_REPORT_NAME
        if not final.is_file():
            shutil.copy2(report_dir / DRAFT_REPORT_NAME, final)
        return {"status": "NO_AI_TASKS", "final_report": str(final), "resolved": 0, "unresolved": 0, "rejected": 0}
    progress(f"授權 AI 任務數量：{len(tasks)}")
    resolved = []
    unresolved = []
    rejected = []
    raw_results = []
    total_prompt_tokens = 0
    total_completion_tokens = 0
    shipment_context = {"header": report["shipments"][0].get("header", {}), "req_number_resolution": report["shipments"][0].get("req_number_resolution", {}), "item_count": len(report["shipments"][0].get("items", []))}
    for index, task in enumerate(tasks, start=1):
        progress(f"AI 任務 {index}/{len(tasks)}：{task.get('field')}；task_id={task.get('task_id')}")
        try:
            evidence = compact_task_evidence(task, report_dir, manifest)
            result, usage = call_api([{"role": "system", "content": system_prompt()}, {"role": "user", "content": task_prompt(task, evidence, shipment_context)}])
            total_prompt_tokens += int(usage.get("prompt_tokens", 0) or 0)
            total_completion_tokens += int(usage.get("completion_tokens", 0) or 0)
            validated = validate_model_result(task, result, report)
            raw_results.append(validated)
            if validated["status"] == "resolved":
                apply_result(report, validated)
                resolved.append(validated)
                progress(f"任務已解決：{task.get('field')} = {validated.get('value')}；信心度={validated.get('confidence'):.2f}")
            else:
                unresolved.append(validated)
                progress(f"任務無法可靠解決：{task.get('field')}；將轉為人手填寫", "警告")
        except Exception as exc:
            rejected_result = {"task_id": task.get("task_id"), "entity_type": task.get("entity_type"), "field": task.get("field"), "line_id": task.get("line_id"), "status": "rejected", "value": None, "confidence": 0.0, "reason": f"{type(exc).__name__}: {exc}", "provenance": []}
            rejected.append(rejected_result)
            raw_results.append(rejected_result)
            progress(f"任務被拒絕：{task.get('field')}；{type(exc).__name__}: {exc}", "錯誤")
    recalculate_dependencies(report)
    final = write_final_workbook(report_dir, report, resolved, unresolved + rejected)
    response = {"schema_version": "1.2", "job_id": request.get("job_id"), "generated_at": utc_now(), "source_policy": "JSON_ONLY", "task_results": raw_results, "summary": {"total": len(tasks), "resolved": len(resolved), "unresolved": len(unresolved), "rejected": len(rejected)}}
    if KEEP_AI_RESPONSE_JSON:
        save_json(report_dir / "ai_review_response.json", response)
    report["step3"] = {"status": "COMPLETED", "resolved": len(resolved), "unresolved": len(unresolved), "rejected": len(rejected), "final_report": str(final), "generated_at": utc_now()}
    save_json(report_dir / "extracted_report_after_ai.json", report)
    summary = {"status": "COMPLETED", "generated_at": utc_now(), "report_dir": str(report_dir), "final_report": str(final), "resolved": len(resolved), "unresolved": len(unresolved), "rejected": len(rejected), "duration_seconds": round(time.time() - started, 2), "usage": {"prompt_tokens": total_prompt_tokens, "completion_tokens": total_completion_tokens}}
    save_json(report_dir / "step3_processing_summary.json", summary)
    progress(f"Step 3 完成：已解決={len(resolved)}；未解決={len(unresolved)}；拒絕={len(rejected)}")
    progress(f"最終報表：{final}")
    return summary


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--report-folder")
    parser.add_argument("--start-date")
    parser.add_argument("--end-date")
    parser.add_argument("--check", action="store_true")
    args = parser.parse_args()
    try:
        if args.check:
            print(json.dumps({"api_base_url": API_BASE_URL, "model_name": MODEL_NAME, "api_key_configured": bool(API_KEY.strip() and API_KEY != "your-api-key"), "attachment_root_exists": ATTACHMENT_ROOT.is_dir(), "today_hong_kong": str(datetime.now(HK_TIMEZONE).date())}, ensure_ascii=False, indent=2))
            return 0
        validate_configuration()
        progress("Step 3 AI 覆核及報表更新程序啟動")
        progress("證據政策：只讀 JSON；不會將 PDF、XLSX 附件或 MSG 傳送至 API")
        specific = args.report_folder if args.report_folder is not None else SPECIFIC_REPORT_FOLDER
        if specific:
            report_folders = [Path(specific).expanduser().resolve()]
            missing_dates = []
        else:
            start, end = selected_period(args.start_date, args.end_date)
            progress(f"香港日期範圍：{start} 至 {end}")
            report_folders, missing_dates = discover_report_folders(start, end)
        progress(f"待處理報告資料夾：{len(report_folders)}")
        results = []
        failures = []
        for index, report_dir in enumerate(report_folders, start=1):
            progress(f"總進度 {index}/{len(report_folders)}：{report_dir.parent.name}")
            try:
                results.append(process_report_folder(report_dir))
            except Exception as exc:
                failures.append({"report_dir": str(report_dir), "error": f"{type(exc).__name__}: {exc}"})
                progress(f"報告處理失敗：{report_dir.parent.name}；{type(exc).__name__}: {exc}", "錯誤")
        batch_summary = {"generated_at": utc_now(), "selected": len(report_folders), "completed": len(results), "failed": len(failures), "missing_date_folders": [str(path) for path in missing_dates], "results": results, "failures": failures}
        save_json(ASSET_DIR / "step3_last_batch_summary.json", batch_summary)
        print("\n" + "=" * 80)
        print("Step 3 技術處理摘要")
        print(f"已選取報告：{len(report_folders)}")
        print(f"成功完成：{len(results)}")
        print(f"處理失敗：{len(failures)}")
        print(f"缺少日期資料夾：{len(missing_dates)}")
        print("=" * 80)
        return 2 if failures else 0
    except Exception as exc:
        progress(f"程式執行失敗：{type(exc).__name__}: {exc}", "致命錯誤")
        return 2
    finally:
        try:
            input("\n按 Enter 鍵關閉視窗...")
        except Exception:
            pass


if __name__ == "__main__":
    raise SystemExit(main())
