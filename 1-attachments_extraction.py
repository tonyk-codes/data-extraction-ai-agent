import json
import logging
import re
import shutil
import sys
import traceback
from datetime import date, datetime, time, timedelta
from pathlib import Path
from zoneinfo import ZoneInfo

import pythoncom
import win32com.client

OUTPUT_ROOT = Path(r"C:\Scripts\海外逆向到货\Attachments\海外逆向到货")

SUBJECT_KEYWORDS = (
    "海运到货",
    "空运到货",
)

ALLOWED_EXTENSIONS = {
    ".pdf",
    ".xlsx",
}

#2026-12-01
#r"Z:\单证与关务组\AI Agent\海外逆向到货"
SEARCH_START_DATE = None
SEARCH_END_DATE = None
MSG_FOLDER_PATH = r"C:\Scripts\海外逆向到货\Database"

REPROCESS_ALREADY_PROCESSED = False

HK_TIMEZONE = ZoneInfo("Asia/Hong_Kong")
OL_FOLDER_INBOX = 6
OL_MAIL_ITEM = 43
OL_SAVE_AS_MSG = 3
INVALID_WINDOWS_CHARS = re.compile(r'[<>:"/\\|?*\x00-\x1f]')
RESERVED_WINDOWS_NAMES = {
    "CON", "PRN", "AUX", "NUL",
    *(f"COM{i}" for i in range(1, 10)),
    *(f"LPT{i}" for i in range(1, 10)),
}
LABEL_SEPARATOR = r"[\s._:/\\#=\-]*"
NUMBER_PATTERN = r"(?P<number>[A-Z0-9][A-Z0-9._/\-]{2,80})"
REFERENCE_PATTERNS = (
    (
        "MBL",
        re.compile(
            rf"(?<![A-Z0-9])(?:M{LABEL_SEPARATOR}B{LABEL_SEPARATOR}L|MASTER{LABEL_SEPARATOR}B{LABEL_SEPARATOR}L|MASTER{LABEL_SEPARATOR}BILL{LABEL_SEPARATOR}OF{LABEL_SEPARATOR}LADING){LABEL_SEPARATOR}(?:(?:NO|NUMBER|NUM|REF|REFERENCE){LABEL_SEPARATOR})?{NUMBER_PATTERN}",
            re.IGNORECASE,
        ),
    ),
    (
        "BL",
        re.compile(
            rf"(?<![A-Z0-9])(?:B{LABEL_SEPARATOR}L|BILL{LABEL_SEPARATOR}OF{LABEL_SEPARATOR}LADING){LABEL_SEPARATOR}(?:(?:NO|NUMBER|NUM|REF|REFERENCE){LABEL_SEPARATOR})?{NUMBER_PATTERN}",
            re.IGNORECASE,
        ),
    ),
    (
        "HBL",
        re.compile(
            rf"(?<![A-Z0-9])(?:H{LABEL_SEPARATOR}B{LABEL_SEPARATOR}L|HOUSE{LABEL_SEPARATOR}B{LABEL_SEPARATOR}L|HOUSE{LABEL_SEPARATOR}BILL{LABEL_SEPARATOR}OF{LABEL_SEPARATOR}LADING){LABEL_SEPARATOR}(?:(?:NO|NUMBER|NUM|REF|REFERENCE){LABEL_SEPARATOR})?{NUMBER_PATTERN}",
            re.IGNORECASE,
        ),
    ),
)


def get_script_directory() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    if "__file__" in globals():
        return Path(__file__).resolve().parent
    return Path.cwd().resolve()


SCRIPT_DIRECTORY = get_script_directory()
LOCAL_LOG_FILE = SCRIPT_DIRECTORY / "outlook_attachment_downloader.log"
NETWORK_LOG_FILE = OUTPUT_ROOT / "outlook_attachment_downloader.log"
STATE_FILE = OUTPUT_ROOT / "processed_outlook_messages.json"


def display_progress(message: str) -> None:
    print(f"[{datetime.now():%H:%M:%S}] {message}", flush=True)


def add_file_handler(logger: logging.Logger, path: Path) -> bool:
    try:
        path.parent.mkdir(parents=True, exist_ok=True)
        handler = logging.FileHandler(path, encoding="utf-8", mode="a")
        handler.setFormatter(logging.Formatter("%(asctime)s | %(levelname)s | %(message)s"))
        logger.addHandler(handler)
        return True
    except Exception:
        return False


def configure_logging() -> logging.Logger:
    logger = logging.getLogger("outlook_attachment_downloader")
    logger.setLevel(logging.INFO)
    logger.propagate = False
    for handler in list(logger.handlers):
        logger.removeHandler(handler)
        try:
            handler.close()
        except Exception:
            pass
    add_file_handler(logger, LOCAL_LOG_FILE)
    try:
        same_location = LOCAL_LOG_FILE.resolve() == NETWORK_LOG_FILE.resolve()
    except Exception:
        same_location = str(LOCAL_LOG_FILE).casefold() == str(NETWORK_LOG_FILE).casefold()
    if not same_location:
        add_file_handler(logger, NETWORK_LOG_FILE)
    logger.info("=" * 80)
    logger.info("程式已啟動")
    return logger


LOGGER = configure_logging()


def parse_config_date(value, variable_name: str) -> date | None:
    if value is None:
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, str):
        value = value.strip()
        if not value:
            return None
        try:
            return datetime.strptime(value, "%Y-%m-%d").date()
        except ValueError as exc:
            raise ValueError(f"{variable_name} 必須使用 YYYY-MM-DD 格式：{value!r}") from exc
    raise TypeError(f"{variable_name} 格式無效。")


def get_search_period() -> tuple[datetime, datetime]:
    today_hk = datetime.now(HK_TIMEZONE).date()
    start_date = parse_config_date(SEARCH_START_DATE, "SEARCH_START_DATE")
    end_date = parse_config_date(SEARCH_END_DATE, "SEARCH_END_DATE")
    if start_date is None and end_date is None:
        start_date = today_hk
        end_date = today_hk
    elif start_date is None:
        start_date = end_date
    elif end_date is None:
        end_date = start_date
    if start_date is None or end_date is None:
        raise RuntimeError("無法決定搜尋日期範圍。")
    if start_date > end_date:
        raise ValueError("SEARCH_START_DATE 唔可以遲過 SEARCH_END_DATE。")
    return datetime.combine(start_date, time.min), datetime.combine(end_date + timedelta(days=1), time.min)


def outlook_datetime_to_hk_naive(value) -> datetime:
    if value is None:
        return datetime.now(HK_TIMEZONE).replace(tzinfo=None)
    if getattr(value, "tzinfo", None) is not None:
        try:
            return value.astimezone(HK_TIMEZONE).replace(tzinfo=None)
        except Exception:
            LOGGER.exception("轉換 Outlook 時間失敗。")
    return datetime(value.year, value.month, value.day, value.hour, value.minute, value.second, getattr(value, "microsecond", 0))


def sanitize_name(value: str, fallback: str, max_length: int | None = None) -> str:
    value = INVALID_WINDOWS_CHARS.sub("_", str(value or "").strip())
    value = re.sub(r"\s+", " ", value).rstrip(". ")
    if not value:
        value = fallback
    if value.split(".", 1)[0].upper() in RESERVED_WINDOWS_NAMES:
        value = f"_{value}"
    if max_length is not None:
        value = value[:max_length].rstrip(". ")
    return value or fallback


def clean_reference_number(value: str) -> str | None:
    if not value:
        return None
    value = value.strip(" \t\r\n:;,#/\\|()[]{}'\"")
    value = re.sub(r"\s+", "", value)
    value = value.replace("/", "-").replace("\\", "-").rstrip("._-")
    if len(value) < 3 or not re.search(r"\d", value):
        return None
    return sanitize_name(value, "REFERENCE_UNKNOWN", 100)


def extract_shipping_reference(subject: str) -> tuple[str | None, str | None]:
    for reference_type, pattern in REFERENCE_PATTERNS:
        for match in pattern.finditer(str(subject or "")):
            number = clean_reference_number(match.group("number"))
            if number:
                return number, reference_type
    return None, None


def get_destination_folder_name(subject: str) -> tuple[str, str]:
    number, reference_type = extract_shipping_reference(subject)
    if number and reference_type:
        return number, reference_type
    return sanitize_name(subject, "NO_SUBJECT", 180), "FALLBACK"


def unique_file_path(folder: Path, filename: str) -> Path:
    source = Path(str(filename or "file"))
    stem = sanitize_name(source.stem, "file", 150)
    suffix = source.suffix
    destination = folder / f"{stem}{suffix}"
    counter = 2
    while destination.exists():
        counter_text = f"_{counter}"
        destination = folder / f"{stem[:max(1, 150 - len(counter_text))]}{counter_text}{suffix}"
        counter += 1
    return destination


def load_processed_ids() -> set[str]:
    if not STATE_FILE.exists():
        return set()
    try:
        with STATE_FILE.open("r", encoding="utf-8") as file:
            data = json.load(file)
        return {str(value) for value in data.get("processed_entry_ids", []) if value}
    except Exception:
        LOGGER.exception("讀取已處理郵件記錄失敗。")
        return set()


def save_processed_ids(processed_ids: set[str]) -> None:
    OUTPUT_ROOT.mkdir(parents=True, exist_ok=True)
    temporary_file = STATE_FILE.with_suffix(".json.tmp")
    data = {
        "updated_at_hk": datetime.now(HK_TIMEZONE).isoformat(timespec="seconds"),
        "processed_entry_ids": sorted(processed_ids),
    }
    with temporary_file.open("w", encoding="utf-8") as file:
        json.dump(data, file, ensure_ascii=False, indent=2)
    temporary_file.replace(STATE_FILE)


def message_received_time(message) -> datetime:
    try:
        return outlook_datetime_to_hk_naive(getattr(message, "ReceivedTime", None))
    except Exception:
        LOGGER.exception("未能讀取郵件收件時間。")
        return datetime.now(HK_TIMEZONE).replace(tzinfo=None)


def is_reply_or_forward(message, subject: str) -> bool:
    if re.match(r"^(?:re|fw|fwd|回覆|回复|轉寄|转发)\s*[:：]", subject, re.IGNORECASE):
        return True
    return bool(str(getattr(message, "InReplyTo", "") or "").strip())


def scan_supported_attachments(message) -> list[tuple[int, str]]:
    attachments = message.Attachments
    supported = []
    for index in range(1, int(getattr(attachments, "Count", 0)) + 1):
        filename = str(getattr(attachments.Item(index), "FileName", "") or "")
        if Path(filename).suffix.lower() in ALLOWED_EXTENSIONS:
            supported.append((index, filename))
    return supported


def save_message_file(message, mbl_folder: Path, subject: str, local_msg_path: Path | None) -> Path:
    mbl_folder.mkdir(parents=True, exist_ok=True)
    if local_msg_path is not None:
        destination = unique_file_path(mbl_folder, local_msg_path.name)
        shutil.copy2(local_msg_path, destination)
        display_progress(f"已複製 MSG 檔案：{destination}")
        return destination
    filename = f"{sanitize_name(subject, 'NO_SUBJECT', 150)}.msg"
    destination = unique_file_path(mbl_folder, filename)
    message.SaveAs(str(destination), OL_SAVE_AS_MSG)
    display_progress(f"已儲存 Outlook 郵件為 MSG：{destination}")
    return destination


def download_attachments(message, supported: list[tuple[int, str]], mbl_folder: Path) -> tuple[int, int]:
    attachments_folder = mbl_folder / "Attachments"
    downloaded = 0
    failed = 0
    display_progress(f"郵件共有 {len(supported)} 個支援附件。")
    for index, filename in supported:
        try:
            attachment = message.Attachments.Item(index)
            display_progress(f"下載附件 {index}：{filename}")
            attachments_folder.mkdir(parents=True, exist_ok=True)
            destination = unique_file_path(attachments_folder, filename)
            attachment.SaveAsFile(str(destination))
            downloaded += 1
            display_progress(f"附件下載完成：{destination}")
        except Exception:
            failed += 1
            display_progress(f"附件 {index} 下載失敗：{filename}")
            LOGGER.exception("附件 %d 下載失敗：%s", index, filename)
    return downloaded, failed


def process_message(message, source_id: str, processed_ids: set[str], local_msg_path: Path | None = None) -> dict[str, int]:
    result = {"matched": 0, "already_processed": 0, "skipped_reply_forward": 0, "skipped_no_supported_attachment": 0, "processed": 0, "downloaded": 0, "failed": 0, "fallback": 0, "msg_saved": 0}
    subject = str(getattr(message, "Subject", "") or "").strip()
    display_progress(f"檢查郵件：{subject or '[沒有主旨]'}")
    if not any(keyword.casefold() in subject.casefold() for keyword in SUBJECT_KEYWORDS):
        display_progress("略過：郵件主旨冇指定關鍵字。")
        return result
    result["matched"] = 1
    if is_reply_or_forward(message, subject):
        result["skipped_reply_forward"] = 1
        display_progress("略過：郵件係回覆或轉寄。")
        return result
    if not REPROCESS_ALREADY_PROCESSED and source_id and source_id in processed_ids:
        result["already_processed"] = 1
        display_progress("略過：郵件之前已經處理。")
        return result
    supported = scan_supported_attachments(message)
    if not supported:
        result["skipped_no_supported_attachment"] = 1
        display_progress("略過：郵件冇 PDF 或 XLSX 附件。")
        return result
    folder_name, reference_type = get_destination_folder_name(subject)
    received_time = message_received_time(message)
    mbl_folder = OUTPUT_ROOT / received_time.strftime("%Y-%m-%d") / folder_name
    if reference_type == "FALLBACK":
        result["fallback"] = 1
        display_progress(f"搵唔到有效 MBL、BL 或 HBL，使用完整郵件主旨：{folder_name}")
    else:
        display_progress(f"搵到 {reference_type}：{folder_name}")
    operation_failed = False
    try:
        save_message_file(message, mbl_folder, subject, local_msg_path)
        result["msg_saved"] = 1
    except Exception:
        operation_failed = True
        result["failed"] += 1
        display_progress("MSG 檔案儲存或複製失敗。")
        LOGGER.exception("MSG 檔案儲存或複製失敗。")
    downloaded, failed = download_attachments(message, supported, mbl_folder)
    result["downloaded"] = downloaded
    result["failed"] += failed
    if operation_failed or failed > 0:
        display_progress("因為有檔案處理失敗，郵件唔會標記為已處理。")
        return result
    if source_id:
        processed_ids.add(source_id)
        save_processed_ids(processed_ids)
    result["processed"] = 1
    display_progress("郵件處理完成。")
    return result


def add_result(summary: dict[str, int], result: dict[str, int]) -> None:
    for key, value in result.items():
        summary[key] += value


def process_outlook_inbox(namespace, processed_ids: set[str], summary: dict[str, int]) -> None:
    search_start, search_end_exclusive = get_search_period()
    display_end = (search_end_exclusive - timedelta(days=1)).date()
    display_progress(f"模式：Outlook 收件箱，搜尋日期：{search_start.date()} 至 {display_end}")
    items = namespace.GetDefaultFolder(OL_FOLDER_INBOX).Items
    items.Sort("[ReceivedTime]", True)
    display_progress(f"收件箱共有 {int(items.Count)} 個項目。")
    for item_index in range(1, int(items.Count) + 1):
        try:
            item = items.Item(item_index)
            if getattr(item, "Class", None) != OL_MAIL_ITEM:
                continue
            received_time = outlook_datetime_to_hk_naive(item.ReceivedTime)
            if received_time >= search_end_exclusive:
                continue
            if received_time < search_start:
                display_progress("已到達搜尋日期之前嘅郵件，停止掃描。")
                break
            summary["scanned"] += 1
            entry_id = str(getattr(item, "EntryID", "") or "").strip()
            add_result(summary, process_message(item, entry_id, processed_ids))
        except Exception:
            display_progress(f"處理收件箱第 {item_index} 個項目時發生錯誤。")
            LOGGER.exception("處理收件箱第 %d 個項目時發生錯誤。", item_index)


def msg_file_state_id(msg_path: Path) -> str:
    stat = msg_path.stat()
    return f"MSG::{msg_path.resolve()}::{stat.st_size}::{stat.st_mtime_ns}"


def process_msg_folder(namespace, processed_ids: set[str], summary: dict[str, int]) -> None:
    folder = Path(str(MSG_FOLDER_PATH)).expanduser()
    if not folder.exists():
        raise FileNotFoundError(f"MSG 資料夾不存在：{folder}")
    if not folder.is_dir():
        raise NotADirectoryError(f"MSG 路徑唔係資料夾：{folder}")
    msg_files = sorted((path for path in folder.rglob("*") if path.is_file() and path.suffix.casefold() == ".msg"), key=lambda path: str(path).casefold())
    display_progress("模式：MSG 資料夾，日期設定將會忽略。")
    display_progress(f"搵到 {len(msg_files)} 個 MSG 檔案，包括子資料夾。")
    for index, msg_path in enumerate(msg_files, start=1):
        message = None
        try:
            summary["scanned"] += 1
            display_progress(f"開啟 MSG {index}/{len(msg_files)}：{msg_path}")
            source_id = msg_file_state_id(msg_path)
            message = namespace.OpenSharedItem(str(msg_path.resolve()))
            if getattr(message, "Class", None) != OL_MAIL_ITEM:
                display_progress("略過：呢個 MSG 檔案唔係普通郵件項目。")
                continue
            add_result(summary, process_message(message, source_id, processed_ids, msg_path))
        except Exception:
            display_progress(f"處理 MSG 檔案失敗：{msg_path}")
            LOGGER.exception("處理 MSG 檔案失敗：%s", msg_path)
        finally:
            if message is not None:
                try:
                    message.Close(1)
                except Exception:
                    pass


def main() -> None:
    summary = {"scanned": 0, "matched": 0, "already_processed": 0, "skipped_reply_forward": 0, "skipped_no_supported_attachment": 0, "processed": 0, "downloaded": 0, "failed": 0, "fallback": 0, "msg_saved": 0}
    display_progress("Outlook 附件下載程式開始執行。")
    pythoncom.CoInitialize()
    try:
        display_progress("正在連接 Outlook，請等候...")
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        processed_ids = load_processed_ids()
        if MSG_FOLDER_PATH is None or not str(MSG_FOLDER_PATH).strip():
            process_outlook_inbox(namespace, processed_ids, summary)
        else:
            process_msg_folder(namespace, processed_ids, summary)
        result_message = (
            "處理完成。\n"
            f"已掃描項目：{summary['scanned']}\n"
            f"符合關鍵字郵件：{summary['matched']}\n"
            f"略過已處理郵件：{summary['already_processed']}\n"
            f"略過回覆或轉寄郵件：{summary['skipped_reply_forward']}\n"
            f"略過冇 PDF 或 XLSX 附件郵件：{summary['skipped_no_supported_attachment']}\n"
            f"今次處理郵件：{summary['processed']}\n"
            f"已儲存或複製 MSG：{summary['msg_saved']}\n"
            f"成功下載附件：{summary['downloaded']}\n"
            f"檔案處理失敗：{summary['failed']}\n"
            f"使用主旨作資料夾名稱：{summary['fallback']}"
        )
        print("\n" + "=" * 60)
        print(result_message)
        print("=" * 60)
        LOGGER.info(result_message.replace("\n", " | "))
    except Exception as exc:
        LOGGER.exception("程式執行失敗。")
        display_progress(f"程式執行失敗：{type(exc).__name__}: {exc}")
    finally:
        pythoncom.CoUninitialize()
        LOGGER.info("程式執行完畢")
        for handler in LOGGER.handlers:
            try:
                handler.flush()
            except Exception:
                pass


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        LOGGER.error("嚴重錯誤：%s", traceback.format_exc())
        display_progress(f"發生嚴重錯誤：{type(exc).__name__}: {exc}")
    finally:
        try:
            input("\n按 Enter 鍵關閉視窗...")
        except Exception:
            pass
