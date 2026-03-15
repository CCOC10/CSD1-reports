#!/opt/homebrew/bin/python3
from __future__ import annotations

import argparse
import json
import os
import re
import shutil
import subprocess
import tempfile
from copy import copy
from datetime import date, datetime
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from typing import Any
from urllib.parse import quote, unquote, urlparse
from xml.etree import ElementTree as ET
from xml.sax.saxutils import escape
from zipfile import ZIP_DEFLATED, ZipFile

PROJECT_ROOT = Path(__file__).resolve().parents[1]
TEMPLATE_PATH = PROJECT_ROOT / "FORM connect" / "bank-general.docx"
GENERATED_DIR = PROJECT_ROOT / "backend" / "generated"
SAMPLE_PAYLOAD_PATH = PROJECT_ROOT / "backend" / "sample-bank-general.json"
DEFAULT_LEADER_NAME = "พันตำรวจตรีมณเฑียร ธงเทียน"
DEFAULT_LEADER_PHONE = "0616591447"

WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
XML_NS = "http://www.w3.org/XML/1998/namespace"
NS = {"w": WORD_NS}

ET.register_namespace("w", WORD_NS)

TH_MONTHS_FULL = [
    "มกราคม",
    "กุมภาพันธ์",
    "มีนาคม",
    "เมษายน",
    "พฤษภาคม",
    "มิถุนายน",
    "กรกฎาคม",
    "สิงหาคม",
    "กันยายน",
    "ตุลาคม",
    "พฤศจิกายน",
    "ธันวาคม",
]

BANK_PRINT_META = {
    "BBL": {
        "recipient": "กรรมการผู้จัดการใหญ่ ธนาคารกรุงเทพ จำกัด (มหาชน) สำนักงานใหญ่",
        "address": "สำนักงานใหญ่ธนาคารกรุงเทพ จำกัด (มหาชน)",
        "short_name": "กรุงเทพ",
    },
    "KBANK": {
        "recipient": "กรรมการผู้จัดการใหญ่ ธนาคารกสิกรไทย จำกัด (มหาชน) สำนักงานใหญ่",
        "address": "400/22 ถนนพหลโยธิน แขวงสามเสนใน เขตพญาไท กรุงเทพมหานคร",
        "short_name": "กสิกรไทย",
    },
    "KTB": {
        "recipient": "กรรมการผู้จัดการใหญ่ ธนาคารกรุงไทย จำกัด (มหาชน) สำนักงานใหญ่",
        "address": "10 ถนนสุขุมวิท แขวงคลองเตย เขตคลองเตย กรุงเทพมหานคร",
        "short_name": "กรุงไทย",
    },
    "SCB": {
        "recipient": "กรรมการผู้จัดการใหญ่ ธนาคารไทยพาณิชย์ จำกัด (มหาชน) สำนักงานใหญ่",
        "address": "9 ถนนรัชดาภิเษก แขวงจตุจักร เขตจตุจักร กรุงเทพมหานคร 10900",
        "short_name": "ไทยพาณิชย์",
    },
    "TTB": {
        "recipient": "กรรมการผู้จัดการใหญ่ ธนาคารทหารไทยธนชาต จำกัด (มหาชน) สำนักงานใหญ่",
        "address": "3000 ถนนพหลโยธิน แขวงจอมพล เขตจตุจักร กรุงเทพมหานคร 10900",
        "short_name": "ทหารไทยธนชาต",
    },
    "BAY": {
        "recipient": "กรรมการผู้จัดการใหญ่ ธนาคารกรุงศรีอยุธยา จำกัด (มหาชน) สำนักงานใหญ่",
        "address": "1222 ถนนพระรามที่ 3 แขวงบางโพงพาง เขตยานนาวา กรุงเทพ 10120",
        "short_name": "กรุงศรีอยุธยา",
    },
    "GSB": {
        "recipient": "กรรมการผู้จัดการ ธนาคารออมสิน สำนักงานใหญ่",
        "address": "470 ถนนพหลโยธิน แขวงสามเสนใน เขตพญาไท กรุงเทพมหานคร 10400",
        "short_name": "ออมสิน",
    },
}


def as_text(value: Any) -> str:
    return str("" if value is None else value).strip()


def first_filled(*values: Any) -> str:
    for value in values:
        text = as_text(value)
        if text:
            return text
    return ""


def as_array(value: Any) -> list[Any]:
    return value if isinstance(value, list) else []


def sanitize_digits(value: Any) -> str:
    return re.sub(r"\D+", "", as_text(value))


def sanitize_doc_number(value: Any) -> str:
    return re.sub(r"[^0-9]+", "", as_text(value))


def parse_date(value: Any) -> date | None:
    raw = as_text(value)
    if not raw:
        return None
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y"):
        try:
            return datetime.strptime(raw, fmt).date()
        except ValueError:
            continue
    try:
        return datetime.fromisoformat(raw).date()
    except ValueError:
        return None


def format_thai_long_date(value: Any) -> str:
    raw = as_text(value)
    if not raw:
        return "..."
    if any(month in raw for month in TH_MONTHS_FULL):
        return raw
    parsed = parse_date(raw)
    if not parsed:
        return raw
    return f"{parsed.day} {TH_MONTHS_FULL[parsed.month - 1]} {parsed.year + 543}"


def format_thai_issue_date_full(value: Any) -> str:
    raw = as_text(value)
    if raw and "พุทธศักราช" in raw and "เดือน" in raw:
        return raw
    parsed = parse_date(raw) if raw else date.today()
    if not parsed:
        parsed = date.today()
    return (
        f"{parsed.day} เดือน {TH_MONTHS_FULL[parsed.month - 1]} "
        f"พุทธศักราช {parsed.year + 543}"
    )


def get_bank_meta(bank_code: str) -> dict[str, str]:
    code = as_text(bank_code).upper()
    meta = BANK_PRINT_META.get(code, {})
    return {
        "recipient": first_filled(meta.get("recipient"), code),
        "address": first_filled(meta.get("address"), code),
        "short_name": first_filled(meta.get("short_name"), code),
    }


def get_bank_accounts(payload: dict[str, Any]) -> list[dict[str, str]]:
    canonical = as_array(payload.get("bank_accounts"))
    if canonical:
        return [
            {
                "account_no": sanitize_digits(item.get("account_no") or item.get("accountNo")),
                "account_name": first_filled(item.get("account_name"), item.get("accountName")),
            }
            for item in canonical
        ]
    legacy_nos = as_array(payload.get("bankAccounts"))
    legacy_names = as_array(payload.get("bankAccountNames"))
    accounts = []
    for index, account_no in enumerate(legacy_nos):
        accounts.append(
            {
                "account_no": sanitize_digits(account_no),
                "account_name": first_filled(legacy_names[index] if index < len(legacy_names) else ""),
            }
        )
    return accounts


def build_canonical_payload(source: dict[str, Any]) -> dict[str, Any]:
    bank_code = first_filled(source.get("bank_code"), source.get("bankCode")).upper()
    meta = get_bank_meta(bank_code)
    return {
        "unit": first_filled(source.get("unit"), "ชป.15"),
        "bank_code": bank_code,
        "doc_number": sanitize_doc_number(first_filled(source.get("doc_number"), source.get("docNumber"))),
        "issue_date_th_full": format_thai_issue_date_full(
            first_filled(source.get("issue_date_th_full"), source.get("issue_date"), source.get("issueDate"))
        ),
        "bank_name_th": first_filled(source.get("bank_name_th"), meta["short_name"], bank_code),
        "recipient_address": first_filled(source.get("recipient_address"), source.get("bank_address"), meta["address"]),
        "case_headline": first_filled(
            source.get("case_headline"),
            source.get("case_detail"),
            source.get("caseDetail"),
            "คดีที่อยู่ระหว่างการสืบสวนสอบสวน",
        ),
        "bank_request_mode": first_filled(
            source.get("bank_request_mode"),
            source.get("bankRequestMode"),
            "fullaccount",
        ).lower(),
        "bank_accounts": get_bank_accounts(source),
        "period_start_th": format_thai_long_date(
            first_filled(source.get("period_start_th"), source.get("period_start"), source.get("bankDateStart"))
        ),
        "period_end_th": format_thai_long_date(
            first_filled(source.get("period_end_th"), source.get("period_end"), source.get("bankDateEnd"))
        ),
    }


def get_compatibility_report(payload: dict[str, Any]) -> dict[str, Any]:
    blockers: list[str] = []
    warnings: list[str] = []
    mode = first_filled(payload.get("bank_request_mode"), "fullaccount").lower()
    accounts = payload.get("bank_accounts") or []

    if mode != "fullaccount":
        blockers.append(f"template นี้รองรับเฉพาะ fullaccount แต่ได้รับ {mode or 'unknown'}")
    if not payload.get("doc_number"):
        blockers.append("missing doc_number")
    if not payload.get("issue_date_th_full"):
        blockers.append("missing issue_date_th")
    if not payload.get("bank_name_th"):
        blockers.append("missing bank_name")
    if not payload.get("recipient_address"):
        blockers.append("missing bank_address")
    if not payload.get("case_headline"):
        blockers.append("missing case_headline")
    if not payload.get("period_start_th"):
        blockers.append("missing statement_date_from")
    if not payload.get("period_end_th"):
        blockers.append("missing statement_date_to")
    if not accounts:
        blockers.append("missing bank account for accounts_no/accounts_name")
    if len(accounts) > 1:
        warnings.append("template นี้มีช่องบัญชีเดียว ระบบจะใช้เฉพาะบัญชีแรก")
    warnings.append("ผู้ลงนามและท้ายเอกสารยังเป็นข้อความ hardcoded ใน DOCX ฉบับนี้")

    return {
        "compatible": not blockers,
        "blockers": blockers,
        "warnings": warnings,
    }


def build_legacy_placeholder_map(payload: dict[str, Any]) -> dict[str, str]:
    accounts = payload.get("bank_accounts") or []
    first_account = accounts[0] if accounts else {}
    return {
        "doc_number": payload.get("doc_number", ""),
        "issue_date_th": payload.get("issue_date_th_full", ""),
        "bank_name": payload.get("bank_name_th", ""),
        "bank_address": payload.get("recipient_address", ""),
        "case_headline": payload.get("case_headline", ""),
        "accounts_no": first_filled(first_account.get("account_no"), payload.get("account_no")),
        "accounts_name": first_filled(first_account.get("account_name"), payload.get("account_name")),
        "statement_date_from": payload.get("period_start_th", ""),
        "statement_date_to": payload.get("period_end_th", ""),
        "leaderName": first_filled(payload.get("leader_name"), DEFAULT_LEADER_NAME),
        "leaderphoneno": first_filled(payload.get("leader_phone"), DEFAULT_LEADER_PHONE),
    }


def replace_placeholders_in_document_xml(xml_text: str, placeholder_map: dict[str, str]) -> str:
    token_re = re.compile(r"(<w:t(?:\s+[^>]*)?>)(.*?)(</w:t>)", re.DOTALL)
    tokens: list[dict[str, Any]] = []
    combined_parts: list[str] = []
    char_map: list[tuple[int, int]] = []

    for match in token_re.finditer(xml_text):
        open_tag, raw_text, close_tag = match.groups()
        token_index = len(tokens)
        tokens.append(
            {
                "open_tag": open_tag,
                "raw_text": raw_text,
                "close_tag": close_tag,
            }
        )
        for offset, char in enumerate(raw_text):
            char_map.append((token_index, offset))
            combined_parts.append(char)

    combined_text = "".join(combined_parts)
    matches = []
    for match in re.finditer(r"\{\{\s*([A-Za-z0-9_]+)\s*\}\}", combined_text):
        key = match.group(1)
        if key not in placeholder_map:
            continue
        matches.append(
            {
                "key": key,
                "start": match.start(),
                "end": match.end(),
                "replacement": str(placeholder_map.get(key, "")),
            }
        )

    for item in reversed(matches):
        start_ref = char_map[item["start"]] if item["start"] < len(char_map) else None
        end_ref = char_map[item["end"] - 1] if item["end"] - 1 < len(char_map) else None
        if not start_ref or not end_ref:
            continue

        start_token = tokens[start_ref[0]]
        end_token = tokens[end_ref[0]]
        start_text = start_token["raw_text"]
        end_text = end_token["raw_text"]
        prefix = start_text[: start_ref[1]]
        suffix = end_text[end_ref[1] + 1 :]
        replacement = escape(item["replacement"])

        if start_ref[0] == end_ref[0]:
            start_token["raw_text"] = f"{prefix}{replacement}{suffix}"
            continue

        start_token["raw_text"] = f"{prefix}{replacement}"
        for token_index in range(start_ref[0] + 1, end_ref[0]):
            tokens[token_index]["raw_text"] = ""
        end_token["raw_text"] = suffix

    rebuilt_parts: list[str] = []
    last_index = 0
    token_index = 0
    for match in token_re.finditer(xml_text):
        rebuilt_parts.append(xml_text[last_index : match.start()])
        token = tokens[token_index]
        rebuilt_parts.append(f"{token['open_tag']}{token['raw_text']}{token['close_tag']}")
        last_index = match.end()
        token_index += 1
    rebuilt_parts.append(xml_text[last_index:])
    return "".join(rebuilt_parts)


def generate_docx_from_template(
    template_path: Path, output_docx_path: Path, placeholder_map: dict[str, str]
) -> None:
    output_docx_path.parent.mkdir(parents=True, exist_ok=True)
    with ZipFile(template_path, "r") as source_zip, ZipFile(
        output_docx_path, "w", compression=ZIP_DEFLATED
    ) as target_zip:
        for info in source_zip.infolist():
            data = source_zip.read(info.filename)
            if info.filename == "word/document.xml":
                data = replace_placeholders_in_document_xml(
                    data.decode("utf-8"), placeholder_map
                ).encode("utf-8")
            cloned_info = copy(info)
            cloned_info.CRC = 0
            cloned_info.compress_size = 0
            cloned_info.file_size = len(data)
            target_zip.writestr(cloned_info, data)


def find_soffice_binary() -> str:
    env_path = as_text(os.environ.get("SOFFICE_BIN"))
    if env_path and Path(env_path).exists():
        return env_path

    candidates = [
        shutil.which("soffice"),
        shutil.which("libreoffice"),
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",
    ]
    for candidate in candidates:
        if candidate and Path(candidate).exists():
            return str(candidate)
    raise FileNotFoundError(
        "ไม่พบ soffice/libreoffice. ติดตั้ง LibreOffice ก่อน หรือกำหนด SOFFICE_BIN"
    )


def convert_docx_to_pdf(docx_path: Path, output_dir: Path) -> Path:
    soffice_bin = find_soffice_binary()
    output_dir.mkdir(parents=True, exist_ok=True)
    existing_pdf = output_dir / f"{docx_path.stem}.pdf"
    if existing_pdf.exists():
        existing_pdf.unlink()

    with tempfile.TemporaryDirectory(prefix="csd1-lo-profile-") as profile_dir:
        command = [
            soffice_bin,
            "--headless",
            f"-env:UserInstallation=file://{profile_dir}",
            "--convert-to",
            "pdf:writer_pdf_Export",
            "--outdir",
            str(output_dir),
            str(docx_path),
        ]
        result = subprocess.run(command, capture_output=True, text=True)
        if result.returncode != 0:
            raise RuntimeError(
                "LibreOffice แปลง PDF ไม่สำเร็จ: "
                + (result.stderr.strip() or result.stdout.strip() or f"exit {result.returncode}")
            )

    if not existing_pdf.exists():
        raise FileNotFoundError(f"LibreOffice ไม่ได้สร้างไฟล์ PDF: {existing_pdf}")
    return existing_pdf


def build_file_base_name(payload: dict[str, Any]) -> str:
    unit = first_filled(payload.get("unit"), "ชป.15")
    category = first_filled(payload.get("bank_code"), "BANK").upper()
    doc_number = first_filled(payload.get("doc_number"), "0000")
    return f"{unit}-{category}-ตช 0026.21-{doc_number}"


def build_content_disposition(disposition_type: str, file_name: str) -> str:
    suffix = Path(file_name).suffix or ""
    ascii_fallback = f"generated{suffix}"
    encoded_name = quote(file_name, safe="")
    return f"{disposition_type}; filename=\"{ascii_fallback}\"; filename*=UTF-8''{encoded_name}"


def render_bank_general(source_payload: dict[str, Any], *, docx_only: bool = False) -> dict[str, Any]:
    canonical = build_canonical_payload(source_payload)
    report = get_compatibility_report(canonical)
    if not report["compatible"]:
        raise ValueError("; ".join(report["blockers"]))

    file_base_name = build_file_base_name(canonical)
    docx_path = GENERATED_DIR / f"{file_base_name}.docx"
    pdf_path = GENERATED_DIR / f"{file_base_name}.pdf"
    placeholder_map = build_legacy_placeholder_map(canonical)

    generate_docx_from_template(TEMPLATE_PATH, docx_path, placeholder_map)
    resolved_pdf_path: Path | None = None
    if not docx_only:
        resolved_pdf_path = convert_docx_to_pdf(docx_path, GENERATED_DIR)

    return {
        "ok": True,
        "template": str(TEMPLATE_PATH),
        "file_base_name": file_base_name,
        "docx_path": str(docx_path),
        "pdf_path": str(resolved_pdf_path) if resolved_pdf_path else None,
        "docx_only": docx_only,
        "warnings": report["warnings"],
        "placeholder_map": placeholder_map,
    }


class BackendHandler(BaseHTTPRequestHandler):
    server_version = "CSD1DocxPdfBackend/1.0"

    def _write_json(self, status: int, payload: dict[str, Any]) -> None:
        body = json.dumps(payload, ensure_ascii=False, indent=2).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")
        self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
        self.end_headers()
        self.wfile.write(body)

    def _write_bytes(
        self,
        status: int,
        body: bytes,
        content_type: str,
        *,
        content_disposition: str | None = None,
    ) -> None:
        self.send_response(status)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Cache-Control", "no-store")
        if content_disposition:
            self.send_header("Content-Disposition", content_disposition)
        self.end_headers()
        self.wfile.write(body)

    def _resolve_generated_file(self, raw_name: str) -> Path:
        file_name = Path(unquote(raw_name)).name
        if not file_name:
            raise FileNotFoundError("missing generated filename")
        generated_root = GENERATED_DIR.resolve()
        candidate = (generated_root / file_name).resolve()
        if candidate.parent != generated_root or not candidate.exists() or not candidate.is_file():
            raise FileNotFoundError(file_name)
        return candidate

    def _read_json(self) -> dict[str, Any]:
        content_length = int(self.headers.get("Content-Length", "0"))
        raw = self.rfile.read(content_length) if content_length else b"{}"
        if not raw.strip():
            return {}
        return json.loads(raw.decode("utf-8"))

    def do_OPTIONS(self) -> None:
        self._write_json(HTTPStatus.NO_CONTENT, {})

    def do_GET(self) -> None:
        path = urlparse(self.path).path
        if path.startswith("/generated/"):
            try:
                file_path = self._resolve_generated_file(path.removeprefix("/generated/"))
            except FileNotFoundError:
                self._write_json(HTTPStatus.NOT_FOUND, {"ok": False, "error": "generated file not found"})
                return

            if file_path.suffix.lower() == ".pdf":
                self._write_bytes(
                    HTTPStatus.OK,
                    file_path.read_bytes(),
                    "application/pdf",
                    content_disposition=build_content_disposition("inline", file_path.name),
                )
                return

            if file_path.suffix.lower() == ".docx":
                self._write_bytes(
                    HTTPStatus.OK,
                    file_path.read_bytes(),
                    "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    content_disposition=build_content_disposition("attachment", file_path.name),
                )
                return

            self._write_bytes(HTTPStatus.OK, file_path.read_bytes(), "application/octet-stream")
            return

        if path == "/health":
            self._write_json(
                HTTPStatus.OK,
                {
                    "ok": True,
                    "template_exists": TEMPLATE_PATH.exists(),
                    "generated_dir": str(GENERATED_DIR),
                    "soffice": find_soffice_binary() if _soffice_available() else None,
                    "sample_payload": str(SAMPLE_PAYLOAD_PATH),
                },
            )
            return
        self._write_json(
            HTTPStatus.OK,
            {
                "ok": True,
                "service": "CSD1 DOCX/PDF backend",
                "routes": {
                    "GET /health": "health check",
                    "GET /generated/<filename>": "open generated PDF/DOCX file",
                    "POST /api/render/bank-general": "render bank-general.docx -> PDF",
                },
            },
        )

    def do_POST(self) -> None:
        path = urlparse(self.path).path
        try:
            payload = self._read_json()
        except json.JSONDecodeError as error:
            self._write_json(HTTPStatus.BAD_REQUEST, {"ok": False, "error": f"invalid json: {error}"})
            return

        if path == "/api/render/bank-general":
            try:
                result = render_bank_general(payload, docx_only=bool(payload.get("docx_only")))
            except Exception as error:  # noqa: BLE001
                self._write_json(HTTPStatus.BAD_REQUEST, {"ok": False, "error": str(error)})
                return
            self._write_json(HTTPStatus.OK, result)
            return

        self._write_json(HTTPStatus.NOT_FOUND, {"ok": False, "error": f"unknown path: {path}"})


def _soffice_available() -> bool:
    try:
        find_soffice_binary()
        return True
    except FileNotFoundError:
        return False


def run_server(host: str, port: int) -> None:
    GENERATED_DIR.mkdir(parents=True, exist_ok=True)
    server = ThreadingHTTPServer((host, port), BackendHandler)
    print(f"[CSD1 backend] listening on http://{host}:{port}")
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        server.server_close()


def render_from_payload_file(payload_file: Path) -> dict[str, Any]:
    payload = json.loads(payload_file.read_text(encoding="utf-8"))
    return render_bank_general(payload, docx_only=bool(payload.get("docx_only")))


def main() -> int:
    parser = argparse.ArgumentParser(description="DOCX -> PDF backend for CSD1 bank-general template")
    subparsers = parser.add_subparsers(dest="command", required=True)

    serve_parser = subparsers.add_parser("serve", help="start HTTP backend")
    serve_parser.add_argument("--host", default="127.0.0.1")
    serve_parser.add_argument("--port", type=int, default=8765)

    render_parser = subparsers.add_parser("render", help="render from a JSON payload file")
    render_parser.add_argument(
        "--payload",
        type=Path,
        default=SAMPLE_PAYLOAD_PATH,
        help="path to a JSON payload file",
    )
    render_parser.add_argument(
        "--docx-only",
        action="store_true",
        default=False,
        help="generate only DOCX and skip PDF conversion",
    )

    args = parser.parse_args()

    if args.command == "serve":
        run_server(args.host, args.port)
        return 0

    if args.command == "render":
        result = render_bank_general(
            json.loads(args.payload.read_text(encoding="utf-8")),
            docx_only=bool(args.docx_only),
        )
        print(json.dumps(result, ensure_ascii=False, indent=2))
        return 0

    parser.error("unknown command")
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
