import os
import io
import csv
import sys
import json
import base64
import re
import time
import socket
from datetime import datetime
from zoneinfo import ZoneInfo
from typing import List, Dict, Any
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaIoBaseUpload
from googleapiclient.errors import HttpError

# =========================
# CONFIGURAÇÕES
# =========================
FOLDER_ID = "1f5Z0f73IZD4rBEssNb9OVtADLVZzttaF"
DEST_FOLDER_ID = "1QvBJJDh5UzcewdU8EYl2OscVEYkhWtO-"
DEST_CSV_NAME = "COMPILADO.csv"

CSV_START_ROW = 3
CSV_START_COL = 1  # A

SOURCE_SPREADSHEET_IDS = [
    "1OTHF2ytEOjGgfE49paARXkz9GjaklOQC_UhiXwUjC2E",
    "1rj2V7CxbZwkan63eCeLkH9G00Gi041IZNC6vwEgq6yI",
    "1oS619l3x_D1mXkvDpw8vs91G6ipZmsK83JqEIwPj7Uk",
    "1FO5tyhXygbbzSmmTGdnm45j4DD_rRFQgEheN8T8Wy70",
    "1dNwj8qWTl1k92PxI9iXwaNZYITnxuKP-kOF1QnZK3Iw",
    "1NV0oObhLHAqnSpJKmeBBHQQxcxwlRh14TKQwO561GEw",
    "1rzT8o6XZi4v8j7CYLky3BD3sT5IPjv1PRb45ipBfbw4",
    "1sGHf-zWXoxjnO20QBw2KWX39BSCzT8rzHdEz1hL7jyU",
    "1gN2tR_LCuRnVCQ9tm2UURnVuMlJPVNEjvmo02TwFQCI",
    "1XmpY8mqkRou-CRY68j1ljHH8W8zcROy7wnwMMSfbV7o",
    "1bqGmVwMVvWP7KtyE3gDsLyOtV8Zwvo76AY49HJI7QLk",
]
SOURCE_SHEET_NAME = "Plan_Principal"
SOURCE_RANGE_A1 = "B5:BX"

SCOPES = [
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/spreadsheets.readonly",
]
LOCAL_CREDENTIALS_FILE = "service_account.json"

API_TIMEOUT_SECONDS = 300
API_MAX_RETRIES = 5

socket.setdefaulttimeout(API_TIMEOUT_SECONDS)


# =========================
# EXECUÇÃO COM RETRY
# =========================
def execute_with_retries(request, description: str = "requisição"):
    last_error = None
    for attempt in range(API_MAX_RETRIES):
        try:
            return request.execute(num_retries=2)
        except HttpError as e:
            last_error = e
            status = getattr(e.resp, "status", None)
            retryable = status in {429, 500, 502, 503, 504}
            if not retryable or attempt == API_MAX_RETRIES - 1:
                raise
            wait_seconds = 2 ** attempt
            print(f"Falha HTTP em {description} (status={status}). Tentando novamente em {wait_seconds}s...")
            time.sleep(wait_seconds)
        except (TimeoutError, socket.timeout, OSError) as e:
            last_error = e
            if attempt == API_MAX_RETRIES - 1:
                raise
            wait_seconds = 2 ** attempt
            print(f"Timeout/erro de rede em {description}. Tentando novamente em {wait_seconds}s...")
            time.sleep(wait_seconds)
    raise last_error


# =========================
# CSV - AUMENTA LIMITE
# =========================
def configure_csv_field_limit():
    limit = sys.maxsize
    while True:
        try:
            csv.field_size_limit(limit)
            configured = csv.field_size_limit()
            print(f"Limite do CSV configurado para: {configured}")
            return configured
        except OverflowError:
            limit //= 10

configure_csv_field_limit()


# =========================
# UTILITÁRIOS
# =========================
def cell_has_value(cell: Any) -> bool:
    if cell is None:
        return False
    return str(cell).strip() != ""

def row_has_any_value(row: List[Any]) -> bool:
    return any(cell_has_value(cell) for cell in row)

def remove_fully_blank_rows(values: List[List[Any]]) -> List[List[Any]]:
    return [row for row in values if row_has_any_value(row)]

def filter_rows_where_first_column_has_value(values: List[List[Any]]) -> List[List[Any]]:
    return [row for row in values if cell_has_value(row[0] if row else "")]

def column_letter_to_number(letter: str) -> int:
    result = 0
    for char in letter.upper():
        result = result * 26 + (ord(char) - ord("A") + 1)
    return result

def get_range_width(a1_range: str) -> int:
    match = re.match(r"([A-Z]+)\d*:([A-Z]+)", a1_range.upper())
    if not match:
        raise ValueError(f"Não foi possível calcular a largura do range: {a1_range}")
    start_col = column_letter_to_number(match.group(1))
    end_col = column_letter_to_number(match.group(2))
    return end_col - start_col + 1

def pad_rows_to_width(values: List[List[Any]], width: int) -> List[List[Any]]:
    return [
        (list(row) + [""] * (width - len(row)))[:width]
        for row in values
    ]


# =========================
# NORMALIZAÇÃO NUMÉRICA
# =========================
def is_grouped_thousands(value: str, sep: str) -> bool:
    parts = value.split(sep)
    if len(parts) <= 1:
        return False
    if not all(part.isdigit() for part in parts):
        return False
    if not (1 <= len(parts[0]) <= 3):
        return False
    return all(len(part) == 3 for part in parts[1:])

def normalize_numeric_string(value: Any):
    if value is None:
        return ""
    if not isinstance(value, str):
        return value
    original = value
    s = value.strip().replace("\u00A0", " ")
    if s == "" or s.startswith("'"):
        return s[1:].strip() if s.startswith("'") else s
    is_percent = s.endswith("%")
    if is_percent:
        s = s[:-1].strip()
    s = s.replace("R$", "").replace("$", "").strip()
    negative = False
    if s.startswith("(") and s.endswith(")"):
        negative, s = True, s[1:-1].strip()
    if s.startswith("-"):
        negative, s = True, s[1:].strip()
    s = s.replace(" ", "")
    if not re.fullmatch(r"[\d\.,]+", s):
        return original
    if re.fullmatch(r"\d+", s) and len(s) > 1 and s.startswith("0") and not is_percent:
        return original
    if "." in s and "," in s:
        last_dot, last_comma = s.rfind("."), s.rfind(",")
        if last_comma > last_dot:
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    elif "," in s:
        if is_grouped_thousands(s, ","):
            s = s.replace(",", "")
        elif s.count(",") == 1:
            left, right = s.split(",")
            if right.isdigit() and 1 <= len(right) <= 6:
                s = left + "." + right
            else:
                return original
        else:
            return original
    elif "." in s:
        if is_grouped_thousands(s, "."):
            s = s.replace(".", "")
        elif s.count(".") == 1:
            left, right = s.split(".")
            if right.isdigit() and 1 <= len(right) <= 6:
                pass  # já está no formato correto
            else:
                return original
        else:
            return original
    try:
        number = float(s) if "." in s else int(s)
        if negative:
            number = -number
        if is_percent:
            return f"{number}%"
        return number
    except ValueError:
        return original


# =========================
# AUTENTICAÇÃO
# =========================
def get_credentials() -> Credentials:
    credentials_b64 = os.getenv("GOOGLE_CREDENTIALS_B64")
    if credentials_b64:
        print("Usando credenciais da variável GOOGLE_CREDENTIALS_B64...")
        credentials_info = json.loads(base64.b64decode(credentials_b64).decode("utf-8"))
        return Credentials.from_service_account_info(credentials_info, scopes=SCOPES)
    if os.path.exists(LOCAL_CREDENTIALS_FILE):
        print(f"Usando credenciais do arquivo local: {LOCAL_CREDENTIALS_FILE}")
        return Credentials.from_service_account_file(LOCAL_CREDENTIALS_FILE, scopes=SCOPES)
    raise FileNotFoundError(
        "Credenciais não encontradas. Defina GOOGLE_CREDENTIALS_B64 ou adicione service_account.json."
    )

def get_services():
    creds = get_credentials()
    drive_service = build("drive", "v3", credentials=creds)
    sheets_service = build("sheets", "v4", credentials=creds)
    return drive_service, sheets_service


# =========================
# GOOGLE DRIVE - LEITURA
# =========================
def list_csv_files_in_folder(drive_service, folder_id: str) -> List[Dict[str, str]]:
    files = []
    page_token = None
    query = f"'{folder_id}' in parents and trashed = false and name contains '.csv'"
    while True:
        response = execute_with_retries(
            drive_service.files().list(
                q=query,
                fields="nextPageToken, files(id, name, mimeType)",
                pageToken=page_token,
                pageSize=1000,
                supportsAllDrives=True,
                includeItemsFromAllDrives=True
            ),
            description="listagem de CSVs no Drive"
        )
        files.extend(f for f in response.get("files", []) if f["name"].lower().endswith(".csv"))
        page_token = response.get("nextPageToken")
        if not page_token:
            break
    files.sort(key=lambda x: x["name"].lower())
    return files

def download_csv_content(drive_service, file_id: str) -> str:
    request = drive_service.files().get_media(fileId=file_id)
    buffer = io.BytesIO()
    downloader = MediaIoBaseDownload(buffer, request)
    done = False
    while not done:
        _, done = downloader.next_chunk(num_retries=2)
    raw_content = buffer.getvalue()
    for encoding in ("utf-8-sig", "utf-8", "latin-1"):
        try:
            return raw_content.decode(encoding)
        except UnicodeDecodeError:
            continue
    return raw_content.decode("utf-8", errors="replace")


# =========================
# GOOGLE DRIVE - UPLOAD
# =========================
def find_existing_file_in_folder(drive_service, folder_id: str, filename: str) -> str | None:
    query = (
        f"'{folder_id}' in parents and trashed = false "
        f"and name = '{filename}'"
    )
    response = execute_with_retries(
        drive_service.files().list(
            q=query,
            fields="files(id, name)",
            supportsAllDrives=True,
            includeItemsFromAllDrives=True
        ),
        description=f"busca de arquivo existente '{filename}'"
    )
    files = response.get("files", [])
    return files[0]["id"] if files else None

def upload_csv_to_drive(
    drive_service,
    folder_id: str,
    filename: str,
    rows: List[List[Any]]
):
    buffer = io.StringIO()
    writer = csv.writer(buffer, lineterminator="\n")
    for row in rows:
        writer.writerow([str(cell) if cell is not None else "" for cell in row])
    csv_bytes = buffer.getvalue().encode("utf-8-sig")  # BOM para compatibilidade com Excel
    media = MediaIoBaseUpload(
        io.BytesIO(csv_bytes),
        mimetype="text/csv",
        resumable=True
    )
    existing_id = find_existing_file_in_folder(drive_service, folder_id, filename)
    if existing_id:
        print(f"Arquivo '{filename}' já existe (ID: {existing_id}). Substituindo...")
        execute_with_retries(
            drive_service.files().update(
                fileId=existing_id,
                media_body=media,
                supportsAllDrives=True
            ),
            description=f"atualização do arquivo '{filename}'"
        )
        print(f"Arquivo '{filename}' atualizado com sucesso.")
    else:
        print(f"Criando novo arquivo '{filename}' na pasta {folder_id}...")
        file_metadata = {
            "name": filename,
            "parents": [folder_id],
            "mimeType": "text/csv"
        }
        execute_with_retries(
            drive_service.files().create(
                body=file_metadata,
                media_body=media,
                fields="id",
                supportsAllDrives=True
            ),
            description=f"criação do arquivo '{filename}'"
        )
        print(f"Arquivo '{filename}' criado com sucesso.")


# =========================
# CSV / CONSOLIDAÇÃO
# =========================
def detect_csv_dialect(csv_text: str):
    sample = csv_text[:10000]
    try:
        return csv.Sniffer().sniff(sample, delimiters=",;\t|")
    except Exception:
        class SimpleDialect(csv.Dialect):
            delimiter = ","
            quotechar = '"'
            doublequote = True
            skipinitialspace = False
            lineterminator = "\n"
            quoting = csv.QUOTE_MINIMAL
        return SimpleDialect

def parse_csv_text(csv_text: str) -> List[List[str]]:
    dialect = detect_csv_dialect(csv_text)
    reader = csv.reader(io.StringIO(csv_text, newline=""), dialect=dialect)
    return [
        [str(cell) for cell in row]
        for row in reader
        if row_has_any_value(row)
    ]

def normalize_header(header: List[str]) -> List[str]:
    return [str(col).strip().lower() for col in header]

def merge_csvs(file_contents: List[str]) -> List[List[str]]:
    merged_rows: List[List[str]] = []
    first_header_normalized = None
    for content in file_contents:
        rows = parse_csv_text(content)
        if not rows:
            continue
        header_normalized = normalize_header(rows[0])
        data_rows = rows[1:] if len(rows) > 1 else []
        if first_header_normalized is None:
            first_header_normalized = header_normalized
            merged_rows.append(rows[0])
            merged_rows.extend(data_rows)
        elif header_normalized == first_header_normalized:
            merged_rows.extend(data_rows)
        else:
            merged_rows.extend(rows)
    if not merged_rows:
        return []
    max_cols = max(len(row) for row in merged_rows)
    return [row + [""] * (max_cols - len(row)) for row in merged_rows]


# =========================
# GOOGLE SHEETS - LEITURA
# =========================
def get_sheet_range_values(
    sheets_service,
    spreadsheet_id: str,
    sheet_name: str,
    range_a1: str
) -> List[List[Any]]:
    full_range = f"{sheet_name}!{range_a1}"
    response = execute_with_retries(
        sheets_service.spreadsheets().values().get(
            spreadsheetId=spreadsheet_id,
            range=full_range,
            valueRenderOption="FORMATTED_VALUE",
            dateTimeRenderOption="FORMATTED_STRING",
            majorDimension="ROWS"
        ),
        description=f"leitura de {full_range}"
    )
    return response.get("values", [])

def collect_source_sheets_data(
    sheets_service,
    spreadsheet_ids: List[str],
    sheet_name: str,
    range_a1: str
) -> List[List[Any]]:
    all_rows = []
    expected_width = get_range_width(range_a1)
    for spreadsheet_id in spreadsheet_ids:
        try:
            print(f"Lendo {sheet_name}!{range_a1} da planilha {spreadsheet_id}...")
            rows = get_sheet_range_values(sheets_service, spreadsheet_id, sheet_name, range_a1)
            if not rows:
                print(f" - Nenhum dado encontrado em {spreadsheet_id}")
                continue
            rows = pad_rows_to_width(rows, expected_width)
            rows = remove_fully_blank_rows(rows)
            rows = filter_rows_where_first_column_has_value(rows)
            print(f" - {len(rows)} linha(s) aproveitada(s)")
            all_rows.extend(rows)
        except Exception as e:
            print(f" - Erro ao ler {spreadsheet_id}: {e}")
    return all_rows


# =========================
# NORMALIZAÇÃO DE DADOS
# =========================
def normalize_rows(values: List[List[Any]], skip_first_row: bool = False) -> List[List[Any]]:
    result = []
    for i, row in enumerate(values):
        if skip_first_row and i == 0:
            result.append(row)
        else:
            result.append([normalize_numeric_string(cell) for cell in row])
    return result


# =========================
# MAIN
# =========================
def main():
    drive_service, sheets_service = get_services()
    all_rows: List[List[Any]] = []

    # -------------------------------------------------
    # 1) CONSOLIDA CSVs
    # -------------------------------------------------
    print("Listando arquivos CSV na pasta...")
    files = list_csv_files_in_folder(drive_service, FOLDER_ID)

    if files:
        print(f"{len(files)} arquivo(s) CSV encontrado(s):")
        for f in files:
            print(f" - {f['name']}")

        csv_contents = []
        for f in files:
            print(f"Baixando: {f['name']}")
            csv_contents.append(download_csv_content(drive_service, f["id"]))

        print("Mesclando arquivos CSV...")
        merged_csv_data = merge_csvs(csv_contents)

        if merged_csv_data:
            print(f"Linhas dos CSVs antes da limpeza: {len(merged_csv_data)}")
            merged_csv_data = remove_fully_blank_rows(merged_csv_data)
            merged_csv_data = normalize_rows(merged_csv_data, skip_first_row=True)
            print(f"Linhas dos CSVs após limpeza: {len(merged_csv_data)}")
            all_rows.extend(merged_csv_data)
        else:
            print("Nenhum dado útil encontrado nos CSVs.")
    else:
        print("Nenhum arquivo CSV encontrado na pasta.")

    # -------------------------------------------------
    # 2) LÊ PLAN_PRINCIPAL!B5:BX DAS PLANILHAS DE ORIGEM
    # -------------------------------------------------
    print("Coletando dados das planilhas de origem...")
    source_raw_rows = collect_source_sheets_data(
        sheets_service=sheets_service,
        spreadsheet_ids=SOURCE_SPREADSHEET_IDS,
        sheet_name=SOURCE_SHEET_NAME,
        range_a1=SOURCE_RANGE_A1
    )

    if source_raw_rows:
        print(f"Linhas coletadas das planilhas de origem: {len(source_raw_rows)}")
        source_raw_rows = normalize_rows(source_raw_rows, skip_first_row=False)
        source_raw_rows = remove_fully_blank_rows(source_raw_rows)
        print(f"Linhas das planilhas de origem após limpeza: {len(source_raw_rows)}")
        all_rows.extend(source_raw_rows)
    else:
        print("Nenhum dado encontrado nas planilhas de origem.")

    # -------------------------------------------------
    # 3) SALVA COMO CSV NO GOOGLE DRIVE
    # -------------------------------------------------
    if not all_rows:
        print("Nenhum dado para salvar. Encerrando.")
        return

    print(f"Total de linhas a salvar: {len(all_rows)}")
    print(f"Fazendo upload de '{DEST_CSV_NAME}' para a pasta {DEST_FOLDER_ID}...")
    upload_csv_to_drive(
        drive_service=drive_service,
        folder_id=DEST_FOLDER_ID,
        filename=DEST_CSV_NAME,
        rows=all_rows
    )
    print("Processo concluído com sucesso.")


if __name__ == "__main__":
    main()
