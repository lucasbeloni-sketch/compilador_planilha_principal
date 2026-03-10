import os
import io
import csv
import sys
import json
import base64
import re
import time
import socket
from datetime import datetime, date
from zoneinfo import ZoneInfo
from typing import List, Dict, Any

from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
from googleapiclient.errors import HttpError


# =========================
# CONFIGURAÇÕES
# =========================
FOLDER_ID = "1f5Z0f73IZD4rBEssNb9OVtADLVZzttaF"
DEST_SPREADSHEET_ID = "1B_ZAktVrIoY_qGg9vhjMabmNqGMeHODtWPR8nmFp61A"
DEST_SHEET_NAME = "Planejamento"

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
SOURCE_START_COL_IN_DEST = 1  # A

# Colunas da origem que devem virar data real
SOURCE_DATE_COLUMNS_LETTERS = ["B", "BR", "BS", "BT"]

SCOPES = [
    "https://www.googleapis.com/auth/drive.readonly",
    "https://www.googleapis.com/auth/spreadsheets",
]

LOCAL_CREDENTIALS_FILE = "service_account.json"
WRITE_CHUNK_SIZE = 3000
FORMAT_CHUNK_ROWS = 5000
API_TIMEOUT_SECONDS = 300
API_MAX_RETRIES = 5

SHEETS_DATE_EPOCH = date(1899, 12, 30)

# Define timeout global antes de criar os serviços Google
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
            print(
                f"Falha HTTP em {description} (status={status}). "
                f"Tentando novamente em {wait_seconds}s..."
            )
            time.sleep(wait_seconds)

        except (TimeoutError, socket.timeout, OSError) as e:
            last_error = e

            if attempt == API_MAX_RETRIES - 1:
                raise

            wait_seconds = 2 ** attempt
            print(
                f"Timeout/erro de rede em {description}. "
                f"Tentando novamente em {wait_seconds}s..."
            )
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
    filtered = []

    for row in values:
        first_cell = row[0] if row else ""
        if cell_has_value(first_cell):
            filtered.append(row)

    return filtered


def column_number_to_letter(n: int) -> str:
    result = ""
    while n > 0:
        n, remainder = divmod(n - 1, 26)
        result = chr(65 + remainder) + result
    return result


def column_letter_to_number(letter: str) -> int:
    result = 0
    for char in letter.upper():
        result = result * 26 + (ord(char) - ord("A") + 1)
    return result


def get_range_start_column_letter(a1_range: str) -> str:
    match = re.match(r"([A-Z]+)\d*:([A-Z]+)", a1_range.upper())
    if not match:
        raise ValueError(f"Não foi possível identificar a coluna inicial do range: {a1_range}")
    return match.group(1)


def get_range_width(a1_range: str) -> int:
    match = re.match(r"([A-Z]+)\d*:([A-Z]+)", a1_range.upper())
    if not match:
        raise ValueError(f"Não foi possível calcular a largura do range: {a1_range}")

    start_col = column_letter_to_number(match.group(1))
    end_col = column_letter_to_number(match.group(2))
    return end_col - start_col + 1


def get_relative_column_indexes_for_range(
    column_letters: List[str],
    a1_range: str
) -> List[int]:
    start_col_letter = get_range_start_column_letter(a1_range)
    start_col_number = column_letter_to_number(start_col_letter)

    indexes = []
    for col_letter in column_letters:
        absolute_col_number = column_letter_to_number(col_letter)
        relative_idx = absolute_col_number - start_col_number

        if relative_idx < 0:
            raise ValueError(
                f"A coluna {col_letter} está antes do início do range {a1_range}."
            )

        indexes.append(relative_idx)

    return indexes


def pad_rows_to_width(values: List[List[Any]], width: int) -> List[List[Any]]:
    padded = []
    for row in values:
        row_list = list(row)
        if len(row_list) < width:
            row_list = row_list + [""] * (width - len(row_list))
        else:
            row_list = row_list[:width]
        padded.append(row_list)
    return padded


def iter_row_chunks(start_row: int, num_rows: int, chunk_size: int):
    for offset in range(0, num_rows, chunk_size):
        yield start_row + offset, min(chunk_size, num_rows - offset)


# =========================
# AUTENTICAÇÃO
# =========================
def get_credentials() -> Credentials:
    credentials_b64 = os.getenv("GOOGLE_CREDENTIALS_B64")

    if credentials_b64:
        print("Usando credenciais da variável GOOGLE_CREDENTIALS_B64...")
        credentials_info = json.loads(
            base64.b64decode(credentials_b64).decode("utf-8")
        )
        return Credentials.from_service_account_info(
            credentials_info,
            scopes=SCOPES
        )

    if os.path.exists(LOCAL_CREDENTIALS_FILE):
        print(f"Usando credenciais do arquivo local: {LOCAL_CREDENTIALS_FILE}")
        return Credentials.from_service_account_file(
            LOCAL_CREDENTIALS_FILE,
            scopes=SCOPES
        )

    raise FileNotFoundError(
        "Credenciais não encontradas. Defina a variável GOOGLE_CREDENTIALS_B64 "
        "ou adicione o arquivo service_account.json."
    )


def get_services():
    creds = get_credentials()
    drive_service = build("drive", "v3", credentials=creds)
    sheets_service = build("sheets", "v4", credentials=creds)
    return drive_service, sheets_service


# =========================
# GOOGLE DRIVE
# =========================
def list_csv_files_in_folder(drive_service, folder_id: str) -> List[Dict[str, str]]:
    files = []
    page_token = None

    query = (
        f"'{folder_id}' in parents "
        f"and trashed = false "
        f"and name contains '.csv'"
    )

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

        batch_files = response.get("files", [])
        for file in batch_files:
            if file["name"].lower().endswith(".csv"):
                files.append(file)

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
        try:
            _, done = downloader.next_chunk(num_retries=2)
        except (HttpError, TimeoutError, socket.timeout, OSError):
            raise

    raw_content = buffer.getvalue()

    for encoding in ("utf-8-sig", "utf-8", "latin-1"):
        try:
            return raw_content.decode(encoding)
        except UnicodeDecodeError:
            continue

    return raw_content.decode("utf-8", errors="replace")


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

    rows = []
    for row in reader:
        cleaned_row = [str(cell) for cell in row]
        if row_has_any_value(cleaned_row):
            rows.append(cleaned_row)

    return rows


def normalize_header(header: List[str]) -> List[str]:
    return [str(col).strip().lower() for col in header]


def merge_csvs(file_contents: List[str]) -> List[List[str]]:
    merged_rows: List[List[str]] = []
    first_header = None
    first_header_normalized = None

    for content in file_contents:
        rows = parse_csv_text(content)

        if not rows:
            continue

        header = rows[0]
        data_rows = rows[1:] if len(rows) > 1 else []
        header_normalized = normalize_header(header)

        if first_header is None:
            first_header = header
            first_header_normalized = header_normalized
            merged_rows.append(header)
            merged_rows.extend(data_rows)
            continue

        if header_normalized == first_header_normalized:
            merged_rows.extend(data_rows)
        else:
            merged_rows.extend(rows)

    if not merged_rows:
        return []

    max_cols = max(len(row) for row in merged_rows)
    normalized_rows = [
        row + [""] * (max_cols - len(row))
        for row in merged_rows
    ]

    return normalized_rows


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
            rows = get_sheet_range_values(
                sheets_service=sheets_service,
                spreadsheet_id=spreadsheet_id,
                sheet_name=sheet_name,
                range_a1=range_a1
            )

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
# CONVERSÃO DE TEXTO -> NÚMERO / DATA
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

    if s == "":
        return ""

    if s.startswith("'"):
        s = s[1:].strip()

    is_percent = False
    if s.endswith("%"):
        is_percent = True
        s = s[:-1].strip()

    s = s.replace("R$", "").replace("$", "").strip()

    negative = False
    if s.startswith("(") and s.endswith(")"):
        negative = True
        s = s[1:-1].strip()

    if s.startswith("-"):
        negative = True
        s = s[1:].strip()

    s = s.replace(" ", "")

    if not re.fullmatch(r"[\d\.,]+", s):
        return original

    if re.fullmatch(r"\d+", s) and len(s) > 1 and s.startswith("0") and not is_percent:
        return original

    if "." in s and "," in s:
        last_dot = s.rfind(".")
        last_comma = s.rfind(",")

        if last_comma > last_dot:
            s = s.replace(".", "")
            s = s.replace(",", ".")
        else:
            s = s.replace(",", "")

    elif "," in s:
        if is_grouped_thousands(s, ","):
            s = s.replace(",", "")
        else:
            if s.count(",") == 1:
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
        else:
            if s.count(".") == 1:
                left, right = s.split(".")
                if right.isdigit() and 1 <= len(right) <= 6:
                    s = left + "." + right
                else:
                    return original
            else:
                return original

    try:
        if "." in s:
            number = float(s)
        else:
            number = int(s)

        if negative:
            number = -number

        if is_percent:
            return float(number) / 100

        return number

    except ValueError:
        return original


def convert_display_date_to_serial(value: Any):
    if value is None:
        return ""

    if not isinstance(value, str):
        return value

    s = value.strip()

    if s == "":
        return ""

    if s.startswith("'"):
        s = s[1:].strip()

    match = re.match(r"^(\d{2}/\d{2}/\d{4})(?:\s*-\s*.*)?$", s)
    if not match:
        return value

    try:
        parsed_date = datetime.strptime(match.group(1), "%d/%m/%Y").date()
        return (parsed_date - SHEETS_DATE_EPOCH).days
    except ValueError:
        return value


def convert_csv_rows_for_sheets(values: List[List[Any]]) -> List[List[Any]]:
    converted = []

    for row_idx, row in enumerate(values):
        new_row = []

        for col_idx, cell in enumerate(row):
            if row_idx > 0 and col_idx == 0:
                converted_date = convert_display_date_to_serial(cell)
                if converted_date != cell:
                    new_row.append(converted_date)
                else:
                    new_row.append(normalize_numeric_string(cell))
            else:
                if row_idx == 0:
                    new_row.append(cell)
                else:
                    new_row.append(normalize_numeric_string(cell))

        converted.append(new_row)

    return converted


def convert_source_rows_for_sheets(
    values: List[List[Any]],
    date_column_indexes: List[int]
) -> List[List[Any]]:
    date_column_indexes_set = set(date_column_indexes)
    converted = []

    for row in values:
        new_row = []

        for idx, cell in enumerate(row):
            if idx in date_column_indexes_set:
                converted_date = convert_display_date_to_serial(cell)
                if converted_date != cell:
                    new_row.append(converted_date)
                else:
                    new_row.append(normalize_numeric_string(cell))
            else:
                new_row.append(normalize_numeric_string(cell))

        converted.append(new_row)

    return converted


# =========================
# DETECÇÃO / FORMATAÇÃO DE %
# =========================
def is_percentage_text(value: Any) -> bool:
    if not isinstance(value, str):
        return False

    s = value.strip()
    if s.startswith("'"):
        s = s[1:].strip()

    return bool(re.fullmatch(r"-?\d+(?:[.,]\d+)?%", s))


def detect_percentage_columns(
    values: List[List[Any]],
    skip_first_row: bool = False,
    threshold: float = 0.6
) -> List[int]:
    if not values:
        return []

    data_rows = values[1:] if skip_first_row and len(values) > 1 else values
    if not data_rows:
        return []

    max_cols = max(len(row) for row in data_rows)
    percentage_columns = []

    for col_idx in range(max_cols):
        non_empty_count = 0
        percent_count = 0

        for row in data_rows:
            cell = row[col_idx] if col_idx < len(row) else ""

            if isinstance(cell, str) and cell.strip() != "":
                non_empty_count += 1
                if is_percentage_text(cell):
                    percent_count += 1

        if non_empty_count > 0 and (percent_count / non_empty_count) >= threshold:
            percentage_columns.append(col_idx)

    return percentage_columns


def get_sheet_id(sheets_service, spreadsheet_id: str, sheet_name: str) -> int:
    metadata = execute_with_retries(
        sheets_service.spreadsheets().get(
            spreadsheetId=spreadsheet_id,
            fields="sheets(properties(sheetId,title))"
        ),
        description=f"obtenção do sheetId de {sheet_name}"
    )

    for sheet in metadata.get("sheets", []):
        props = sheet.get("properties", {})
        if props.get("title") == sheet_name:
            return props.get("sheetId")

    raise ValueError(f"Aba '{sheet_name}' não encontrada.")


def get_sheet_properties(sheets_service, spreadsheet_id: str, sheet_name: str) -> Dict[str, Any]:
    metadata = execute_with_retries(
        sheets_service.spreadsheets().get(
            spreadsheetId=spreadsheet_id,
            fields="sheets.properties(sheetId,title,gridProperties)"
        ),
        description=f"obtenção das propriedades da aba {sheet_name}"
    )

    for sheet in metadata.get("sheets", []):
        props = sheet.get("properties", {})
        if props.get("title") == sheet_name:
            return props

    raise ValueError(f"Aba '{sheet_name}' não encontrada.")


def apply_percentage_format(
    sheets_service,
    spreadsheet_id: str,
    sheet_name: str,
    percentage_columns: List[int],
    start_row: int,
    start_col: int,
    num_rows: int
):
    if not percentage_columns or num_rows <= 0:
        return

    sheet_id = get_sheet_id(sheets_service, spreadsheet_id, sheet_name)

    for chunk_start_row, chunk_num_rows in iter_row_chunks(start_row, num_rows, FORMAT_CHUNK_ROWS):
        requests = []

        for col_idx in percentage_columns:
            requests.append({
                "repeatCell": {
                    "range": {
                        "sheetId": sheet_id,
                        "startRowIndex": chunk_start_row - 1,
                        "endRowIndex": chunk_start_row - 1 + chunk_num_rows,
                        "startColumnIndex": start_col - 1 + col_idx,
                        "endColumnIndex": start_col - 1 + col_idx + 1
                    },
                    "cell": {
                        "userEnteredFormat": {
                            "numberFormat": {
                                "type": "PERCENT"
                            }
                        }
                    },
                    "fields": "userEnteredFormat.numberFormat"
                }
            })

        execute_with_retries(
            sheets_service.spreadsheets().batchUpdate(
                spreadsheetId=spreadsheet_id,
                body={"requests": requests}
            ),
            description=(
                f"formatação percentual em {sheet_name} "
                f"(linhas {chunk_start_row}-{chunk_start_row + chunk_num_rows - 1})"
            )
        )


def apply_date_format(
    sheets_service,
    spreadsheet_id: str,
    sheet_name: str,
    date_columns: List[int],
    start_row: int,
    start_col: int,
    num_rows: int
):
    if not date_columns or num_rows <= 0:
        return

    sheet_id = get_sheet_id(sheets_service, spreadsheet_id, sheet_name)

    for chunk_start_row, chunk_num_rows in iter_row_chunks(start_row, num_rows, FORMAT_CHUNK_ROWS):
        requests = []

        for col_idx in date_columns:
            requests.append({
                "repeatCell": {
                    "range": {
                        "sheetId": sheet_id,
                        "startRowIndex": chunk_start_row - 1,
                        "endRowIndex": chunk_start_row - 1 + chunk_num_rows,
                        "startColumnIndex": start_col - 1 + col_idx,
                        "endColumnIndex": start_col - 1 + col_idx + 1
                    },
                    "cell": {
                        "userEnteredFormat": {
                            "numberFormat": {
                                "type": "DATE",
                                "pattern": "dd/mm/yyyy"
                            }
                        }
                    },
                    "fields": "userEnteredFormat.numberFormat"
                }
            })

        execute_with_retries(
            sheets_service.spreadsheets().batchUpdate(
                spreadsheetId=spreadsheet_id,
                body={"requests": requests}
            ),
            description=(
                f"formatação de data em {sheet_name} "
                f"(linhas {chunk_start_row}-{chunk_start_row + chunk_num_rows - 1})"
            )
        )


def sort_planejamento_by_column_a(
    sheets_service,
    spreadsheet_id: str,
    sheet_name: str,
    header_row: int,
    total_written_rows: int,
    last_column_letter: str = "BW"
):
    # total_written_rows inclui a linha de cabeçalho
    if total_written_rows <= 1:
        print("Não há linhas de dados suficientes para ordenar.")
        return

    sheet_id = get_sheet_id(sheets_service, spreadsheet_id, sheet_name)

    data_start_row_index = header_row  # linha 4 em índice 0-based
    data_end_row_index = header_row + total_written_rows - 1  # exclusivo
    end_column_index = column_letter_to_number(last_column_letter)  # exclusivo

    execute_with_retries(
        sheets_service.spreadsheets().batchUpdate(
            spreadsheetId=spreadsheet_id,
            body={
                "requests": [
                    {
                        "sortRange": {
                            "range": {
                                "sheetId": sheet_id,
                                "startRowIndex": data_start_row_index,
                                "endRowIndex": data_end_row_index,
                                "startColumnIndex": 0,              # A
                                "endColumnIndex": end_column_index  # BW
                            },
                            "sortSpecs": [
                                {
                                    "dimensionIndex": 0,             # coluna A
                                    "sortOrder": "ASCENDING"
                                }
                            ]
                        }
                    }
                ]
            }
        ),
        description=f"ordenação de {sheet_name}!A4:{last_column_letter} pela coluna A"
    )


# =========================
# GOOGLE SHEETS - ESCRITA
# =========================
def build_range(sheet_name: str, start_row: int, start_col: int, num_rows: int, num_cols: int) -> str:
    start_col_letter = column_number_to_letter(start_col)
    end_col_letter = column_number_to_letter(start_col + num_cols - 1)
    end_row = start_row + num_rows - 1
    return f"{sheet_name}!{start_col_letter}{start_row}:{end_col_letter}{end_row}"


def clear_target_range(sheets_service, spreadsheet_id: str, sheet_name: str):
    clear_range = f"{sheet_name}!A3:ZZZ"
    execute_with_retries(
        sheets_service.spreadsheets().values().clear(
            spreadsheetId=spreadsheet_id,
            range=clear_range,
            body={}
        ),
        description=f"limpeza da faixa {clear_range}"
    )


def ensure_sheet_has_capacity(
    sheets_service,
    spreadsheet_id: str,
    sheet_name: str,
    required_rows: int,
    required_cols: int
):
    props = get_sheet_properties(sheets_service, spreadsheet_id, sheet_name)
    sheet_id = props["sheetId"]
    grid = props.get("gridProperties", {})

    current_rows = grid.get("rowCount", 0)
    current_cols = grid.get("columnCount", 0)

    requests = []

    if required_rows > current_rows:
        rows_to_add = required_rows - current_rows
        print(f"Adicionando {rows_to_add} linha(s) na aba {sheet_name}...")
        requests.append({
            "appendDimension": {
                "sheetId": sheet_id,
                "dimension": "ROWS",
                "length": rows_to_add
            }
        })

    if required_cols > current_cols:
        cols_to_add = required_cols - current_cols
        print(f"Adicionando {cols_to_add} coluna(s) na aba {sheet_name}...")
        requests.append({
            "appendDimension": {
                "sheetId": sheet_id,
                "dimension": "COLUMNS",
                "length": cols_to_add
            }
        })

    if requests:
        execute_with_retries(
            sheets_service.spreadsheets().batchUpdate(
                spreadsheetId=spreadsheet_id,
                body={"requests": requests}
            ),
            description=f"expansão da grade da aba {sheet_name}"
        )


def write_to_sheet_in_chunks(
    sheets_service,
    spreadsheet_id: str,
    sheet_name: str,
    start_row: int,
    start_col: int,
    values: List[List[Any]],
    chunk_size: int = WRITE_CHUNK_SIZE
):
    if not values:
        print("Nenhum dado para gravar na planilha.")
        return

    total_rows = len(values)
    total_cols = max(len(row) for row in values) if values else 0

    print(f"Total de linhas para gravação: {total_rows}")
    print(f"Total de colunas para gravação: {total_cols}")

    required_end_row = start_row + total_rows - 1
    required_end_col = start_col + total_cols - 1

    ensure_sheet_has_capacity(
        sheets_service=sheets_service,
        spreadsheet_id=spreadsheet_id,
        sheet_name=sheet_name,
        required_rows=required_end_row,
        required_cols=required_end_col
    )

    current_row = start_row

    for i in range(0, total_rows, chunk_size):
        chunk = values[i:i + chunk_size]
        target_range = build_range(
            sheet_name=sheet_name,
            start_row=current_row,
            start_col=start_col,
            num_rows=len(chunk),
            num_cols=total_cols
        )

        print(f"Gravando linhas {i + 1} até {i + len(chunk)} em {target_range}...")

        execute_with_retries(
            sheets_service.spreadsheets().values().update(
                spreadsheetId=spreadsheet_id,
                range=target_range,
                valueInputOption="RAW",
                body={"values": chunk}
            ),
            description=f"gravação em {target_range}"
        )

        current_row += len(chunk)


def write_timestamp_to_c2(sheets_service, spreadsheet_id: str, sheet_name: str):
    timestamp = datetime.now(ZoneInfo("America/Recife")).strftime("%d/%m/%Y %H:%M:%S")

    execute_with_retries(
        sheets_service.spreadsheets().values().update(
            spreadsheetId=spreadsheet_id,
            range=f"{sheet_name}!C2",
            valueInputOption="RAW",
            body={"values": [[timestamp]]}
        ),
        description=f"gravação do timestamp em {sheet_name}!C2"
    )


# =========================
# MAIN
# =========================
def main():
    drive_service, sheets_service = get_services()

    print("Limpando faixa de destino...")
    clear_target_range(sheets_service, DEST_SPREADSHEET_ID, DEST_SHEET_NAME)

    # -------------------------------------------------
    # 1) CONSOLIDA CSVs E ESCREVE A PARTIR DE A3
    # -------------------------------------------------
    print("Listando arquivos CSV na pasta...")
    files = list_csv_files_in_folder(drive_service, FOLDER_ID)

    csv_rows_written = 0

    if files:
        print(f"{len(files)} arquivo(s) CSV encontrado(s):")
        for f in files:
            print(f" - {f['name']}")

        csv_contents = []
        for f in files:
            print(f"Baixando: {f['name']}")
            content = download_csv_content(drive_service, f["id"])
            csv_contents.append(content)

        print("Mesclando arquivos CSV...")
        merged_csv_data = merge_csvs(csv_contents)

        if merged_csv_data:
            print(f"Total final de linhas dos CSVs antes da limpeza: {len(merged_csv_data)}")
            print(f"Total final de colunas dos CSVs: {max(len(row) for row in merged_csv_data)}")

            print("Detectando colunas de porcentagem dos CSVs...")
            csv_percentage_columns = detect_percentage_columns(
                merged_csv_data,
                skip_first_row=True,
                threshold=0.6
            )
            print(f"Colunas de porcentagem dos CSVs: {csv_percentage_columns}")

            print("Convertendo valores dos CSVs...")
            prepared_csv_data = convert_csv_rows_for_sheets(merged_csv_data)

            print("Removendo linhas totalmente em branco dos CSVs...")
            prepared_csv_data = remove_fully_blank_rows(prepared_csv_data)
            print(f"Total final de linhas dos CSVs após remover vazias: {len(prepared_csv_data)}")

            if prepared_csv_data:
                print("Gravando bloco dos CSVs...")
                write_to_sheet_in_chunks(
                    sheets_service=sheets_service,
                    spreadsheet_id=DEST_SPREADSHEET_ID,
                    sheet_name=DEST_SHEET_NAME,
                    start_row=CSV_START_ROW,
                    start_col=CSV_START_COL,
                    values=prepared_csv_data
                )

                csv_rows_written = len(prepared_csv_data)

                if csv_rows_written > 1:
                    print("Aplicando formatação de data na coluna A dos CSVs...")
                    apply_date_format(
                        sheets_service=sheets_service,
                        spreadsheet_id=DEST_SPREADSHEET_ID,
                        sheet_name=DEST_SHEET_NAME,
                        date_columns=[0],
                        start_row=CSV_START_ROW + 1,  # pula cabeçalho
                        start_col=CSV_START_COL,
                        num_rows=csv_rows_written - 1
                    )

                print("Aplicando formatação de porcentagem nos CSVs...")
                apply_percentage_format(
                    sheets_service=sheets_service,
                    spreadsheet_id=DEST_SPREADSHEET_ID,
                    sheet_name=DEST_SHEET_NAME,
                    percentage_columns=csv_percentage_columns,
                    start_row=CSV_START_ROW,
                    start_col=CSV_START_COL,
                    num_rows=csv_rows_written
                )
        else:
            print("Nenhum dado útil encontrado nos CSVs.")
    else:
        print("Nenhum arquivo CSV encontrado na pasta.")

    # -------------------------------------------------
    # 2) LÊ PLAN_PRINCIPAL!B5:BX DAS 11 PLANILHAS
    #    E ESCREVE ABAIXO DO BLOCO DOS CSVs, A PARTIR DA COLUNA A
    #    CONSIDERANDO APENAS LINHAS ONDE A COLUNA B DA ORIGEM TENHA VALOR
    # -------------------------------------------------
    append_start_row = CSV_START_ROW + csv_rows_written
    source_rows_written = 0

    print("Coletando dados das planilhas de origem...")
    source_raw_rows = collect_source_sheets_data(
        sheets_service=sheets_service,
        spreadsheet_ids=SOURCE_SPREADSHEET_IDS,
        sheet_name=SOURCE_SHEET_NAME,
        range_a1=SOURCE_RANGE_A1
    )

    if source_raw_rows:
        print(f"Total de linhas coletadas das planilhas de origem: {len(source_raw_rows)}")

        source_date_column_indexes = get_relative_column_indexes_for_range(
            SOURCE_DATE_COLUMNS_LETTERS,
            SOURCE_RANGE_A1
        )
        print(f"Colunas de data das planilhas de origem: {source_date_column_indexes}")

        print("Detectando colunas de porcentagem das planilhas de origem...")
        source_percentage_columns = detect_percentage_columns(
            source_raw_rows,
            skip_first_row=False,
            threshold=0.6
        )
        print(f"Colunas de porcentagem das planilhas de origem: {source_percentage_columns}")

        print("Convertendo valores das planilhas de origem...")
        prepared_source_rows = convert_source_rows_for_sheets(
            source_raw_rows,
            date_column_indexes=source_date_column_indexes
        )

        print("Removendo linhas totalmente em branco das planilhas de origem...")
        prepared_source_rows = remove_fully_blank_rows(prepared_source_rows)
        print(f"Total final de linhas das planilhas de origem: {len(prepared_source_rows)}")

        if prepared_source_rows:
            print(
                f"Gravando bloco das planilhas de origem a partir de "
                f"{DEST_SHEET_NAME}!{column_number_to_letter(SOURCE_START_COL_IN_DEST)}{append_start_row}..."
            )

            write_to_sheet_in_chunks(
                sheets_service=sheets_service,
                spreadsheet_id=DEST_SPREADSHEET_ID,
                sheet_name=DEST_SHEET_NAME,
                start_row=append_start_row,
                start_col=SOURCE_START_COL_IN_DEST,
                values=prepared_source_rows
            )

            source_rows_written = len(prepared_source_rows)

            print("Aplicando formatação de data nas colunas das planilhas de origem...")
            apply_date_format(
                sheets_service=sheets_service,
                spreadsheet_id=DEST_SPREADSHEET_ID,
                sheet_name=DEST_SHEET_NAME,
                date_columns=source_date_column_indexes,
                start_row=append_start_row,
                start_col=SOURCE_START_COL_IN_DEST,
                num_rows=len(prepared_source_rows)
            )

            print("Aplicando formatação de porcentagem nas planilhas de origem...")
            apply_percentage_format(
                sheets_service=sheets_service,
                spreadsheet_id=DEST_SPREADSHEET_ID,
                sheet_name=DEST_SHEET_NAME,
                percentage_columns=source_percentage_columns,
                start_row=append_start_row,
                start_col=SOURCE_START_COL_IN_DEST,
                num_rows=len(prepared_source_rows)
            )
        else:
            print("Nenhuma linha útil restou nas planilhas de origem após limpeza.")
    else:
        print("Nenhum dado encontrado nas planilhas de origem.")

    total_written_rows = csv_rows_written + source_rows_written

    print("Ordenando intervalo A4:BW pela coluna A...")
    sort_planejamento_by_column_a(
        sheets_service=sheets_service,
        spreadsheet_id=DEST_SPREADSHEET_ID,
        sheet_name=DEST_SHEET_NAME,
        header_row=CSV_START_ROW,
        total_written_rows=total_written_rows,
        last_column_letter="BW"
    )

    print("Gravando timestamp em C2...")
    write_timestamp_to_c2(
        sheets_service=sheets_service,
        spreadsheet_id=DEST_SPREADSHEET_ID,
        sheet_name=DEST_SHEET_NAME
    )

    print("Processo concluído com sucesso.")


if __name__ == "__main__":
    main()
