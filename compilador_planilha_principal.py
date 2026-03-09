import os
import io
import csv
import sys
import json
import base64
import re
from typing import List, Dict, Any

from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload

# =========================
# CONFIGURAÇÕES
# =========================
FOLDER_ID = "1f5Z0f73IZD4rBEssNb9OVtADLVZzttaF"
SPREADSHEET_ID = "1B_ZAktVrIoY_qGg9vhjMabmNqGMeHODtWPR8nmFp61A"
SHEET_NAME = "Planejamento"
START_ROW = 3
START_COL = 1  # A = 1

SCOPES = [
    "https://www.googleapis.com/auth/drive.readonly",
    "https://www.googleapis.com/auth/spreadsheets",
]

LOCAL_CREDENTIALS_FILE = "service_account.json"
WRITE_CHUNK_SIZE = 3000


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
        response = drive_service.files().list(
            q=query,
            fields="nextPageToken, files(id, name, mimeType)",
            pageToken=page_token,
            pageSize=1000,
            supportsAllDrives=True,
            includeItemsFromAllDrives=True
        ).execute()

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
        _, done = downloader.next_chunk()

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
        rows.append([str(cell) for cell in row])

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
# CONVERSÃO DE TEXTO -> NÚMERO
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

    # Remove apóstrofo inicial
    if s.startswith("'"):
        s = s[1:].strip()

    # Detecta porcentagem
    is_percent = False
    if s.endswith("%"):
        is_percent = True
        s = s[:-1].strip()

    # Remove moeda
    s = s.replace("R$", "").replace("$", "").strip()

    # Negativo
    negative = False
    if s.startswith("(") and s.endswith(")"):
        negative = True
        s = s[1:-1].strip()

    if s.startswith("-"):
        negative = True
        s = s[1:].strip()

    s = s.replace(" ", "")

    # Se tiver letras, mantém original
    if not re.fullmatch(r"[\d\.,]+", s):
        return original

    # Evita converter códigos com zero à esquerda,
    # exceto quando for percentual
    if re.fullmatch(r"\d+", s) and len(s) > 1 and s.startswith("0") and not is_percent:
        return original

    # Tem ponto e vírgula
    if "." in s and "," in s:
        last_dot = s.rfind(".")
        last_comma = s.rfind(",")

        if last_comma > last_dot:
            # BR: 1.234,56
            s = s.replace(".", "")
            s = s.replace(",", ".")
        else:
            # EN: 1,234.56
            s = s.replace(",", "")

    # Só vírgula
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

    # Só ponto
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


def convert_rows_for_sheets(values: List[List[str]]) -> List[List[object]]:
    converted = []
    for row in values:
        converted.append([normalize_numeric_string(cell) for cell in row])
    return converted


# =========================
# DETECÇÃO / FORMATAÇÃO DE %
# =========================
def is_percentage_text(value) -> bool:
    if not isinstance(value, str):
        return False

    s = value.strip()
    if s.startswith("'"):
        s = s[1:].strip()

    return bool(re.fullmatch(r"-?\d+(?:[.,]\d+)?%", s))


def detect_percentage_columns(values: List[List[str]]) -> List[int]:
    if not values or len(values) <= 1:
        return []

    max_cols = max(len(row) for row in values)
    percentage_columns = []
    data_rows = values[1:]  # ignora cabeçalho

    for col_idx in range(max_cols):
        non_empty_count = 0
        percent_count = 0

        for row in data_rows:
            cell = row[col_idx] if col_idx < len(row) else ""

            if isinstance(cell, str) and cell.strip() != "":
                non_empty_count += 1
                if is_percentage_text(cell):
                    percent_count += 1

        if non_empty_count > 0 and percent_count == non_empty_count:
            percentage_columns.append(col_idx)

    return percentage_columns


def get_sheet_id(sheets_service, spreadsheet_id: str, sheet_name: str) -> int:
    metadata = sheets_service.spreadsheets().get(
        spreadsheetId=spreadsheet_id,
        fields="sheets(properties(sheetId,title))"
    ).execute()

    for sheet in metadata.get("sheets", []):
        props = sheet.get("properties", {})
        if props.get("title") == sheet_name:
            return props.get("sheetId")

    raise ValueError(f"Aba '{sheet_name}' não encontrada.")


def apply_percentage_format(
    sheets_service,
    spreadsheet_id: str,
    sheet_name: str,
    percentage_columns: List[int],
    start_row: int,
    start_col: int,
    num_data_rows: int
):
    if not percentage_columns or num_data_rows <= 0:
        return

    sheet_id = get_sheet_id(sheets_service, spreadsheet_id, sheet_name)

    requests = []
    for col_idx in percentage_columns:
        requests.append({
            "repeatCell": {
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": start_row - 1,
                    "endRowIndex": start_row - 1 + num_data_rows,
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

    if requests:
        sheets_service.spreadsheets().batchUpdate(
            spreadsheetId=spreadsheet_id,
            body={"requests": requests}
        ).execute()


# =========================
# GOOGLE SHEETS
# =========================
def column_number_to_letter(n: int) -> str:
    result = ""
    while n > 0:
        n, remainder = divmod(n - 1, 26)
        result = chr(65 + remainder) + result
    return result


def build_range(sheet_name: str, start_row: int, start_col: int, num_rows: int, num_cols: int) -> str:
    start_col_letter = column_number_to_letter(start_col)
    end_col_letter = column_number_to_letter(start_col + num_cols - 1)
    end_row = start_row + num_rows - 1
    return f"{sheet_name}!{start_col_letter}{start_row}:{end_col_letter}{end_row}"


def clear_target_range(sheets_service, spreadsheet_id: str, sheet_name: str):
    clear_range = f"{sheet_name}!A3:ZZZ"
    sheets_service.spreadsheets().values().clear(
        spreadsheetId=spreadsheet_id,
        range=clear_range,
        body={}
    ).execute()


def write_to_sheet_in_chunks(
    sheets_service,
    spreadsheet_id: str,
    sheet_name: str,
    start_row: int,
    start_col: int,
    values: List[List[object]],
    chunk_size: int = WRITE_CHUNK_SIZE
):
    if not values:
        print("Nenhum dado para gravar na planilha.")
        return

    total_rows = len(values)
    total_cols = max(len(row) for row in values) if values else 0

    print(f"Total de linhas para gravação: {total_rows}")
    print(f"Total de colunas para gravação: {total_cols}")

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

        sheets_service.spreadsheets().values().update(
            spreadsheetId=spreadsheet_id,
            range=target_range,
            valueInputOption="RAW",
            body={"values": chunk}
        ).execute()

        current_row += len(chunk)


# =========================
# MAIN
# =========================
def main():
    drive_service, sheets_service = get_services()

    print("Listando arquivos CSV na pasta...")
    files = list_csv_files_in_folder(drive_service, FOLDER_ID)

    if not files:
        print("Nenhum arquivo CSV encontrado na pasta.")
        return

    print(f"{len(files)} arquivo(s) CSV encontrado(s):")
    for f in files:
        print(f" - {f['name']}")

    csv_contents = []
    for f in files:
        print(f"Baixando: {f['name']}")
        content = download_csv_content(drive_service, f["id"])
        csv_contents.append(content)

    print("Mesclando arquivos...")
    merged_data = merge_csvs(csv_contents)

    if not merged_data:
        print("Nenhum dado foi gerado após a mesclagem.")
        return

    print(f"Total final de linhas: {len(merged_data)}")
    print(f"Total final de colunas: {max(len(row) for row in merged_data)}")

    print("Detectando colunas de porcentagem...")
    percentage_columns = detect_percentage_columns(merged_data)
    print(f"Colunas de porcentagem detectadas: {percentage_columns}")

    print("Limpando faixa de destino...")
    clear_target_range(sheets_service, SPREADSHEET_ID, SHEET_NAME)

    print("Convertendo valores para tipos numéricos quando aplicável...")
    prepared_data = convert_rows_for_sheets(merged_data)

    print("Gravando dados na planilha...")
    write_to_sheet_in_chunks(
        sheets_service=sheets_service,
        spreadsheet_id=SPREADSHEET_ID,
        sheet_name=SHEET_NAME,
        start_row=START_ROW,
        start_col=START_COL,
        values=prepared_data
    )

    print("Aplicando formatação de porcentagem...")
    apply_percentage_format(
        sheets_service=sheets_service,
        spreadsheet_id=SPREADSHEET_ID,
        sheet_name=SHEET_NAME,
        percentage_columns=percentage_columns,
        start_row=START_ROW + 1,   # pula o cabeçalho
        start_col=START_COL,
        num_data_rows=len(prepared_data) - 1
    )

    print("Processo concluído com sucesso.")


if __name__ == "__main__":
    main()
