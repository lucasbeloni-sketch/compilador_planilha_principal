import csv
import io
import sys
from typing import List

from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload

# =========================
# CONFIGURAÇÕES
# =========================
FOLDER_ID = "1f5Z0f73IZD4rBEssNb9OVtADLVZzttaF"
SPREADSHEET_ID = "1B_ZAktVrIoY_qGg9vhjMabmNqGMeHODtWPR8nmFp61A"
SHEET_NAME = "Planejamento"
START_CELL = "A3"
SERVICE_ACCOUNT_FILE = "service_account.json"

SCOPES = [
    "https://www.googleapis.com/auth/drive.readonly",
    "https://www.googleapis.com/auth/spreadsheets",
]


def get_services():
    creds = Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE,
        scopes=SCOPES
    )

    drive_service = build("drive", "v3", credentials=creds)
    sheets_service = build("sheets", "v4", credentials=creds)

    return drive_service, sheets_service


def list_csv_files_in_folder(drive_service, folder_id: str) -> List[dict]:
    files = []
    page_token = None

    query = (
        f"'{folder_id}' in parents "
        f"and trashed = false "
        f"and mimeType = 'text/csv'"
    )

    while True:
        response = drive_service.files().list(
            q=query,
            fields="nextPageToken, files(id, name)",
            pageToken=page_token,
            pageSize=1000,
            supportsAllDrives=True,
            includeItemsFromAllDrives=True
        ).execute()

        files.extend(response.get("files", []))
        page_token = response.get("nextPageToken")

        if not page_token:
            break

    # Ordena por nome para manter previsibilidade
    files.sort(key=lambda x: x["name"].lower())
    return files


def download_csv_content(drive_service, file_id: str) -> str:
    request = drive_service.files().get_media(fileId=file_id)
    buffer = io.BytesIO()
    downloader = MediaIoBaseDownload(buffer, request)

    done = False
    while not done:
        _, done = downloader.next_chunk()

    return buffer.getvalue().decode("utf-8-sig", errors="replace")


def parse_csv_text(csv_text: str) -> List[List[str]]:
    reader = csv.reader(io.StringIO(csv_text))
    rows = []
    for row in reader:
        # Mantém inclusive linhas vazias, se necessário
        rows.append([str(cell) for cell in row])
    return rows


def merge_csvs(file_contents: List[str]) -> List[List[str]]:
    merged_rows: List[List[str]] = []
    first_header = None

    for index, content in enumerate(file_contents):
        rows = parse_csv_text(content)

        if not rows:
            continue

        header = rows[0]
        data_rows = rows[1:] if len(rows) > 1 else []

        if index == 0:
            first_header = header
            merged_rows.append(header)
            merged_rows.extend(data_rows)
        else:
            # Se o cabeçalho for igual ao primeiro, ignora
            if header == first_header:
                merged_rows.extend(data_rows)
            else:
                # Se for diferente, inclui tudo para não perder informação
                merged_rows.extend(rows)

    # Padroniza o número de colunas
    max_cols = max((len(row) for row in merged_rows), default=0)
    normalized = [row + [""] * (max_cols - len(row)) for row in merged_rows]

    return normalized


def clear_target_range(sheets_service, spreadsheet_id: str, sheet_name: str):
    # Limpa uma faixa grande a partir de A3
    clear_range = f"{sheet_name}!A3:ZZZ"
    sheets_service.spreadsheets().values().clear(
        spreadsheetId=spreadsheet_id,
        range=clear_range,
        body={}
    ).execute()


def write_to_sheet(
    sheets_service,
    spreadsheet_id: str,
    sheet_name: str,
    start_cell: str,
    values: List[List[str]]
):
    if not values:
        print("Nenhum dado para gravar na planilha.")
        return

    target_range = f"{sheet_name}!{start_cell}"

    sheets_service.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id,
        range=target_range,
        valueInputOption="RAW",
        body={"values": values}
    ).execute()


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

    print("Limpando faixa de destino...")
    clear_target_range(sheets_service, SPREADSHEET_ID, SHEET_NAME)

    print("Gravando dados na planilha...")
    write_to_sheet(
        sheets_service,
        SPREADSHEET_ID,
        SHEET_NAME,
        START_CELL,
        merged_data
    )

    print("Processo concluído com sucesso.")


if __name__ == "__main__":
    main()
