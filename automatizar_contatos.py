#!/usr/bin/env python3
"""Consolida telefones de Assertiva, Consulta Dados e Lemit."""

from __future__ import annotations

import argparse
import csv
import io
import re
import sys
import unicodedata
from collections import OrderedDict
from datetime import datetime
from pathlib import Path

try:
    from openpyxl import Workbook, load_workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter
except ImportError as exc:
    raise SystemExit(
        "Falta instalar a biblioteca openpyxl. Abra o Prompt de Comando e execute: "
        "py -m pip install openpyxl"
    ) from exc


PLATFORM_ORDER = ("ASSERTIVA", "CONSULTA DADOS", "LEMIT")
PLATFORM_COLORS = {
    "ASSERTIVA": ("1F4E78", "D9EAF7"),
    "CONSULTA DADOS": ("C65911", "FCE4D6"),
    "LEMIT": ("548235", "E2F0D9"),
}


def clean_digits(value: object) -> str:
    if value is None or isinstance(value, bool):
        return ""
    if isinstance(value, (int, float)):
        if isinstance(value, float) and not value.is_integer():
            return ""
        value = str(int(value))
    return re.sub(r"\D", "", str(value))


def normalize_cnpj(value: object) -> str:
    digits = clean_digits(value)
    if not digits or len(digits) > 14:
        return ""
    return digits.zfill(14)


def normalize_phone(value: object) -> str:
    digits = clean_digits(value)
    if len(digits) in (12, 13) and digits.startswith("55"):
        digits = digits[2:]
    # A Consulta Dados pode trazer o prefixo de discagem nacional "00"
    # junto do número. Preserve esse valor, conforme o modelo validado.
    if len(digits) == 12 and digits.startswith("00"):
        return digits
    return digits if 8 <= len(digits) <= 11 else ""


def normalize_header(value: object) -> str:
    text = unicodedata.normalize("NFKD", str(value or ""))
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    return re.sub(r"[^A-Z0-9]+", "", text.upper())


def decode_csv(path: Path) -> list[list[str]]:
    raw = path.read_bytes()
    decoded = None
    for encoding in ("utf-8-sig", "cp1252", "latin-1"):
        try:
            decoded = raw.decode(encoding)
            break
        except UnicodeDecodeError:
            continue
    if decoded is None:
        raise ValueError(f"Não consegui ler o texto do arquivo {path.name}.")

    first_line = next((line for line in decoded.splitlines() if line.strip()), "")
    delimiter = ";" if first_line.count(";") >= 3 else None
    if delimiter is None:
        try:
            delimiter = csv.Sniffer().sniff(decoded[:4096], delimiters=",\t;").delimiter
        except csv.Error:
            delimiter = ","
    rows = [row for row in csv.reader(io.StringIO(decoded), delimiter=delimiter) if any(cell.strip() for cell in row)]
    if len(rows) < 2:
        raise ValueError(f"O arquivo {path.name} não tem cabeçalho e linhas de dados.")
    return rows


def find_header(headers: list[object], wanted: tuple[str, ...]) -> int:
    normalized = [normalize_header(value) for value in headers]
    for name in wanted:
        target = normalize_header(name)
        if target in normalized:
            return normalized.index(target)
    raise ValueError(f"Não encontrei a coluna {wanted[0]}.")


def make_company(companies: OrderedDict, cnpj: str) -> dict:
    if cnpj not in companies:
        companies[cnpj] = {"name": "", "phones": {platform: [] for platform in PLATFORM_ORDER}}
    return companies[cnpj]


def put_phone_slots(company: dict, platform: str, slots: list[str | None]) -> None:
    current = company["phones"][platform]
    if not current:
        company["phones"][platform] = slots
        return
    if len(current) < len(slots):
        current.extend([None] * (len(slots) - len(current)))
    extras = []
    for index, phone in enumerate(slots):
        if not phone:
            continue
        if current[index] is None:
            current[index] = phone
        elif current[index] != phone and phone not in current and phone not in extras:
            extras.append(phone)
    current.extend(extras)


def load_assertiva(path: Path, companies: OrderedDict) -> tuple[list[str], set[str], int]:
    rows = decode_csv(path)
    headers = rows[0]
    cnpj_col = find_header(headers, ("CNPJ",))
    name_col = find_header(headers, ("RAZAO", "RAZAO SOCIAL"))
    phone_cols = []
    for index, header in enumerate(headers):
        key = normalize_header(header)
        if key.startswith("TELEFONE") or key == "CELULAR1" or (key.startswith("SOCIO") and "TELEFONE" in key):
            phone_cols.append((index, str(header).strip() or f"TELEFONE{len(phone_cols)+1}"))
    if not phone_cols:
        raise ValueError("Não encontrei colunas de telefone no arquivo da Assertiva.")

    keys: set[str] = set()
    invalid = 0
    for row in rows[1:]:
        cnpj = normalize_cnpj(row[cnpj_col] if cnpj_col < len(row) else None)
        if not cnpj:
            invalid += 1
            continue
        keys.add(cnpj)
        company = make_company(companies, cnpj)
        if not company["name"] and name_col < len(row):
            company["name"] = str(row[name_col] or "").strip()
        slots = [normalize_phone(row[i] if i < len(row) else None) or None for i, _ in phone_cols]
        put_phone_slots(company, "ASSERTIVA", slots)
    return [name for _, name in phone_cols], keys, invalid


def load_consulta(path: Path, companies: OrderedDict) -> tuple[list[str], set[str], int, str]:
    workbook = load_workbook(path, read_only=True, data_only=True)
    preferred = next((sheet for sheet in workbook.worksheets if normalize_header(sheet.title) == "ENRIQUECIMENTO"), None)
    candidates = ([preferred] if preferred else []) + [sheet for sheet in workbook.worksheets if sheet is not preferred]
    selected = None
    columns = None
    for sheet in candidates:
        first = list(next(sheet.iter_rows(min_row=1, max_row=1, values_only=True), ()))
        try:
            columns = (
                find_header(first, ("CNPJ",)),
                find_header(first, ("RAZAO SOCIAL", "RAZAO")),
                find_header(first, ("TELEFONE EMPRESA",)),
                find_header(first, ("SOCIO TELEFONE", "TELEFONE SOCIO")),
            )
        except ValueError:
            continue
        selected = sheet
        break
    if selected is None or columns is None:
        workbook.close()
        raise ValueError("Não encontrei uma aba da Consulta Dados com CNPJ, razão social e telefones da empresa e do sócio.")

    cnpj_col, name_col, company_phone_col, partner_phone_col = columns
    keys: set[str] = set()
    invalid = 0
    for row in selected.iter_rows(min_row=2, values_only=True):
        cnpj = normalize_cnpj(row[cnpj_col] if cnpj_col < len(row) else None)
        if not cnpj:
            invalid += 1
            continue
        keys.add(cnpj)
        company = make_company(companies, cnpj)
        if not company["name"] and name_col < len(row):
            company["name"] = str(row[name_col] or "").strip()
        slots = [
            normalize_phone(row[company_phone_col] if company_phone_col < len(row) else None) or None,
            normalize_phone(row[partner_phone_col] if partner_phone_col < len(row) else None) or None,
        ]
        put_phone_slots(company, "CONSULTA DADOS", slots)
    workbook.close()
    return ["Telefone empresa", "Sócio telefone"], keys, invalid, selected.title


def load_lemit(path: Path, companies: OrderedDict) -> tuple[list[str], set[str], int]:
    rows = decode_csv(path)
    headers = rows[0]
    cnpj_col = find_header(headers, ("CPF/CNPJ", "CNPJ"))
    name_col = find_header(headers, ("NOME",))
    phone_groups = []
    # Alguns CSVs da Lemit vêm com uma célula de cabeçalho vazia ou incorreta
    # entre DDD e FONE. Procure o FONE no mesmo grupo, até duas colunas adiante.
    for index, header in enumerate(headers):
        if normalize_header(header) != "DDD":
            continue
        fone_col = next(
            (candidate for candidate in range(index + 1, min(index + 4, len(headers)))
             if normalize_header(headers[candidate]) == "FONE"),
            None,
        )
        if fone_col is not None:
            phone_groups.append((index, fone_col))
    if not phone_groups:
        raise ValueError("Não encontrei os campos DDD/FONE no arquivo da Lemit.")

    # O modelo de trabalho define a ordem da lista-base (Assertiva/Consulta).
    # No CSV recebido da Lemit, a identificação NOME/CPF-CNPJ fica adiantada
    # em relação aos grupos DDD/FONE depois de algumas linhas. Para preservar
    # a associação validada no modelo, vincule os grupos de telefone pela ordem
    # dos registros à mesma ordem de empresas da lista-base.
    company_order = list(companies.keys())
    name_index = {}
    for known_cnpj, known_company in companies.items():
        normalized_name = normalize_header(known_company["name"])
        if normalized_name:
            name_index.setdefault(normalized_name, set()).add(known_cnpj)

    keys: set[str] = set()
    invalid = 0
    for row_index, row in enumerate(rows[1:]):
        if row_index < len(company_order):
            cnpj = company_order[row_index]
        else:
            # Caso a Lemit traga empresas extras além da lista-base, use a
            # identificação da própria linha somente para esses registros.
            raw_cnpj = row[cnpj_col] if cnpj_col < len(row) else None
            raw_text = str(raw_cnpj or "").strip()
            cnpj = ""
            if re.search(r"[Ee][+-]?\d", raw_text) and name_col < len(row):
                candidates = name_index.get(normalize_header(row[name_col]), set())
                if len(candidates) == 1:
                    cnpj = next(iter(candidates))
            if not cnpj:
                cnpj = normalize_cnpj(raw_cnpj)
            if not cnpj:
                invalid += 1
                continue
        keys.add(cnpj)
        company = make_company(companies, cnpj)
        if not company["name"] and name_col < len(row):
            company["name"] = str(row[name_col] or "").strip()
        slots = []
        for ddd_col, fone_col in phone_groups:
            ddd = clean_digits(row[ddd_col] if ddd_col < len(row) else None)
            fone = clean_digits(row[fone_col] if fone_col < len(row) else None)
            if not fone:
                continue
            combined = fone if len(fone) >= 10 else ddd + fone
            phone = normalize_phone(combined)
            if phone and phone not in slots:
                slots.append(phone)
        put_phone_slots(company, "LEMIT", slots)
    return [f"FONE {i+1}" for i in range(len(phone_groups))], keys, invalid


def format_consulta_phone(phone: str) -> str:
    if len(phone) == 10:
        return f"({phone[:2]}) {phone[2:6]}-{phone[6:]}"
    if len(phone) == 11:
        return f"({phone[:2]}) {phone[2:7]}-{phone[7:]}"
    if len(phone) == 12 and phone.startswith("00"):
        return phone
    return phone


def origin_labels(company: dict, phone: str) -> str:
    return ", ".join(
        platform
        for platform in PLATFORM_ORDER
        if phone in {value for value in company["phones"][platform] if value}
    )


def shared_contacts(companies: OrderedDict) -> list[tuple[str, str, str, str]]:
    results = []
    for cnpj, company in companies.items():
        by_platform = {
            platform: {value for value in company["phones"][platform] if value}
            for platform in PLATFORM_ORDER
        }
        all_numbers = set().union(*by_platform.values())
        for phone in sorted(all_numbers):
            origins = [platform for platform in PLATFORM_ORDER if phone in by_platform[platform]]
            if len(origins) > 1:
                results.append((cnpj, company["name"], phone, ", ".join(origins)))
    return results


def style_group_header(ws, row: int, start: int, end: int, title: str, fill_color: str) -> None:
    fill = PatternFill("solid", fgColor=fill_color)
    for col in range(start, end + 1):
        cell = ws.cell(row=row, column=col)
        cell.fill = fill
        cell.font = Font(name="Aptos", size=10, bold=True, color="FFFFFF")
        cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.cell(row=row, column=start).value = title


def build_workbook(companies, output_path, model_path):
    """Preenche o modelo de uma aba com os três blocos lado a lado."""
    if not model_path.is_file():
        raise FileNotFoundError(f"Modelo de planilha não encontrado: {model_path}")
    wb = load_workbook(model_path)
    ws = wb["Folha1"]
    max_col = ws.max_column
    prototype = {col: ws.cell(3, col)._style for col in range(1, max_col + 1)}
    # Complete o estilo das colunas extras da Lemit usando o par final do modelo.
    for col in range(61, max_col + 1):
        prototype[col] = ws.cell(3, 60 if col % 2 else 59)._style
    for row in ws.iter_rows(min_row=3):
        for cell in row:
            cell.value = None

    thin_gray = Side(style="thin", color="D9E1F2")
    for row_index, (cnpj, company) in enumerate(companies.items(), start=3):
        for col in range(1, max_col + 1):
            if prototype.get(col):
                ws.cell(row_index, col)._style = prototype[col]
        ws.cell(row_index, 1, cnpj).number_format = "@"
        ws.cell(row_index, 2, company["name"] or "")

        # Assertiva ocupa C:AT: telefone e fonte em colunas alternadas.
        phones = company["phones"]["ASSERTIVA"]
        for index in range(22):
            phone = phones[index] if index < len(phones) else None
            phone_col = 3 + index * 2
            if phone:
                ws.cell(row_index, phone_col, phone).number_format = "@"
                ws.cell(row_index, phone_col + 1, origin_labels(company, phone))

        # Consulta Dados: telefone da empresa e telefone do sócio.
        for index in range(2):
            consulta_phones = company["phones"]["CONSULTA DADOS"]
            phone = consulta_phones[index] if index < len(consulta_phones) else None
            phone_col = 47 + index * 2  # AU e AW
            if phone:
                ws.cell(row_index, phone_col, format_consulta_phone(phone)).number_format = "@"
                ws.cell(row_index, phone_col + 1, origin_labels(company, phone))

        # O arquivo da Lemit repete DDD/FONE; o modelo reserva cinco espaços.
        lemit_phones = company["phones"]["LEMIT"]
        for index in range(5):
            phone = lemit_phones[index] if index < len(lemit_phones) else None
            phone_col = 51 + index * 2  # AY em diante
            if phone:
                ws.cell(row_index, phone_col, phone).number_format = "@"
                ws.cell(row_index, phone_col + 1, origin_labels(company, phone))

        for cell in ws[row_index][:max_col]:
            cell.border = Border(bottom=thin_gray)

    # Abra na posição inicial do modelo. A seleção original estava em AZ45;
    # combinada ao congelamento de painéis, fazia o Excel abrir deslocado à direita.
    ws.auto_filter.ref = None
    ws.freeze_panes = None
    if ws.sheet_view.selection:
        ws.sheet_view.selection = [ws.sheet_view.selection[0]]
        ws.sheet_view.selection[0].activeCell = "A1"
        ws.sheet_view.selection[0].sqref = "A1"
    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)

def select_files():
    import tkinter as tk
    from tkinter import filedialog

    root = tk.Tk()
    root.withdraw()
    choices = (
        ("assertiva", "Selecione o CSV da Assertiva", [("CSV", "*.csv")]),
        ("consulta", "Selecione o Excel da Consulta Dados", [("Excel", "*.xlsx *.xlsm")]),
        ("lemit", "Selecione o CSV da Lemit", [("CSV", "*.csv")]),
    )
    paths = {}
    for key, title, filetypes in choices:
        chosen = filedialog.askopenfilename(title=title, filetypes=filetypes)
        if not chosen:
            root.destroy()
            return None
        paths[key] = Path(chosen)
    default = paths["assertiva"].parent / f"Lista consolidada {datetime.now():%Y%m%d_%H%M}.xlsx"
    chosen_output = filedialog.asksaveasfilename(
        title="Salvar planilha consolidada",
        initialdir=str(default.parent),
        initialfile=default.name,
        defaultextension=".xlsx",
        filetypes=[("Excel", "*.xlsx")],
    )
    root.destroy()
    if not chosen_output:
        return None
    return paths, Path(chosen_output)


def parse_args():
    parser = argparse.ArgumentParser(description="Consolida telefones e origens de três plataformas.")
    parser.add_argument("--assertiva", type=Path)
    parser.add_argument("--consulta", type=Path)
    parser.add_argument("--lemit", type=Path)
    parser.add_argument("--saida", type=Path)
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    gui = not any((args.assertiva, args.consulta, args.lemit, args.saida))
    if gui:
        selected = select_files()
        if selected is None:
            return 0
        input_paths, output_path = selected
    else:
        if not (args.assertiva and args.consulta and args.lemit and args.saida):
            raise SystemExit("Informe os quatro caminhos: --assertiva, --consulta, --lemit e --saida; ou execute sem parâmetros para abrir as janelas de seleção.")
        input_paths = {"assertiva": args.assertiva, "consulta": args.consulta, "lemit": args.lemit}
        output_path = args.saida

    for key, path in input_paths.items():
        if not path.is_file():
            raise FileNotFoundError(f"Arquivo não encontrado para {key}: {path}")
    if output_path.resolve() in {path.resolve() for path in input_paths.values()}:
        raise ValueError("Escolha um caminho de saída diferente dos arquivos de entrada.")

    companies: OrderedDict[str, dict] = OrderedDict()
    assertiva_fields, assertiva_keys, assertiva_invalid = load_assertiva(input_paths["assertiva"], companies)
    consulta_fields, consulta_keys, consulta_invalid, consulta_sheet = load_consulta(input_paths["consulta"], companies)
    lemit_fields, lemit_keys, lemit_invalid = load_lemit(input_paths["lemit"], companies)
    platform_fields = {"ASSERTIVA": assertiva_fields, "CONSULTA DADOS": consulta_fields, "LEMIT": lemit_fields}
    source_cnpjs = {"ASSERTIVA": assertiva_keys, "CONSULTA DADOS": consulta_keys, "LEMIT": lemit_keys}
    invalid_rows = {"ASSERTIVA": assertiva_invalid, "CONSULTA DADOS": consulta_invalid, "LEMIT": lemit_invalid}
    model_path = Path(__file__).with_name("Modelo consolidacao.xlsx")
    build_workbook(companies, output_path, model_path)
    message = f"Planilha criada com {len(companies)} empresas. Os três blocos estão na mesma aba e cada telefone traz sua origem.\n\n{output_path}"
    if gui:
        import tkinter as tk
        from tkinter import messagebox

        root = tk.Tk()
        root.withdraw()
        messagebox.showinfo("Consolidação concluída", message)
        root.destroy()
    else:
        print(message)
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as exc:
        print(f"Erro: {exc}", file=sys.stderr)
        raise


