from __future__ import annotations
import io
import re
import zipfile
from collections import Counter
from collections.abc import Iterable
import xml.etree.ElementTree as ET
from xml.parsers import expat
import pandas as pd


def parse_xml(content: bytes):
    if not content.strip():
        raise ValueError("O arquivo está vazio.")
    if len(content) > 50 * 1024 * 1024:
        raise ValueError("O limite é de 50 MB por XML.")
    # Block DTDs and entities, including in UTF-16 documents.
    parser = expat.ParserCreate()
    def reject(*args):
        raise ValueError("XML com DTD ou entidades não é aceito.")
    parser.StartDoctypeDeclHandler = reject
    parser.EntityDeclHandler = reject
    parser.ExternalEntityRefHandler = reject
    try:
        parser.Parse(content, True)
        return ET.fromstring(content)
    except (expat.ExpatError, ET.ParseError) as exc:
        raise ValueError(f"XML inválido: {exc}") from exc


def record_groups(root):
    groups = {}
    def walk(node, path):
        groups.setdefault(path, []).append(node)
        for child in node:
            walk(child, path + (child.tag,))
    walk(root, (root.tag,))
    return groups


def path_label(path):
    # Namespace URIs remain visible to distinguish equally named tags.
    return " / ".join(path)


def flatten_record(node):
    row = {}
    def walk(element, path):
        for name, value in element.attrib.items():
            row[f"{path}/@{name}"] = value
        text = (element.text or "").strip()
        if text or not len(element):
            row[path] = text
        totals = Counter(child.tag for child in element)
        seen = Counter()
        for child in element:
            seen[child.tag] += 1
            suffix = f"[{seen[child.tag]}]" if totals[child.tag] > 1 else ""
            child_path = f"{path}/{child.tag}{suffix}"
            walk(child, child_path)
            if child.tail and child.tail.strip():
                row[child_path + "/#tail"] = child.tail.strip()
    walk(node, node.tag)
    return row


def records_to_dataframe(nodes):
    rows = [flatten_record(node) for node in nodes]
    frame = pd.DataFrame(rows).fillna("")
    if not len(frame.columns):
        raise ValueError("O grupo selecionado não contém dados.")
    if len(frame) > 1048575 or len(frame.columns) > 16384:
        raise ValueError("Os dados ultrapassam o limite de linhas ou colunas do Excel. Selecione outro grupo.")
    if any(len(str(value)) > 32767 for row in rows for value in row.values()):
        raise ValueError("Há um campo maior que o limite de 32.767 caracteres por célula do Excel.")
    if any(len(str(column)) > 32767 for column in frame.columns):
        raise ValueError("Há um nome de coluna maior que o limite do Excel.")
    return frame


def _safe_filename(value: str) -> str:
    value = value.strip().replace(" ", "_")
    value = re.sub(r"[<>:\"/\\|?*\x00-\x1f]", "_", value)
    value = re.sub(r"_+", "_", value).strip("._")
    return value[:150] or "planilha_convertida"


def normalize_output_name(value: str) -> str:
    value = re.sub(r"\.xlsx$", "", value.strip(), flags=re.IGNORECASE)
    return f"{_safe_filename(value)}.xlsx"


def dataframe_to_xlsx(dataframe: pd.DataFrame) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(
        output,
        engine="xlsxwriter",
        engine_kwargs={
            "options": {
                "strings_to_formulas": False,
                "strings_to_urls": False,
                "strings_to_numbers": False,
            }
        },
    ) as writer:
        # dtype=str na leitura e strings_to_* desativados garantem que TRUE/FALSE
        # continuem texto normal no Excel, sem checkbox ou conversão booleana.
        dataframe.to_excel(writer, sheet_name="Dados", index=False, header=False, startrow=1)
        workbook = writer.book
        worksheet = writer.sheets["Dados"]

        header_format = workbook.add_format(
            {
                "bold": True,
                "font_color": "#FFFFFF",
                "bg_color": "#1F4E78",
                "align": "center",
                "valign": "vcenter",
                "border": 1,
                "border_color": "#D6E4F0",
                "text_wrap": True,
            }
        )
        text_format = workbook.add_format({"num_format": "@", "valign": "vcenter"})
        worksheet.write_row(0, 0, list(dataframe.columns), header_format)
        worksheet.set_row(0, 34)
        worksheet.freeze_panes(1, 0)
        worksheet.hide_gridlines(2)

        row_count = len(dataframe)
        column_count = len(dataframe.columns)
        if column_count:
            columns = [{"header": str(column)} for column in dataframe.columns]
            worksheet.add_table(
                0,
                0,
                max(row_count, 1),
                column_count - 1,
                {
                    "name": "TabelaDados",
                    "columns": columns,
                    "style": "Table Style Medium 2",
                },
            )

        sample = dataframe.head(1000)
        for index, column in enumerate(dataframe.columns):
            values = sample[column].astype("string") if column in sample else pd.Series(dtype="string")
            observed = max([len(str(column)), *(len(str(value)) for value in values)], default=len(str(column)))
            width = min(max(observed + 2, 12), 38)
            worksheet.set_column(index, index, width, text_format)

        if row_count == 0:
            worksheet.write(1, 0, "Sem registros")

    return output.getvalue()


def build_zip(files: Iterable[tuple[str, bytes]]) -> bytes:
    output = io.BytesIO()
    with zipfile.ZipFile(output, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        used_names: set[str] = set()
        for raw_name, content in files:
            name = normalize_output_name(raw_name)
            base = name[:-5]
            candidate = name
            counter = 2
            while candidate.casefold() in used_names:
                candidate = f"{base}_{counter}.xlsx"
                counter += 1
            used_names.add(candidate.casefold())
            archive.writestr(candidate, content)
    return output.getvalue()


SS = "{urn:schemas-microsoft-com:office:spreadsheet}"


def spreadsheetml_tables(root):
    """Return named tables, or None for a generic XML document."""
    if root.tag != SS + "Workbook":
        return None
    result = []
    for sheet in root.findall(SS + "Worksheet"):
        table = sheet.find(SS + "Table")
        if table is None:
            continue
        rows = []
        for row in table.findall(SS + "Row"):
            row_index = int(row.get(SS + "Index", len(rows) + 1))
            if not len(rows) < row_index <= 1048576:
                raise ValueError("Índice de linha inválido ou acima do limite do Excel.")
            rows.extend([[] for _ in range(row_index - len(rows) - 1)])
            values = []
            for cell in row.findall(SS + "Cell"):
                index = int(cell.get(SS + "Index", len(values) + 1))
                span = int(cell.get(SS + "MergeAcross", "0"))
                if not len(values) < index <= 16384 or span < 0 or index + span > 16384:
                    raise ValueError("Índice de coluna inválido ou acima do limite do Excel.")
                values.extend([""] * (index - len(values) - 1))
                data = cell.find(SS + "Data")
                value = "" if data is None else "".join(data.itertext())
                if len(value) > 32767:
                    raise ValueError("Campo acima do limite de 32.767 caracteres do Excel.")
                values.append(value)
                values.extend([""] * span)
            rows.append(values)
        if not rows or not any(rows):
            continue
        width = max(map(len, rows))
        rows = [r + [""] * (width - len(r)) for r in rows]
        headers = []
        used = set()
        for i, raw in enumerate(rows[0], 1):
            base = raw or f"Coluna_{i}"
            name = base
            n = 2
            while name.casefold() in used:
                name = f"{base}_{n}"
                n += 1
            used.add(name.casefold())
            headers.append(name)
        result.append((sheet.get(SS + "Name", "Dados"), pd.DataFrame(rows[1:], columns=headers)))
    if not result:
        raise ValueError("O XML de Excel não contém planilhas com dados.")
    return result
