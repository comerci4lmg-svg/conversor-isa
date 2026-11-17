import hashlib
from pathlib import Path
import streamlit as st
from converter import (parse_xml, record_groups, path_label, records_to_dataframe,
                       dataframe_to_xlsx, normalize_output_name, build_zip, spreadsheetml_tables)

st.set_page_config(page_title="XML para Excel", page_icon="📊", layout="wide")
st.title("Conversor XML para Excel")
st.write("Envie seus XMLs, confira os registros e baixe os arquivos XLSX.")
st.caption("Códigos, datas, zeros à esquerda e TRUE/FALSE são preservados como texto.")
uploads = st.file_uploader("Selecione os arquivos XML", type=["xml"], accept_multiple_files=True)
if not uploads:
    st.info("Selecione um ou mais arquivos para começar.")
    st.stop()

prepared = []
invalid = False
for i, upload in enumerate(uploads):
    try:
        content = upload.getvalue()
        token = hashlib.sha256(content).hexdigest()
        root = parse_xml(content)
        tables = spreadsheetml_tables(root)
        if tables is not None:
            with st.expander(upload.name, expanded=True):
                st.success("XML de Excel reconhecido automaticamente.")
                st.caption("A primeira linha é usada como cabeçalho. Valores são preservados como texto; estilos e fórmulas originais não são transferidos. Cada aba gera um XLSX separado.")
                for j, (sheet_name, frame) in enumerate(tables):
                    st.write(f"{sheet_name}: {len(frame)} registros · {len(frame.columns)} colunas")
                    st.dataframe(frame.head(10), hide_index=True, use_container_width=True)
                    suffix = f"_{sheet_name}" if len(tables) > 1 else ""
                    name = st.text_input("Nome do XLSX", Path(upload.name).stem + suffix + ".xlsx", key=f"name_{i}_{j}_{token}")
                    prepared.append((normalize_output_name(name), frame))
            continue
        groups = record_groups(root)
        paths = list(groups)
        # Prefer repeated parent elements, then repeated leaves, then root.
        suggested = max(paths, key=lambda path: (len(groups[path]) > 1 and any(len(n) for n in groups[path]),
                                                len(groups[path]) > 1, len(groups[path]), -len(path)))
        with st.expander(upload.name, expanded=True):
            selected = st.selectbox("Grupo que representa cada linha", paths,
                                    index=paths.index(suggested),
                                    format_func=lambda path: f"{path_label(path)} ({len(groups[path])} registros)",
                                    key=f"group_{i}_{token}")
            st.caption("Cada elemento desse grupo vira uma linha. Campos internos viram colunas; repetições recebem [1], [2] etc. Atributos usam @. Dados fora do grupo selecionado não são incluídos.")
            frame = records_to_dataframe(groups[selected])
            st.write(f"{len(frame)} registros · {len(frame.columns)} colunas")
            st.dataframe(frame.head(10), hide_index=True, use_container_width=True)
            name = st.text_input("Nome do XLSX", Path(upload.name).stem + ".xlsx", key=f"name_{i}_{token}")
            prepared.append((normalize_output_name(name), frame))
    except Exception as exc:
        invalid = True
        st.error(f"Não foi possível ler {upload.name}: {exc}")

# Tie downloads to both data and selection; changing inputs removes stale results.
signature = hashlib.sha256()
for name, frame in prepared:
    signature.update(name.encode())
    signature.update(frame.to_json(force_ascii=False).encode())
key = signature.hexdigest()
if invalid:
    st.warning("Remova ou corrija os arquivos com erro antes de converter.")
    st.stop()
if st.button("Converter para XLSX", type="primary", use_container_width=True):
    try:
        with st.spinner("Convertendo..."):
            generated = [(name, dataframe_to_xlsx(frame)) for name, frame in prepared]
        st.session_state["result"] = (key, generated)
        st.success("Conversão concluída!")
    except Exception as exc:
        st.session_state.pop("result", None)
        st.error(f"Não foi possível converter: {exc}")
result = st.session_state.get("result")
if result and result[0] == key:
    for i, (name, data) in enumerate(result[1]):
        st.download_button(f"Baixar {name}", data, name,
                           "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key=f"download_{i}")
    if len(result[1]) > 1:
        st.download_button("Baixar todas em ZIP", build_zip(result[1]), "planilhas_convertidas.zip", "application/zip")
