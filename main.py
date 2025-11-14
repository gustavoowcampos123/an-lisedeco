import io
import streamlit as st
import pandas as pd


# --------- Funções auxiliares --------- #
def init_session_state():
    if "tabela_base" not in st.session_state:
        st.session_state["tabela_base"] = None
    if "base_dict" not in st.session_state:
        st.session_state["base_dict"] = {}
    if "orcamento_linhas" not in st.session_state:
        st.session_state["orcamento_linhas"] = []
    if "co_gerado" not in st.session_state:
        st.session_state["co_gerado"] = False


def carregar_tabela_base(uploaded_file):
    """
    Lê o Excel e tenta encontrar uma aba com as colunas:
    Item, un, Quantidade Total, Valor Total.

    A sua planilha tem esses cabeçalhos a partir da linha 6 (header=5),
    então o código tenta usar esse padrão em todas as abas.
    """
    try:
        xls = pd.ExcelFile(uploaded_file)
        colunas_esperadas = ["Item", "un", "Quantidade Total", "Valor Total"]

        df_encontrado = None
        aba_encontrada = None

        # tenta em todas as abas do arquivo
        for sheet in xls.sheet_names:
            try:
                # muitas planilhas "modelo AW" têm o cabeçalho na linha 6 -> header=5
                df_tmp = pd.read_excel(xls, sheet_name=sheet, header=5)
            except Exception:
                continue

            if all(col in df_tmp.columns for col in colunas_esperadas):
                df_encontrado = df_tmp
                aba_encontrada = sheet
                break

        if df_encontrado is None:
            st.error(
                "Não encontrei nenhuma aba com as colunas "
                "**Item**, **un**, **Quantidade Total** e **Valor Total**.\n\n"
                "Confirme se a planilha segue esse modelo."
            )
            return None

        df = df_encontrado.copy()

        # Mantém apenas as colunas importantes + demais (se quiser olhar depois)
        # Aqui garantimos que as colunas esperadas existem
        # e filtramos somente as linhas válidas
        df = df[
            (df["Item"].notna())
            & (df["Quantidade Total"].notna())
            & (df["Valor Total"].notna())
        ]

        # Evita divisão por zero
        df["Quantidade Total"] = df["Quantidade Total"].replace(0, pd.NA)

        # Calcula valor unitário a partir de Valor Total / Quantidade Total
        df["valor_unitario"] = df["Valor Total"] / df["Quantidade Total"]

        # Remove linhas sem valor_unitario válido
        df = df.dropna(subset=["valor_unitario"])

        # Feedback da aba utilizada
        st.sidebar.success(f"Tabela base carregada da aba: **{aba_encontrada}**")

        return df

    except Exception as e:
        st.error(f"Erro ao ler o arquivo Excel: {e}")
        return None


def montar_base_dict(df):
    base = {}
    for _, row in df.iterrows():
        item = row["Item"]
        base[item] = {
            "un": row["un"],
            "valor_unitario": float(row["valor_unitario"]),
        }
    return base


def inicializar_orcamento():
    st.session_state["orcamento_linhas"] = []
    st.session_state["co_gerado"] = True


def adicionar_linha():
    st.session_state["orcamento_linhas"].append(
        {
            "item": None,
            "un": "",
            "valor_unitario": 0.0,
            "quantidade": 0.0,
            "total_linha": 0.0,
        }
    )


def gerar_dataframe_orcamento():
    if not st.session_state["orcamento_linhas"]:
        return pd.DataFrame()

    df = pd.DataFrame(st.session_state["orcamento_linhas"])
    total_geral = df["total_linha"].sum()

    # Linha de total geral
    total_row = {
        "item": "TOTAL GERAL",
        "un": "",
        "valor_unitario": "",
        "quantidade": "",
        "total_linha": total_geral,
    }
    df_total = pd.concat([df, pd.DataFrame([total_row])], ignore_index=True)
    return df_total


def gerar_excel_download(df):
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Orcamento_CO")
    buffer.seek(0)
    return buffer


# --------- Interface Streamlit --------- #
def main():
    st.set_page_config(
        page_title="Montagem de Orçamento (CO)",
        layout="wide",
    )

    init_session_state()

    st.title("Montagem de Orçamento (CO)")
    st.markdown(
        "Aplicação para gerar orçamento a partir de uma planilha base em Excel "
        "com as colunas **Item**, **un**, **Quantidade Total** e **Valor Total**."
    )

    # --------- Sidebar: Tabela base --------- #
    st.sidebar.header("Configuração da Tabela Base")

    uploaded_file = st.sidebar.file_uploader(
        "Carregue a planilha base (.xlsx)",
        type=["xlsx"],
        help=(
            "A planilha deve conter as colunas: "
            "Item, un, Quantidade Total, Valor Total (como na Planilha Spark)."
        ),
    )

    if uploaded_file is not None:
        df_base = carregar_tabela_base(uploaded_file)
        if df_base is not None:
            st.session_state["tabela_base"] = df_base
            st.session_state["base_dict"] = montar_base_dict(df_base)
    else:
        st.sidebar.info("Nenhuma planilha carregada ainda.")

    # Preview da tabela base
    if st.session_state["tabela_base"] is not None:
        with st.expander("Ver tabela base carregada"):
            st.dataframe(st.session_state["tabela_base"][["Item", "un", "Quantidade Total", "Valor Total", "valor_unitario"]])

    st.markdown("---")

    # --------- Botão GERAR CO --------- #
    col1, col2 = st.columns([1, 3])
    with col1:
        gerar_co_btn = st.button("GERAR CO", type="primary")

    if gerar_co_btn:
        if st.session_state["tabela_base"] is None:
            st.error("Por favor, carregue a tabela base em Excel antes de gerar o CO.")
        else:
            inicializar_orcamento()

    # --------- Montagem do Orçamento --------- #
    if st.session_state["co_gerado"] and st.session_state["tabela_base"] is not None:
        st.subheader("Montagem do Orçamento (CO)")

        # Botão de adicionar linha
        if st.button("Adicionar linha"):
            adicionar_linha()

        base_dict = st.session_state["base_dict"]
        itens_disponiveis = list(base_dict.keys())

        linhas_atualizadas = []
        total_geral = 0.0

        # Renderizar cada linha do orçamento
        for idx, linha in enumerate(st.session_state["orcamento_linhas"]):
            st.markdown(f"**Linha {idx + 1}**")
            col_item, col_un, col_vu, col_qt, col_total = st.columns([3, 1, 2, 2, 2])

            # Sel
