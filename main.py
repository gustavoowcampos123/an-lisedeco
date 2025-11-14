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
    try:
        df = pd.read_excel(uploaded_file)

        colunas_esperadas = ["Item", "un", "Quantidade Total", "Valor Total"]
        for col in colunas_esperadas:
            if col not in df.columns:
                st.error(f"Coluna obrigatória ausente na planilha: **{col}**")
                return None

        # Evita divisão por zero
        df["Quantidade Total"] = df["Quantidade Total"].replace(0, pd.NA)
        df["valor_unitario"] = df["Valor Total"] / df["Quantidade Total"]

        # Remove linhas sem valor_unitario válido
        df = df.dropna(subset=["valor_unitario"])

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
        help="A planilha deve conter as colunas: Item, un, Quantidade Total, Valor Total.",
    )

    if uploaded_file is not None:
        df_base = carregar_tabela_base(uploaded_file)
        if df_base is not None:
            st.session_state["tabela_base"] = df_base
            st.session_state["base_dict"] = montar_base_dict(df_base)
            st.sidebar.success("Tabela base carregada com sucesso!")
    else:
        st.sidebar.info("Nenhuma planilha carregada ainda.")

    # Preview da tabela base
    if st.session_state["tabela_base"] is not None:
        with st.expander("Ver tabela base carregada"):
            st.dataframe(st.session_state["tabela_base"])

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

            # Select de item
            with col_item:
                item = st.selectbox(
                    "Item",
                    options=[""] + itens_disponiveis,
                    index=itens_disponiveis.index(linha["item"]) + 1
                    if linha["item"] in itens_disponiveis
                    else 0,
                    key=f"item_{idx}",
                )
            # Atualiza dados de unidade e valor unitário
            if item and item in base_dict:
                un = base_dict[item]["un"]
                valor_unitario = base_dict[item]["valor_unitario"]
            else:
                un = ""
                valor_unitario = 0.0

            # Unidade (somente leitura)
            with col_un:
                st.text_input("un", value=un, disabled=True, key=f"un_{idx}")

            # Valor unitário (somente leitura)
            with col_vu:
                st.text_input(
                    "Valor Unitário",
                    value=f"{valor_unitario:.2f}",
                    disabled=True,
                    key=f"vu_{idx}",
                )

            # Quantidade
            with col_qt:
                quantidade = st.number_input(
                    "Quantidade",
                    min_value=0.0,
                    step=1.0,
                    value=float(linha["quantidade"]) if linha["quantidade"] else 0.0,
                    key=f"qt_{idx}",
                )

            # Total da linha
            total_linha = valor_unitario * quantidade
            total_geral += total_linha

            with col_total:
                st.text_input(
                    "Total da Linha",
                    value=f"{total_linha:.2f}",
                    disabled=True,
                    key=f"total_{idx}",
                )

            # Guarda linha atualizada
            linhas_atualizadas.append(
                {
                    "item": item if item else None,
                    "un": un,
                    "valor_unitario": valor_unitario,
                    "quantidade": quantidade,
                    "total_linha": total_linha,
                }
            )

            st.markdown("---")

        # Atualiza no session_state
        st.session_state["orcamento_linhas"] = linhas_atualizadas

        # Resumo do orçamento
        st.subheader("Resumo do Orçamento")
        st.markdown(f"**Total Geral do Orçamento: R$ {total_geral:,.2f}**".replace(",", "X").replace(".", ",").replace("X", "."))

        df_orcamento = gerar_dataframe_orcamento()
        if not df_orcamento.empty:
            st.dataframe(df_orcamento)

            excel_bytes = gerar_excel_download(df_orcamento)
            st.download_button(
                label="Baixar Orçamento em Excel",
                data=excel_bytes,
                file_name="orcamento_CO.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

    else:
        st.info("Carregue a tabela base e clique em **GERAR CO** para iniciar o orçamento.")


if __name__ == "__main__":
    main()
