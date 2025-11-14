import io
import streamlit as st
import pandas as pd


# ---------------------
# Inicialização
# ---------------------
def init_session():
    if "tabela_base" not in st.session_state:
        st.session_state.tabela_base = None
    if "orcamento" not in st.session_state:
        st.session_state.orcamento = []
    if "base_dict" not in st.session_state:
        st.session_state.base_dict = {}
    if "co_ok" not in st.session_state:
        st.session_state.co_ok = False


# ---------------------
# Detectar automaticamente o cabeçalho
# ---------------------
def detectar_cabecalho(df):
    colunas_esperadas = ["Item", "un", "Quantidade Total", "Valor Total"]

    for i in range(10):  # tenta nas 10 primeiras linhas
        linha = df.iloc[i].astype(str).str.strip()
        if all(c in linha.values for c in colunas_esperadas):
            return i

    return None


# ---------------------
# Carregar planilha
# ---------------------
def carregar_tabela_base(upload):
    try:
        xls = pd.ExcelFile(upload)
        df_raw = pd.read_excel(xls, sheet_name=0, header=None)

        linha_header = detectar_cabecalho(df_raw)
        if linha_header is None:
            st.error("Não encontrei cabeçalho com as colunas: Item, un, Quantidade Total, Valor Total")
            return None

        df = pd.read_excel(xls, sheet_name=0, header=linha_header)

        # Filtra apenas itens válidos
        df = df[["Item", "un", "Quantidade Total", "Valor Total"]].dropna()

        df["Quantidade Total"] = df["Quantidade Total"].replace(0, pd.NA)
        df["valor_unitario"] = df["Valor Total"] / df["Quantidade Total"]
        df = df.dropna(subset=["valor_unitario"])

        return df

    except Exception as e:
        st.error(f"Erro ao carregar planilha: {e}")
        return None


# ---------------------
# Criar dicionário de preços
# ---------------------
def montar_base_dict(df):
    base = {}
    for _, r in df.iterrows():
        base[r["Item"]] = {
            "un": r["un"],
            "valor_unitario": float(r["valor_unitario"])
        }
    return base


# ---------------------
# Aplicativo
# ---------------------
def main():
    st.set_page_config(page_title="Gerador de CO")
    init_session()

    st.title("Gerador de CO — Base AW/Spark")
    st.write("Carregue sua planilha AW e gere um novo orçamento.")

    # Upload
    upload = st.file_uploader("Carregar planilha base (.xlsx)", type=["xlsx"])

    if upload:
        df = carregar_tabela_base(upload)
        if df is not None:
            st.session_state.tabela_base = df
            st.session_state.base_dict = montar_base_dict(df)
            st.success("Base carregada com sucesso!")
            st.dataframe(df)

    st.markdown("---")

    # Botão gerar CO
    if st.button("GERAR CO"):
        if st.session_state.tabela_base is None:
            st.error("Carregue a tabela base primeiro.")
        else:
            st.session_state.orcamento = []
            st.session_state.co_ok = True

    # Se CO ativo:
    if st.session_state.co_ok and st.session_state.tabela_base is not None:

        st.subheader("Montagem do Orçamento")

        if st.button("Adicionar Linha"):
            st.session_state.orcamento.append(
                {"item": None, "un": "", "vu": 0.0, "qt": 0.0, "total": 0.0}
            )

        base = st.session_state.base_dict

        novas_linhas = []
        total_geral = 0

        for idx, linha in enumerate(st.session_state.orcamento):
            st.markdown(f"### Linha {idx+1}")

            col1, col2, col3, col4, col5 = st.columns([3, 1, 2, 2, 2])

            with col1:
                item = st.selectbox(
                    "Item:",
                    options=[""] + list(base.keys()),
                    index=0,
                    key=f"item_{idx}"
                )

            if item:
                un = base[item]["un"]
                vu = base[item]["valor_unitario"]
            else:
                un = ""
                vu = 0.0

            with col2:
                st.write("un")
                st.text_input("", value=un, disabled=True, key=f"un_{idx}")

            with col3:
                st.write("V. Unitário")
                st.text_input("", value=f"{vu:.2f}", disabled=True, key=f"vu_{idx}")

            with col4:
                qt = st.number_input(
                    "Qtd",
                    min_value=0.0,
                    step=1.0,
                    key=f"qt_{idx}"
                )

            total = qt * vu
            total_geral += total

            with col5:
                st.write("Total")
                st.text_input("", value=f"{total:.2f}", disabled=True, key=f"total_{idx}")

            novas_linhas.append({
                "item": item,
                "un": un,
                "vu": vu,
                "qt": qt,
                "total": total
            })

            st.markdown("---")

        st.session_state.orcamento = novas_linhas

        st.subheader(f"Total Geral: R$ {total_geral:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))

        # Download
        df_final = pd.DataFrame(novas_linhas)
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine="openpyxl") as w:
            df_final.to_excel(w, index=False)
        buffer.seek(0)

        st.download_button(
            "Baixar Orçamento",
            data=buffer,
            file_name="orcamento.xlsx"
        )


if __name__ == "__main__":
    main()
