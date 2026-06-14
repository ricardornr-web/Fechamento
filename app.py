import streamlit as st
from datetime import datetime

import ml_core
import shopee_core

st.set_page_config(
    page_title="Consolidador de Fechamento",
    page_icon="📊",
    layout="centered",
)

st.title("📊 Consolidador de Fechamento")
st.caption("Mercado Livre e Shopee — Ricapet & Thapets")

tab_ml, tab_sp = st.tabs(["🛒 Mercado Livre", "🏪 Shopee"])

# =============================================================================
# ABA MERCADO LIVRE
# =============================================================================

with tab_ml:
    st.subheader("Consolidar Mercado Livre")

    col1, col2 = st.columns(2)
    with col1:
        relatorios_ml = st.file_uploader(
            "Relatórios ML (.xlsx)",
            type="xlsx",
            accept_multiple_files=True,
            key="ml_rel",
            help="Selecione um ou mais arquivos baixados do Mercado Livre "
                 "(devem conter 736787693 ou 1139210125 no nome).",
        )
    with col2:
        tabela_ml = st.file_uploader(
            "TABELA_AUXILIAR.xlsx",
            type="xlsx",
            key="ml_aux",
        )

    btn_ml = st.button("Consolidar Mercado Livre", type="primary", key="btn_ml")

    if btn_ml:
        if not relatorios_ml:
            st.error("Selecione pelo menos um relatório do Mercado Livre.")
        elif not tabela_ml:
            st.error("Selecione a TABELA_AUXILIAR.xlsx.")
        else:
            with st.spinner("Processando..."):
                try:
                    arquivos = [(f.name, f.read()) for f in relatorios_ml]
                    tabela_bytes = tabela_ml.read()
                    xlsx_bytes, logs = ml_core.processar(arquivos, tabela_bytes)

                    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
                    nome_arquivo = f"MercadoLivre_Consolidado_{ts}.xlsx"

                    st.success(
                        f"Consolidado com sucesso! "
                        f"{len(arquivos)} arquivo(s) processado(s)."
                    )
                    st.download_button(
                        label="⬇️  Baixar Excel gerado",
                        data=xlsx_bytes,
                        file_name=nome_arquivo,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="dl_ml",
                    )

                    with st.expander("Log de processamento"):
                        for msg in logs:
                            st.text(msg)

                except Exception as exc:
                    st.error(f"Erro ao processar: {exc}")

# =============================================================================
# ABA SHOPEE
# =============================================================================

with tab_sp:
    st.subheader("Consolidar Shopee")

    col1, col2 = st.columns(2)
    with col1:
        relatorios_sp = st.file_uploader(
            "Relatórios Shopee (.xlsx)",
            type="xlsx",
            accept_multiple_files=True,
            key="sp_rel",
            help="Selecione um ou mais arquivos baixados da Shopee "
                 "(o nome deve começar com 'order.all.').",
        )
    with col2:
        tabela_sp = st.file_uploader(
            "TABELA_AUXILIAR.xlsx",
            type="xlsx",
            key="sp_aux",
        )

    btn_sp = st.button("Consolidar Shopee", type="primary", key="btn_sp")

    if btn_sp:
        if not relatorios_sp:
            st.error("Selecione pelo menos um relatório da Shopee.")
        elif not tabela_sp:
            st.error("Selecione a TABELA_AUXILIAR.xlsx.")
        else:
            with st.spinner("Processando..."):
                try:
                    arquivos = [(f.name, f.read()) for f in relatorios_sp]
                    tabela_bytes = tabela_sp.read()
                    xlsx_bytes, logs = shopee_core.processar(arquivos, tabela_bytes)

                    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
                    nome_arquivo = f"Shopee_Consolidado_{ts}.xlsx"

                    st.success("Consolidado com sucesso!")
                    st.download_button(
                        label="⬇️  Baixar Excel gerado",
                        data=xlsx_bytes,
                        file_name=nome_arquivo,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="dl_sp",
                    )

                    with st.expander("Log de processamento"):
                        for msg in logs:
                            st.text(msg)

                except Exception as exc:
                    st.error(f"Erro ao processar: {exc}")
