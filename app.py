import streamlit as st
from datetime import datetime, date

import ml_core
import ml_api
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

    modo_ml = st.radio(
        "Modo",
        ["🔗 Buscar direto da plataforma (API)", "📂 Enviar arquivo manualmente"],
        horizontal=True,
        label_visibility="collapsed",
    )

    # ------------------------------------------------------------------
    # MODO API
    # ------------------------------------------------------------------
    if modo_ml == "🔗 Buscar direto da plataforma (API)":

        hoje = date.today()
        col_d1, col_d2 = st.columns(2)
        with col_d1:
            data_ini = st.date_input("De", value=hoje.replace(day=1), key="ml_api_ini")
        with col_d2:
            data_fim = st.date_input("Até", value=hoje, key="ml_api_fim")

        tabela_api = st.file_uploader("TABELA_AUXILIAR.xlsx", type="xlsx", key="ml_api_aux")

        btn_api = st.button("🔍 Buscar e Consolidar", type="primary", key="btn_ml_api")

        if btn_api:
            if not tabela_api:
                st.error("Selecione a TABELA_AUXILIAR.xlsx.")
            elif data_ini > data_fim:
                st.error("A data inicial deve ser anterior à data final.")
            else:
                with st.spinner("Buscando pedidos no Mercado Livre..."):
                    try:
                        from_str = f"{data_ini}T00:00:00.000-03:00"
                        to_str   = f"{data_fim}T23:59:59.000-03:00"
                        arquivos = []
                        total    = 0

                        try:
                            cfg = st.secrets["ml_ricapet"]
                            tok, uid = ml_api.get_app_token(cfg["client_id"], cfg["client_secret"])
                            ords = ml_api.fetch_orders(tok, uid, from_str, to_str)
                            st.info(f"Ricapet: {len(ords)} pedido(s) encontrado(s)")
                            if ords:
                                arquivos.append(("relatorio_736787693.xlsx",
                                                 ml_api.orders_to_excel_bytes(ords, "Ricapet")))
                                total += len(ords)
                        except (KeyError, FileNotFoundError):
                            st.warning("Credenciais ml_ricapet não configuradas nos Secrets.")

                        try:
                            cfg = st.secrets["ml_thapets"]
                            tok, uid = ml_api.get_app_token(cfg["client_id"], cfg["client_secret"])
                            ords = ml_api.fetch_orders(tok, uid, from_str, to_str)
                            st.info(f"Thapets: {len(ords)} pedido(s) encontrado(s)")
                            if ords:
                                arquivos.append(("relatorio_1139210125.xlsx",
                                                 ml_api.orders_to_excel_bytes(ords, "Thapets")))
                                total += len(ords)
                        except (KeyError, FileNotFoundError):
                            st.warning("Credenciais ml_thapets não configuradas nos Secrets.")

                        if not arquivos:
                            st.warning("Nenhum pedido encontrado no período selecionado.")
                        else:
                            tabela_bytes = tabela_api.read()
                            xlsx_bytes, logs = ml_core.processar(arquivos, tabela_bytes)

                            ts = datetime.now().strftime("%Y%m%d_%H%M%S")
                            st.success(f"Consolidado! {total} pedido(s) processado(s).")
                            st.download_button(
                                label="⬇️  Baixar Excel gerado",
                                data=xlsx_bytes,
                                file_name=f"MercadoLivre_Consolidado_{ts}.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                key="dl_ml_api",
                            )
                            with st.expander("Log de processamento"):
                                for msg in logs:
                                    st.text(msg)

                    except Exception as exc:
                        st.error(f"Erro: {exc}")

    # ------------------------------------------------------------------
    # MODO ARQUIVO (fallback)
    # ------------------------------------------------------------------
    else:
        col1, col2 = st.columns(2)
        with col1:
            relatorios_ml = st.file_uploader(
                "Relatórios ML (.xlsx)",
                type="xlsx",
                accept_multiple_files=True,
                key="ml_rel",
                help="Arquivos baixados do Mercado Livre "
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
                        arquivos     = [(f.name, f.read()) for f in relatorios_ml]
                        tabela_bytes = tabela_ml.read()
                        xlsx_bytes, logs = ml_core.processar(arquivos, tabela_bytes)

                        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
                        st.success(f"Consolidado com sucesso! {len(arquivos)} arquivo(s) processado(s).")
                        st.download_button(
                            label="⬇️  Baixar Excel gerado",
                            data=xlsx_bytes,
                            file_name=f"MercadoLivre_Consolidado_{ts}.xlsx",
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
            help="Arquivos baixados da Shopee (nome deve começar com 'order.all.').",
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
                    arquivos     = [(f.name, f.read()) for f in relatorios_sp]
                    tabela_bytes = tabela_sp.read()
                    xlsx_bytes, logs = shopee_core.processar(arquivos, tabela_bytes)

                    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
                    st.success("Consolidado com sucesso!")
                    st.download_button(
                        label="⬇️  Baixar Excel gerado",
                        data=xlsx_bytes,
                        file_name=f"Shopee_Consolidado_{ts}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="dl_sp",
                    )
                    with st.expander("Log de processamento"):
                        for msg in logs:
                            st.text(msg)
                except Exception as exc:
                    st.error(f"Erro ao processar: {exc}")
