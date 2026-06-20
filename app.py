import streamlit as st
from datetime import datetime, date

import ml_core
import ml_api
import shopee_core

REDIRECT_URI = "https://fechamento-ke6rzovxkuvjudzaug6pyu.streamlit.app/"

st.set_page_config(
    page_title="Consolidador de Fechamento",
    page_icon="📊",
    layout="centered",
)

@st.cache_resource
def _get_supabase():
    try:
        from supabase import create_client
        url = st.secrets["supabase"]["url"]
        key = st.secrets["supabase"]["key"]
        return create_client(url, key)
    except Exception:
        return None


def _db_save_rt(account: str, refresh_token: str):
    sb = _get_supabase()
    if not sb or not refresh_token:
        return
    try:
        sb.table("ml_connections").upsert(
            {"account": account, "refresh_token": refresh_token}
        ).execute()
    except Exception:
        pass


def _db_load_rt(account: str) -> str:
    sb = _get_supabase()
    if not sb:
        return ""
    try:
        result = (
            sb.table("ml_connections")
            .select("refresh_token")
            .eq("account", account)
            .single()
            .execute()
        )
        return (result.data or {}).get("refresh_token", "")
    except Exception:
        return ""


def _handle_oauth_callback():
    code  = st.query_params.get("code")
    state = st.query_params.get("state", "")

    if not code:
        return

    parts    = state.split("|", 1)
    account  = parts[0]
    verifier = parts[1] if len(parts) > 1 else ""

    if account == "ricapet" and "ml_token_ricapet" in st.session_state:
        st.query_params.clear()
        return
    if account == "thapets" and "ml_token_thapets" in st.session_state:
        st.query_params.clear()
        return

    try:
        if account == "ricapet":
            cfg    = st.secrets["ml_ricapet"]
            tokens = ml_api.exchange_code(cfg["client_id"], cfg["client_secret"],
                                          code, REDIRECT_URI, verifier)
            st.session_state["ml_token_ricapet"]  = tokens
            st.session_state["ml_userid_ricapet"] = ml_api.get_user_id(tokens["access_token"])
            _db_save_rt("ricapet", tokens.get("refresh_token", ""))
        elif account == "thapets":
            cfg    = st.secrets["ml_thapets"]
            tokens = ml_api.exchange_code(cfg["client_id"], cfg["client_secret"],
                                          code, REDIRECT_URI, verifier)
            st.session_state["ml_token_thapets"]  = tokens
            st.session_state["ml_userid_thapets"] = ml_api.get_user_id(tokens["access_token"])
            _db_save_rt("thapets", tokens.get("refresh_token", ""))
    except Exception as e:
        st.session_state["ml_auth_error"] = str(e)

    st.query_params.clear()

_handle_oauth_callback()


def _auto_authenticate():
    for account, secret_key in [("ricapet", "ml_ricapet"), ("thapets", "ml_thapets")]:
        if f"ml_token_{account}" in st.session_state:
            continue
        try:
            cfg = st.secrets[secret_key]
            rt  = _db_load_rt(account) or cfg.get("refresh_token", "")
            if not rt:
                continue
            tokens = ml_api.refresh_access_token(cfg["client_id"], cfg["client_secret"], rt)
            st.session_state[f"ml_token_{account}"]  = tokens
            st.session_state[f"ml_userid_{account}"] = ml_api.get_user_id(tokens["access_token"])
            new_rt = tokens.get("refresh_token", "")
            if new_rt and new_rt != rt:
                _db_save_rt(account, new_rt)
        except Exception:
            pass

_auto_authenticate()

st.title("📊 Consolidador de Fechamento")
st.caption("Mercado Livre e Shopee — Ricapet & Thapets")

tab_ml, tab_sp = st.tabs(["🛒 Mercado Livre", "🏪 Shopee"])

with tab_ml:
    st.subheader("Consolidar Mercado Livre")

    modo_ml = st.radio(
        "Modo",
        ["🔗 Buscar direto da plataforma (API)", "📂 Enviar arquivo manualmente"],
        horizontal=True,
        label_visibility="collapsed",
    )

    if modo_ml == "🔗 Buscar direto da plataforma (API)":

        if st.session_state.get("ml_auth_error"):
            st.error(f"Erro na autenticação: {st.session_state.pop('ml_auth_error')}")

        col1, col2 = st.columns(2)

        with col1:
            st.markdown("**Ricapet**")
            if "ml_token_ricapet" in st.session_state:
                st.success("✅ Conectado")
                if st.button("Desconectar", key="disc_ricapet"):
                    del st.session_state["ml_token_ricapet"]
                    del st.session_state["ml_userid_ricapet"]
                    st.rerun()
            else:
                try:
                    cfg = st.secrets["ml_ricapet"]
                    verifier, challenge = ml_api.generate_pkce()
                    state_r = f"ricapet|{verifier}"
                    url = ml_api.get_auth_url(cfg["client_id"], REDIRECT_URI,
                                              state=state_r, code_challenge=challenge)
                    if st.button("🔗 Conectar conta Ricapet", key="btn_conn_ricapet", type="primary"):
                        st.session_state["_oauth_url"] = url
                        st.rerun()
                except (KeyError, FileNotFoundError):
                    st.warning("Credenciais ml_ricapet não configuradas nos Secrets.")

        with col2:
            st.markdown("**Thapets**")
            if "ml_token_thapets" in st.session_state:
                st.success("✅ Conectado")
                if st.button("Desconectar", key="disc_thapets"):
                    del st.session_state["ml_token_thapets"]
                    del st.session_state["ml_userid_thapets"]
                    st.rerun()
            else:
                try:
                    cfg = st.secrets["ml_thapets"]
                    verifier, challenge = ml_api.generate_pkce()
                    state_t = f"thapets|{verifier}"
                    url = ml_api.get_auth_url(cfg["client_id"], REDIRECT_URI,
                                              state=state_t, code_challenge=challenge)
                    if st.button("🔗 Conectar conta Thapets", key="btn_conn_thapets", type="primary"):
                        st.session_state["_oauth_url"] = url
                        st.rerun()
                except (KeyError, FileNotFoundError):
                    st.warning("Credenciais ml_thapets não configuradas nos Secrets.")

        if "_oauth_url" in st.session_state:
            oauth_url = st.session_state.pop("_oauth_url")
            st.markdown(
                f'<meta http-equiv="refresh" content="0;url={oauth_url}">',
                unsafe_allow_html=True,
            )
            st.stop()

        tem_ricapet = "ml_token_ricapet" in st.session_state
        tem_thapets = "ml_token_thapets" in st.session_state

        if tem_ricapet or tem_thapets:
            st.divider()

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

                            if tem_ricapet:
                                tok = st.session_state["ml_token_ricapet"]
                                uid = st.session_state["ml_userid_ricapet"]
                                st.info(f"🔍 Ricapet user_id: {uid}")
                                ords = ml_api.fetch_orders(tok["access_token"], uid, from_str, to_str)
                                st.info(f"Ricapet: {len(ords)} pedido(s) encontrado(s)")
                                if ords:
                                    arquivos.append(("relatorio_736787693.xlsx",
                                                     ml_api.orders_to_excel_bytes(ords, "Ricapet")))
                                    total += len(ords)

                            if tem_thapets:
                                tok = st.session_state["ml_token_thapets"]
                                uid = st.session_state["ml_userid_thapets"]
                                st.info(f"🔍 Thapets user_id: {uid}")
                                ords = ml_api.fetch_orders(tok["access_token"], uid, from_str, to_str)
                                st.info(f"Thapets: {len(ords)} pedido(s) encontrado(s)")
                                if ords:
                                    arquivos.append(("relatorio_1139210125.xlsx",
                                                     ml_api.orders_to_excel_bytes(ords, "Thapets")))
                                    total += len(ords)

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
                        st.success(f"Consolidado! {len(arquivos)} arquivo(s) processado(s).")
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
