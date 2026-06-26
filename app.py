import streamlit as st
from datetime import datetime, date

import ml_core
import ml_api
import shopee_core

REDIRECT_URI = "https://fechamento-ke6rzovxkuvjudzaug6pyu.streamlit.app/"

st.set_page_config(
    page_title="Fechamento · ML & Shopee",
    page_icon="🛒",
    layout="centered",
)

# =============================================================================
# CSS GLOBAL — esconde chrome do Streamlit e aplica design premium
# =============================================================================

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');

/* ── Esconder elementos padrão do Streamlit ── */
#MainMenu, footer, [data-testid="stToolbar"],
[data-testid="stDecoration"], [data-testid="stStatusWidget"] {
  display: none !important;
}
[data-testid="stHeader"] { background: transparent !important; height: 0 !important; }

/* ── Fonte global ── */
html, body, [class*="css"] {
  font-family: 'Inter', system-ui, sans-serif !important;
}

/* ── Padding do container principal ── */
.block-container {
  padding-top: 0 !important;
  padding-bottom: 48px !important;
  max-width: 820px !important;
}

/* ─────────────────────────────────────────────
   HEADER CUSTOMIZADO
───────────────────────────────────────────── */
.app-header {
  background: #1A1F71;
  margin: -1rem -1rem 0 -1rem;
  padding: 14px 32px;
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 12px;
}
.app-header-brand {
  display: flex;
  align-items: center;
  gap: 12px;
}
.app-header-logo {
  background: #FFE600;
  border-radius: 8px;
  width: 36px; height: 36px;
  display: flex; align-items: center; justify-content: center;
  font-size: 18px;
  flex-shrink: 0;
}
.app-header-name {
  font-size: 15px; font-weight: 700;
  color: #fff; line-height: 1.1; display: block;
}
.app-header-sub {
  font-size: 11px; color: rgba(255,255,255,0.45);
  display: block; letter-spacing: 0.2px;
}
.app-header-pill {
  display: flex; align-items: center; gap: 6px;
  padding: 5px 12px;
  background: rgba(255,255,255,0.08);
  border: 1px solid rgba(255,255,255,0.12);
  border-radius: 20px;
  font-size: 12px; font-weight: 500;
  color: rgba(255,255,255,0.65);
}
.pill-dot {
  width: 7px; height: 7px; border-radius: 50%; flex-shrink: 0;
}
.pill-dot.all  { background: #10B981; box-shadow: 0 0 0 2px rgba(16,185,129,.3); }
.pill-dot.some { background: #FFE600; }
.pill-dot.none { background: #64748B; }

/* ─────────────────────────────────────────────
   HERO BANNER
───────────────────────────────────────────── */
.hero {
  background: linear-gradient(120deg,#FFE600 0%,#FFCB00 60%,#FFB800 100%);
  border-radius: 16px;
  padding: 26px 30px;
  margin: 22px 0 24px;
  display: flex; align-items: center; gap: 18px;
  position: relative; overflow: hidden;
  box-shadow: 0 4px 20px rgba(255,200,0,.22);
}
.hero::after {
  content: "";
  position: absolute; right: -20px; top: -20px;
  width: 150px; height: 150px;
  background: rgba(255,255,255,.13);
  border-radius: 50%;
}
.hero-icon { font-size: 50px; line-height: 1; flex-shrink: 0; position: relative; z-index: 1; }
.hero-copy  { position: relative; z-index: 1; }
.hero-title { font-size: 20px; font-weight: 800; color: #1A1F71; letter-spacing: -.3px; margin: 0 0 4px; }
.hero-sub   { font-size: 13px; color: rgba(26,31,113,.65); margin: 0; }

/* ─────────────────────────────────────────────
   SECTION LABEL
───────────────────────────────────────────── */
.section-lbl {
  font-size: 11px; font-weight: 700; text-transform: uppercase;
  letter-spacing: .8px; color: #94A3B8;
  display: flex; align-items: center; gap: 10px;
  margin-bottom: 14px;
}
.section-lbl::after { content: ""; flex: 1; height: 1px; background: #E2E8F0; }

/* ─────────────────────────────────────────────
   ACCOUNT CARDS
───────────────────────────────────────────── */
.acc-card {
  background: #fff;
  border-radius: 14px;
  border: 2px solid #E2E8F0;
  overflow: hidden;
  box-shadow: 0 1px 4px rgba(0,0,0,.05);
  margin-bottom: 4px;
}
.acc-card.on {
  border-color: #10B981;
  box-shadow: 0 0 0 4px rgba(16,185,129,.09), 0 2px 8px rgba(0,0,0,.06);
}
.acc-top {
  background: #F8FAFC;
  border-bottom: 1px solid #E2E8F0;
  padding: 16px 20px 13px;
}
.acc-title-row {
  display: flex; align-items: center;
  justify-content: space-between; margin-bottom: 9px;
}
.acc-name { font-size: 17px; font-weight: 800; color: #1A1F71; letter-spacing: -.2px; }
.ml-tag {
  background: #FFE600; color: #1A1F71;
  font-size: 9px; font-weight: 800;
  letter-spacing: 1px; text-transform: uppercase;
  padding: 2px 7px; border-radius: 4px;
}
.acc-badge {
  display: inline-flex; align-items: center; gap: 6px;
  padding: 3px 10px; border-radius: 20px;
  font-size: 12px; font-weight: 600;
}
.acc-badge.on  { background: #ECFDF5; color: #065F46; }
.acc-badge.off { background: #F1F5F9; color: #94A3B8; }
.bdot { width: 6px; height: 6px; border-radius: 50%; display: inline-block; }
.bdot.g { background: #10B981; }
.bdot.s { background: #CBD5E1; }

.acc-body { padding: 18px 20px 4px; min-height: 90px; }
.info-row {
  display: flex; justify-content: space-between; align-items: center;
  padding: 7px 0; border-bottom: 1px solid #F1F5F9;
}
.info-row:last-child { border-bottom: none; }
.info-lbl { font-size: 11px; font-weight: 600; text-transform: uppercase;
            letter-spacing: .5px; color: #94A3B8; }
.info-val { font-size: 13px; font-weight: 600; color: #334155; }
.acc-empty { text-align: center; padding: 12px 0 18px; }
.empty-ico  { font-size: 34px; display: block; margin-bottom: 7px; }
.empty-txt  { font-size: 13px; color: #94A3B8; }
.acc-foot { padding: 13px 20px 18px; }

/* ─────────────────────────────────────────────
   PAINEL UNIFICADO DE CONTAS (combo)
───────────────────────────────────────────── */
.acc-panel {
  background: #fff;
  border-radius: 14px;
  border: 1.5px solid #E2E8F0;
  overflow: hidden;
  box-shadow: 0 1px 4px rgba(0,0,0,.05);
  margin-bottom: 12px;
}
.acc-row-item {
  display: flex; align-items: center;
  justify-content: space-between;
  padding: 16px 20px;
  border-bottom: 1px solid #F1F5F9;
  gap: 12px;
}
.acc-row-item:last-child { border-bottom: none; }
.acc-row-item.connected  { background: #FAFFFE; }
.acc-row-left {
  display: flex; align-items: center; gap: 10px; flex: 1;
}
.acc-row-name {
  font-size: 15px; font-weight: 700; color: #1A1F71; letter-spacing: -.1px;
}
.acc-nick {
  font-size: 12px; font-weight: 500; color: #64748B;
}
.acc-row-tag {
  font-size: 11px; font-weight: 600; text-transform: uppercase;
  letter-spacing: .4px; padding: 3px 9px; border-radius: 20px;
  background: #F1F5F9; color: #94A3B8; flex-shrink: 0;
}
.acc-row-tag.connected { background: #ECFDF5; color: #065F46; }

/* ─────────────────────────────────────────────
   PERÍODO / FORM
───────────────────────────────────────────── */
.period-wrap {
  background: #fff;
  border-radius: 14px;
  border: 1.5px solid #E2E8F0;
  padding: 24px;
  box-shadow: 0 1px 4px rgba(0,0,0,.04);
}

/* ─────────────────────────────────────────────
   STREAMLIT WIDGETS — ajuste de estilo
───────────────────────────────────────────── */
/* Botões primários */
[data-testid="stButton"] > button[kind="primary"] {
  background: #1A1F71 !important;
  border: none !important;
  border-radius: 8px !important;
  font-weight: 600 !important;
  font-family: 'Inter', sans-serif !important;
  transition: all .18s ease !important;
}
[data-testid="stButton"] > button[kind="primary"]:hover {
  background: #252ca0 !important;
  box-shadow: 0 4px 14px rgba(26,31,113,.35) !important;
  transform: translateY(-1px) !important;
}
/* Botões secundários */
[data-testid="stButton"] > button[kind="secondary"] {
  border-radius: 8px !important;
  font-weight: 600 !important;
  font-family: 'Inter', sans-serif !important;
}
/* Date inputs */
[data-testid="stDateInput"] input {
  border-radius: 8px !important;
  font-family: 'Inter', sans-serif !important;
}
/* File uploader */
[data-testid="stFileUploader"] {
  border-radius: 10px !important;
}
/* Tabs */
[data-baseweb="tab-list"] {
  border-radius: 10px !important;
  background: #F1F5F9 !important;
  padding: 4px !important;
  gap: 2px !important;
}
[data-baseweb="tab"] {
  border-radius: 8px !important;
  font-weight: 600 !important;
  font-size: 13px !important;
}
[aria-selected="true"][data-baseweb="tab"] {
  background: #fff !important;
  box-shadow: 0 1px 4px rgba(0,0,0,.1) !important;
}
/* Radio */
[data-testid="stRadio"] label {
  font-size: 13px !important;
  font-weight: 500 !important;
}
/* Divider */
hr { border-color: #E2E8F0 !important; }
</style>
""", unsafe_allow_html=True)

# =============================================================================
# SUPABASE
# =============================================================================

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

# =============================================================================
# OAUTH CALLBACK
# =============================================================================

def _handle_oauth_callback():
    code  = st.query_params.get("code")
    state = st.query_params.get("state", "")
    if not code:
        return
    # state format: "account|verifier" or "account|verifier|next_account"
    parts    = state.split("|")
    account  = parts[0]
    verifier = parts[1] if len(parts) > 1 else ""
    next_acc = parts[2] if len(parts) > 2 else ""
    if account == "ricapet" and "ml_token_ricapet" in st.session_state:
        st.query_params.clear(); return
    if account == "thapets" and "ml_token_thapets" in st.session_state:
        st.query_params.clear(); return
    try:
        if account in ("ricapet", "thapets"):
            cfg    = st.secrets[f"ml_{account}"]
            tokens = ml_api.exchange_code(cfg["client_id"], cfg["client_secret"],
                                          code, REDIRECT_URI, verifier)
            info   = ml_api.get_user_info(tokens["access_token"])
            st.session_state[f"ml_token_{account}"]    = tokens
            st.session_state[f"ml_userid_{account}"]   = info["id"]
            st.session_state[f"ml_nickname_{account}"] = info["nickname"]
            _db_save_rt(account, tokens.get("refresh_token", ""))
            # Se veio pedido de conectar a próxima conta, enfileira
            if next_acc in ("ricapet", "thapets") and next_acc != account:
                st.session_state["_connect_next"] = next_acc
    except Exception as e:
        st.session_state["ml_auth_error"] = str(e)
    st.query_params.clear()

_handle_oauth_callback()

# Auto-redirect para a próxima conta (fluxo "Ambas as contas")
if "_connect_next" in st.session_state:
    _next = st.session_state.pop("_connect_next")
    if f"ml_token_{_next}" not in st.session_state:
        try:
            _cfg = st.secrets[f"ml_{_next}"]
            _ver, _chal = ml_api.generate_pkce()
            st.session_state["_oauth_url"] = ml_api.get_auth_url(
                _cfg["client_id"], REDIRECT_URI,
                state=f"{_next}|{_ver}",
                code_challenge=_chal,
            )
        except Exception:
            pass

# =============================================================================
# AUTO-LOGIN
# =============================================================================

def _auto_authenticate():
    for account in ("ricapet", "thapets"):
        if f"ml_token_{account}" in st.session_state:
            continue
        try:
            cfg = st.secrets[f"ml_{account}"]
            rt  = _db_load_rt(account) or cfg.get("refresh_token", "")
            if not rt:
                continue
            tokens = ml_api.refresh_access_token(cfg["client_id"], cfg["client_secret"], rt)
            info   = ml_api.get_user_info(tokens["access_token"])
            st.session_state[f"ml_token_{account}"]    = tokens
            st.session_state[f"ml_userid_{account}"]   = info["id"]
            st.session_state[f"ml_nickname_{account}"] = info["nickname"]
            new_rt = tokens.get("refresh_token", "")
            if new_rt and new_rt != rt:
                _db_save_rt(account, new_rt)
        except Exception:
            pass

_auto_authenticate()

# =============================================================================
# HEADER CUSTOMIZADO
# =============================================================================

r_conn = "ml_token_ricapet" in st.session_state
t_conn = "ml_token_thapets" in st.session_state

if r_conn and t_conn:
    dot_cls, pill_txt = "all",  "Ambas conectadas"
elif r_conn or t_conn:
    dot_cls, pill_txt = "some", "1 conta conectada"
else:
    dot_cls, pill_txt = "none", "Nenhuma conta conectada"

st.markdown(f"""
<div class="app-header">
  <div class="app-header-brand">
    <div class="app-header-logo">🛒</div>
    <div>
      <span class="app-header-name">Fechamento</span>
      <span class="app-header-sub">Mercado Livre &amp; Shopee</span>
    </div>
  </div>
  <div class="app-header-pill">
    <span class="pill-dot {dot_cls}"></span>
    {pill_txt}
  </div>
</div>
""", unsafe_allow_html=True)

# =============================================================================
# TABS
# =============================================================================

tab_ml, tab_sp = st.tabs(["🛒  Mercado Livre", "🏪  Shopee"])

# =============================================================================
# ABA MERCADO LIVRE
# =============================================================================

with tab_ml:

    modo_ml = st.radio(
        "modo",
        ["🔗  Buscar direto da API", "📂  Enviar arquivo manualmente"],
        horizontal=True,
        label_visibility="collapsed",
    )

    # ── MODO API ────────────────────────────────────────────────────────────────
    if modo_ml == "🔗  Buscar direto da API":

        if st.session_state.get("ml_auth_error"):
            st.error(f"Erro na autenticação: {st.session_state.pop('ml_auth_error')}")

        # Hero
        st.markdown("""
        <div class="hero">
          <div class="hero-icon">🛒</div>
          <div class="hero-copy">
            <p class="hero-title">Mercado Livre — Conexão de Contas</p>
            <p class="hero-sub">Conecte Ricapet e Thapets para buscar pedidos direto da plataforma</p>
          </div>
        </div>
        """, unsafe_allow_html=True)

        # Section label
        st.markdown('<div class="section-lbl">Contas conectadas</div>', unsafe_allow_html=True)

        # ── Painel unificado com status das duas contas ──────────────────────
        contas = [
            ("ricapet", "Ricapet"),
            ("thapets", "Thapets"),
        ]
        rows_html = ""
        for account, label in contas:
            connected = f"ml_token_{account}" in st.session_state
            nickname  = st.session_state.get(f"ml_nickname_{account}", "")
            if connected:
                detalhe = f'<span class="acc-nick">@{nickname}</span>' if nickname else ""
                rows_html += f"""
                <div class="acc-row-item connected">
                  <div class="acc-row-left">
                    <span class="bdot g"></span>
                    <span class="acc-row-name">{label}</span>
                    {detalhe}
                  </div>
                  <span class="acc-row-tag connected">Conectada</span>
                </div>"""
            else:
                rows_html += f"""
                <div class="acc-row-item">
                  <div class="acc-row-left">
                    <span class="bdot s"></span>
                    <span class="acc-row-name">{label}</span>
                  </div>
                  <span class="acc-row-tag">Desconectada</span>
                </div>"""

        st.markdown(f'<div class="acc-panel">{rows_html}</div>',
                    unsafe_allow_html=True)

        # ── Combo (dropdown) + botão de conectar ────────────────────────────
        desconectadas_acc = [(acc, lbl) for acc, lbl in contas
                             if f"ml_token_{acc}" not in st.session_state]

        if desconectadas_acc:
            opcoes_labels = [lbl for _, lbl in desconectadas_acc]
            if len(desconectadas_acc) > 1:
                opcoes_labels = ["Ambas as contas"] + opcoes_labels

            col_sel, col_btn = st.columns([3, 2], gap="small")
            with col_sel:
                escolha = st.selectbox(
                    "Conta",
                    opcoes_labels,
                    label_visibility="collapsed",
                    key="sel_conta",
                )
            with col_btn:
                conectar = st.button("🔗  Conectar", type="primary",
                                     use_container_width=True, key="btn_conectar")
            if conectar:
                try:
                    if escolha == "Ambas as contas":
                        # Conecta primeiro account e passa o segundo no state
                        first_acc, _ = desconectadas_acc[0]
                        second_acc, _ = desconectadas_acc[1]
                        cfg = st.secrets[f"ml_{first_acc}"]
                        verifier, challenge = ml_api.generate_pkce()
                        url = ml_api.get_auth_url(
                            cfg["client_id"], REDIRECT_URI,
                            state=f"{first_acc}|{verifier}|{second_acc}",
                            code_challenge=challenge,
                        )
                    else:
                        account = escolha.lower()
                        cfg = st.secrets[f"ml_{account}"]
                        verifier, challenge = ml_api.generate_pkce()
                        url = ml_api.get_auth_url(
                            cfg["client_id"], REDIRECT_URI,
                            state=f"{account}|{verifier}",
                            code_challenge=challenge,
                        )
                    st.session_state["_oauth_url"] = url
                    st.rerun()
                except (KeyError, FileNotFoundError) as e:
                    st.warning(f"Credenciais não configuradas: {e}")

        # ── Botões de desconectar ────────────────────────────────────────────
        conectadas = [(acc, label) for acc, label in contas
                      if f"ml_token_{acc}" in st.session_state]
        if conectadas:
            cols = st.columns(len(conectadas))
            for col_d, (account, label) in zip(cols, conectadas):
                with col_d:
                    if st.button(f"Desconectar {label}", key=f"disc_{account}",
                                 use_container_width=True):
                        for k in (f"ml_token_{account}",
                                  f"ml_userid_{account}",
                                  f"ml_nickname_{account}"):
                            st.session_state.pop(k, None)
                        st.rerun()

        # Redirect OAuth
        if "_oauth_url" in st.session_state:
            oauth_url = st.session_state.pop("_oauth_url")
            st.markdown(
                f'<meta http-equiv="refresh" content="0;url={oauth_url}">',
                unsafe_allow_html=True,
            )
            st.stop()

        # ── Buscar pedidos ───────────────────────────────────────────────────────
        tem_r = "ml_token_ricapet" in st.session_state
        tem_t = "ml_token_thapets" in st.session_state

        if tem_r or tem_t:
            st.divider()
            st.markdown('<div class="section-lbl">Período de busca</div>',
                        unsafe_allow_html=True)

            hoje = date.today()
            with st.container():
                st.markdown('<div class="period-wrap">', unsafe_allow_html=True)
                col_d1, col_d2 = st.columns(2)
                with col_d1:
                    data_ini = st.date_input("De", value=hoje.replace(day=1),
                                             key="ml_api_ini")
                with col_d2:
                    data_fim = st.date_input("Até", value=hoje, key="ml_api_fim")
                tabela_api = st.file_uploader("TABELA_AUXILIAR.xlsx",
                                              type="xlsx", key="ml_api_aux")
                btn_api = st.button("🔍  Buscar e Consolidar", type="primary",
                                    key="btn_ml_api", use_container_width=True)
                st.markdown('</div>', unsafe_allow_html=True)

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
                            if tem_r:
                                tok  = st.session_state["ml_token_ricapet"]
                                uid  = st.session_state["ml_userid_ricapet"]
                                ords = ml_api.fetch_orders(tok["access_token"],
                                                           uid, from_str, to_str)
                                st.info(f"Ricapet: {len(ords)} pedido(s)")
                                if ords:
                                    arquivos.append(("relatorio_736787693.xlsx",
                                                     ml_api.orders_to_excel_bytes(ords, "Ricapet")))
                                    total += len(ords)
                            if tem_t:
                                tok  = st.session_state["ml_token_thapets"]
                                uid  = st.session_state["ml_userid_thapets"]
                                ords = ml_api.fetch_orders(tok["access_token"],
                                                           uid, from_str, to_str)
                                st.info(f"Thapets: {len(ords)} pedido(s)")
                                if ords:
                                    arquivos.append(("relatorio_1139210125.xlsx",
                                                     ml_api.orders_to_excel_bytes(ords, "Thapets")))
                                    total += len(ords)
                            if not arquivos:
                                st.warning("Nenhum pedido encontrado no período.")
                            else:
                                tabela_bytes     = tabela_api.read()
                                xlsx_bytes, logs = ml_core.processar(arquivos, tabela_bytes)
                                ts               = datetime.now().strftime("%Y%m%d_%H%M%S")
                                st.success(f"✅  Consolidado! {total} pedido(s) processado(s).")
                                st.download_button(
                                    "⬇️  Baixar Excel gerado",
                                    data=xlsx_bytes,
                                    file_name=f"MercadoLivre_{ts}.xlsx",
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                    key="dl_ml_api",
                                )
                                with st.expander("Log de processamento"):
                                    for msg in logs:
                                        st.text(msg)
                        except Exception as exc:
                            st.error(f"Erro: {exc}")

    # ── MODO ARQUIVO ─────────────────────────────────────────────────────────────
    else:
        col1, col2 = st.columns(2)
        with col1:
            relatorios_ml = st.file_uploader(
                "Relatórios ML (.xlsx)", type="xlsx",
                accept_multiple_files=True, key="ml_rel",
                help="Arquivos com 736787693 ou 1139210125 no nome.",
            )
        with col2:
            tabela_ml = st.file_uploader("TABELA_AUXILIAR.xlsx",
                                         type="xlsx", key="ml_aux")
        btn_ml = st.button("Consolidar Mercado Livre", type="primary", key="btn_ml")
        if btn_ml:
            if not relatorios_ml:
                st.error("Selecione pelo menos um relatório.")
            elif not tabela_ml:
                st.error("Selecione a TABELA_AUXILIAR.xlsx.")
            else:
                with st.spinner("Processando..."):
                    try:
                        arquivos     = [(f.name, f.read()) for f in relatorios_ml]
                        tabela_bytes = tabela_ml.read()
                        xlsx_bytes, logs = ml_core.processar(arquivos, tabela_bytes)
                        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
                        st.success(f"✅  Consolidado! {len(arquivos)} arquivo(s).")
                        st.download_button(
                            "⬇️  Baixar Excel gerado",
                            data=xlsx_bytes,
                            file_name=f"MercadoLivre_{ts}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key="dl_ml",
                        )
                        with st.expander("Log"):
                            for msg in logs: st.text(msg)
                    except Exception as exc:
                        st.error(f"Erro: {exc}")

# =============================================================================
# ABA SHOPEE
# =============================================================================

with tab_sp:
    st.subheader("Consolidar Shopee")
    col1, col2 = st.columns(2)
    with col1:
        relatorios_sp = st.file_uploader(
            "Relatórios Shopee (.xlsx)", type="xlsx",
            accept_multiple_files=True, key="sp_rel",
            help="Nome deve começar com 'order.all.'",
        )
    with col2:
        tabela_sp = st.file_uploader("TABELA_AUXILIAR.xlsx",
                                     type="xlsx", key="sp_aux")
    btn_sp = st.button("Consolidar Shopee", type="primary", key="btn_sp")
    if btn_sp:
        if not relatorios_sp:
            st.error("Selecione pelo menos um relatório.")
        elif not tabela_sp:
            st.error("Selecione a TABELA_AUXILIAR.xlsx.")
        else:
            with st.spinner("Processando..."):
                try:
                    arquivos     = [(f.name, f.read()) for f in relatorios_sp]
                    tabela_bytes = tabela_sp.read()
                    xlsx_bytes, logs = shopee_core.processar(arquivos, tabela_bytes)
                    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
                    st.success("✅  Consolidado com sucesso!")
                    st.download_button(
                        "⬇️  Baixar Excel gerado",
                        data=xlsx_bytes,
                        file_name=f"Shopee_{ts}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="dl_sp",
                    )
                    with st.expander("Log"):
                        for msg in logs: st.text(msg)
                except Exception as exc:
                    st.error(f"Erro: {exc}")
