from __future__ import annotations

import tempfile
from datetime import datetime
from pathlib import Path

import pandas as pd
import streamlit as st

import itbi_core as core

st.set_page_config(
    page_title="Imóveis SP | ITBI Analytics",
    page_icon="🏙️",
    layout="wide",
    initial_sidebar_state="collapsed",
)

LINCE_LOGO_WHITE = "https://lincepartners.com.br/wp-content/themes/lincepartners2021_vs1/images/logo-branco.png"

st.markdown(
    """
<style>
:root { --navy:#0f2742; --blue:#2563eb; --muted:#60748a; --bg:#f6f8fb; --green:#0f9d76; }
.stApp { background: var(--bg); }
.block-container { max-width: 1450px; padding-top: 1.15rem; padding-bottom: 3rem; }
.hero { background: linear-gradient(135deg,#0f2742,#173b62); color:white; border-radius:22px; padding:22px 26px; margin-bottom:15px; box-shadow:0 8px 30px rgba(15,39,66,.16); display:flex; justify-content:space-between; align-items:center; gap:24px; }
.hero-copy { min-width:0; }
.hero h1 { margin:0; font-size:1.7rem; letter-spacing:-.025em; }
.hero p { margin:.35rem 0 0; color:#d9e6f3; font-size:.94rem; }
.hero-brand { text-align:right; flex:0 0 235px; }
.hero-brand img { width:170px; max-width:100%; object-fit:contain; margin-bottom:5px; }
.hero-brand .madeby { color:#d9e6f3; font-size:.72rem; line-height:1.25; }
.pill { display:inline-block; background:#e8f1ff; color:#164b8d; border-radius:999px; padding:6px 10px; font-weight:700; font-size:.78rem; margin-right:6px; margin-top:10px; }
.status-card { background:white; border:1px solid #e4ebf3; border-radius:16px; padding:13px 16px; box-shadow:0 4px 16px rgba(20,45,70,.04); min-height:72px; }
.status-title { font-weight:800; color:#153653; margin-bottom:3px; }
.status-ok { color:#0f8b65; font-weight:800; }
.status-empty { color:#b7791f; font-weight:800; }
.small-muted { color:#718397; font-size:.83rem; }
.update-box { background:white; border:1px solid #dce7f2; border-radius:18px; padding:18px; margin:10px 0 18px; box-shadow:0 5px 20px rgba(20,45,70,.05); }
div[data-testid="stMetric"] { background:white; border:1px solid #e6edf5; padding:15px 16px; border-radius:16px; box-shadow:0 4px 14px rgba(20,45,70,.04); }
.stButton > button { border-radius:12px; font-weight:700; min-height:42px; }
.stDownloadButton > button { border-radius:12px; font-weight:700; }
div[data-baseweb="select"] > div { border-radius:12px; }
[data-testid="stAlert"] { border-radius:14px; }
@media (max-width: 700px) {
  .block-container { padding-left:.75rem; padding-right:.75rem; padding-top:.65rem; }
  .hero { padding:17px; border-radius:18px; align-items:flex-start; gap:10px; }
  .hero h1 { font-size:1.28rem; }
  .hero p { font-size:.82rem; }
  .hero-brand { flex:0 0 118px; }
  .hero-brand img { width:110px; }
  .hero-brand .madeby { font-size:.58rem; }
}
</style>
""",
    unsafe_allow_html=True,
)


def fmt_period(year: int | None, month: int | None) -> str:
    if not year or not month:
        return "base ainda não preparada"
    return f"{core.MONTH_NAMES.get(month, month).upper()}/{year}"


def status_snapshot():
    try:
        return core.database_status()
    except Exception:
        return 0, 0, None, None, None


def rows_to_dataframe(rows: list[list[object]]) -> pd.DataFrame:
    cols = ["Ano", "Mês", "Aba de origem", "Linha de origem"] + core.CANONICAL_HEADERS
    return pd.DataFrame(rows, columns=cols)


def analytics_matrix(analytics: core.AnalyticsResult) -> pd.DataFrame:
    data = []
    for area in analytics.area_keys:
        row = {"Área Construída": core.format_area(area)}
        for year in analytics.years:
            stats = analytics.cells.get((area, year))
            row[str(year)] = core.format_brl_compact(stats.average) if stats else ""
        data.append(row)
    return pd.DataFrame(data)


def chart_frame(analytics: core.AnalyticsResult) -> pd.DataFrame:
    rows = []
    for area in analytics.area_keys:
        if area is None:
            continue
        for year in analytics.years:
            stats = analytics.cells.get((area, year))
            if stats and stats.average is not None:
                rows.append({"Ano": year, "Área": f"{area} m²", "Preço médio": stats.average})
    return pd.DataFrame(rows)


@st.cache_data(ttl=60, show_spinner=False)
def cached_street_catalog(_row_count: int):
    # row_count participa da chave e renova o catálogo automaticamente após uma atualização.
    return core.list_streets()


def set_year_preset(years: list[int]) -> None:
    st.session_state["years_to_update"] = years


current_year = datetime.now().year
all_years = list(range(2006, current_year + 1))
default_years = list(range(max(2020, 2006), current_year + 1))

if "show_update_panel" not in st.session_state:
    st.session_state["show_update_panel"] = False
if "years_to_update" not in st.session_state:
    st.session_state["years_to_update"] = default_years

years_count, row_count, last_update, latest_year, latest_month = status_snapshot()
period = fmt_period(latest_year, latest_month)
loaded_years = core.indexed_years() if years_count else []

st.markdown(
    f"""
<div class="hero">
  <div class="hero-copy">
    <h1>Imóveis SP — ITBI Analytics</h1>
    <p>Histórico de transações imobiliárias com recolhimento de ITBI — Município de São Paulo</p>
    <div><span class="pill">Base: {period}</span><span class="pill">{years_count} ano(s) indexado(s)</span></div>
  </div>
  <div class="hero-brand">
    <img src="{LINCE_LOGO_WHITE}" alt="Lince Partners">
    <div class="madeby">Made by Italo Ferrara, programador wanna-be</div>
  </div>
</div>
""",
    unsafe_allow_html=True,
)

status_col, update_col = st.columns([2.4, 1], gap="medium", vertical_alignment="center")
with status_col:
    if years_count:
        loaded_text = (
            f"{min(loaded_years)}–{max(loaded_years)}" if loaded_years and len(loaded_years) > 1 else
            (str(loaded_years[0]) if loaded_years else "—")
        )
        st.markdown(
            f"""
<div class="status-card">
  <div class="status-title"><span class="status-ok">● Base preparada</span> &nbsp; • &nbsp; último período: <b>{period}</b></div>
  <div class="small-muted">{row_count:,} registros indexados • anos carregados: {loaded_text}</div>
</div>
""".replace(",", "."),
            unsafe_allow_html=True,
        )
    else:
        st.markdown(
            """
<div class="status-card">
  <div class="status-title"><span class="status-empty">● Base ainda não preparada</span></div>
  <div class="small-muted">Escolha os anos e carregue os dados diretamente da Prefeitura para começar.</div>
</div>
""",
            unsafe_allow_html=True,
        )

with update_col:
    if st.button(
        "🔄 Atualizar Base da Prefeitura",
        type="primary",
        use_container_width=True,
        help="Escolha quais anos baixar/indexar. Anos já carregados são preservados.",
    ):
        st.session_state["show_update_panel"] = not st.session_state["show_update_panel"]

if st.session_state["show_update_panel"]:
    with st.container(border=True):
        st.markdown("#### Atualizar Base da Prefeitura")
        st.caption(
            "Escolha apenas os anos que quer carregar agora. Os anos que já estão na base **não são apagados**. "
            "Para uma consulta mais recente e rápida, 2020 até o ano atual costuma ser um bom ponto de partida."
        )

        p1, p2, p3 = st.columns(3)
        p1.button(
            f"2020–{current_year}",
            use_container_width=True,
            on_click=set_year_preset,
            args=(list(range(2020, current_year + 1)),),
        )
        p2.button(
            "Últimos 5 anos",
            use_container_width=True,
            on_click=set_year_preset,
            args=(list(range(max(2006, current_year - 4), current_year + 1)),),
        )
        p3.button(
            f"Todos (2006–{current_year})",
            use_container_width=True,
            on_click=set_year_preset,
            args=(all_years,),
        )

        selected_years = st.multiselect(
            "Anos a carregar nesta atualização",
            options=all_years,
            key="years_to_update",
            placeholder="Selecione um ou mais anos",
        )
        force_selected = st.checkbox(
            "Baixar novamente e reindexar os anos selecionados",
            value=False,
            help="Normalmente deixe desmarcado. Use se suspeitar que um arquivo histórico foi corrigido pela Prefeitura.",
        )

        if loaded_years:
            st.caption("Já carregados: " + ", ".join(map(str, loaded_years)))

        action_col, close_col = st.columns([2, 1])
        with action_col:
            do_update = st.button(
                "⬇️ Carregar anos selecionados",
                type="primary",
                use_container_width=True,
                disabled=not selected_years,
            )
        with close_col:
            if st.button("Fechar", use_container_width=True):
                st.session_state["show_update_panel"] = False
                st.rerun()

        if do_update:
            progress_bar = st.progress(0.0)
            status_box = st.empty()

            def _progress(message, fraction):
                status_box.info(message)
                if fraction is not None:
                    progress_bar.progress(max(0.0, min(1.0, float(fraction))))

            try:
                summary = core.update_database(
                    force_all=force_selected,
                    progress=_progress,
                    selected_years=selected_years,
                )
                progress_bar.progress(1.0)
                indexed_labels = [str(year) for year, _rows in summary.get("indexed", [])]
                skipped_labels = [str(year) for year in summary.get("skipped", [])]
                message_parts = []
                if indexed_labels:
                    message_parts.append("indexados: " + ", ".join(indexed_labels))
                if skipped_labels:
                    message_parts.append("já estavam atualizados: " + ", ".join(skipped_labels))
                status_box.success("Atualização concluída" + (" — " + " • ".join(message_parts) if message_parts else "."))
                st.cache_data.clear()
                st.session_state.pop("result_rows", None)
                st.session_state["show_update_panel"] = False
                st.rerun()
            except Exception as exc:
                st.error(f"Não foi possível atualizar a base: {exc}")

if years_count == 0:
    st.info(
        "👆 **Comece por aqui:** toque em **Atualizar Base da Prefeitura**, escolha os anos desejados "
        "(por exemplo, 2020 até o ano atual) e carregue a base."
    )

st.markdown("### Pesquisar imóvel")
if years_count:
    street_catalog = cached_street_catalog(row_count)
    street_labels = [item.display for item in street_catalog]
    street_by_label = {item.display: item for item in street_catalog}

    search_col1, search_col2 = st.columns([2.2, 1], gap="large")
    with search_col1:
        selected_street_label = st.selectbox(
            "Logradouro",
            options=street_labels,
            index=None,
            placeholder="Clique aqui e digite parte do nome da rua...",
            filter_mode="contains",
            key="street_choice_v02",
            help="O menu procura por qualquer trecho do nome e mostra também os bairros encontrados na base.",
        )
        chosen_street = street_by_label.get(selected_street_label) if selected_street_label else None
        if chosen_street:
            st.caption(f"Bairro(s) mapeado(s): {chosen_street.neighborhoods or 'não informado'}")

    with search_col2:
        if chosen_street:
            number_suggestions = core.suggest_numbers(
                chosen_street.value,
                selected_street_norm=chosen_street.key,
                limit=10000,
            )
            number_labels = [item.display for item in number_suggestions]
            number_by_label = {item.display: item for item in number_suggestions}
            selected_number_label = st.selectbox(
                "Número",
                options=number_labels,
                index=None,
                placeholder="Clique e digite o número...",
                filter_mode="contains",
                key=f"number_choice_v02_{chosen_street.key}",
            )
            chosen_number = number_by_label.get(selected_number_label) if selected_number_label else None
            if chosen_number and chosen_number.neighborhoods:
                st.caption(f"Bairro: {chosen_number.neighborhoods}")
        else:
            chosen_number = None
            st.selectbox(
                "Número",
                ["Selecione primeiro o logradouro"],
                disabled=True,
                key="number_disabled_v02",
            )

    search_disabled = not (chosen_street and chosen_number)
    if st.button("🔎 Analisar endereço", type="primary", disabled=search_disabled, use_container_width=True):
        with st.spinner("Buscando o histórico do endereço nos anos carregados..."):
            try:
                hits = core.find_hits(chosen_street.value, chosen_number.value, mode="exact")
                if not hits:
                    st.session_state.pop("result_rows", None)
                    st.warning("Nenhuma transação encontrada para esse endereço nos anos carregados.")
                else:
                    result_rows = core.load_source_rows(hits, lambda _message, _fraction: None)
                    st.session_state["result_rows"] = result_rows
                    st.session_state["result_street"] = chosen_street.value
                    st.session_state["result_number"] = chosen_number.value
            except Exception as exc:
                st.error(f"Erro ao consultar o endereço: {exc}")
else:
    st.selectbox(
        "Logradouro",
        ["Carregue a base para habilitar a pesquisa"],
        disabled=True,
        key="street_disabled_v02",
    )
    st.selectbox(
        "Número",
        ["Carregue a base para habilitar a pesquisa"],
        disabled=True,
        key="number_disabled_empty_v02",
    )

rows = st.session_state.get("result_rows")
if rows:
    analytics = core.build_analytics(rows)
    street = st.session_state.get("result_street", "")
    number = st.session_state.get("result_number", "")

    st.markdown(f"### {street}, {number}")
    st.caption(
        "Indicadores e matriz consideram exclusivamente **Natureza da Transação = 1. Compra e Venda**, "
        "o **Valor de Transação (declarado pelo contribuinte)** e a **Área Construída (m²)**."
    )

    k1, k2, k3, k4, k5 = st.columns(5)
    k1.metric("Transações encontradas", analytics.total_records)
    k2.metric("Compras e vendas", analytics.purchase_sale_records)
    k3.metric("Preço médio declarado", core.format_brl_compact(analytics.overall_average) or "—")
    k4.metric("Tipologias por m²", analytics.known_typologies)
    k5.metric("Anos com negócios", analytics.years_with_sales)

    tab1, tab2, tab3 = st.tabs(["📊 Preço por tipologia", "📈 Evolução", "📋 Transações"])

    with tab1:
        matrix = analytics_matrix(analytics)
        if matrix.empty:
            st.info("Não há compras e vendas com preço declarado positivo para montar a matriz.")
        else:
            st.dataframe(matrix, use_container_width=True, hide_index=True, height=min(650, 80 + 35 * len(matrix)))
            st.caption("Cada linha representa a Área Construída arredondada ao m² inteiro; cada célula é o preço médio declarado naquele ano.")

    with tab2:
        chart_df = chart_frame(analytics)
        if chart_df.empty:
            st.info("Não há série histórica suficiente para exibir o gráfico.")
        else:
            pivot = chart_df.pivot(index="Ano", columns="Área", values="Preço médio")
            st.line_chart(pivot, use_container_width=True)
            st.caption("Preço médio declarado das compras e vendas por tipologia de Área Construída.")

    with tab3:
        df = rows_to_dataframe(rows)
        preferred = [
            "Ano", "Mês", "Data de Transação", "Nome do Logradouro", "Número", "Complemento", "Bairro",
            "Natureza de Transação", "Valor de Transação (declarado pelo contribuinte)", "Área Construída (m2)",
            "Descrição do uso (IPTU)", "N° do Cadastro (SQL)", "Matrícula do Imóvel"
        ]
        remaining = [c for c in df.columns if c not in preferred]
        display_df = df[preferred + remaining]
        st.dataframe(display_df, use_container_width=True, hide_index=True, height=620)
        st.caption("O campo **Complemento** é mantido integralmente no dado de origem e no Excel exportado.")

    with tempfile.TemporaryDirectory() as tmpdir:
        outfile = Path(tmpdir) / f"ITBI_{core.normalize_street(street).replace(' ', '_')}_{core.normalize_number(number)}.xlsx"
        core.export_results(outfile, street, number, "exact", rows, analytics)
        excel_bytes = outfile.read_bytes()
    st.download_button(
        "⬇️ Baixar relatório Excel",
        data=excel_bytes,
        file_name=f"ITBI_{core.normalize_street(street).replace(' ', '_')}_{core.normalize_number(number)}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

st.divider()
st.caption("Fonte dos dados: Prefeitura de São Paulo • ITBI Analytics WEB-0.2")
