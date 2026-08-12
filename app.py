from __future__ import annotations

import tempfile
from pathlib import Path

import pandas as pd
import streamlit as st

import itbi_core as core

st.set_page_config(
    page_title="ITBI Analytics SP",
    page_icon="🏙️",
    layout="wide",
    initial_sidebar_state="collapsed",
)

st.markdown(
    """
<style>
:root { --navy:#0f2742; --blue:#2563eb; --muted:#60748a; --bg:#f6f8fb; }
.stApp { background: var(--bg); }
.block-container { max-width: 1450px; padding-top: 1.3rem; padding-bottom: 3rem; }
.hero { background: linear-gradient(135deg,#0f2742,#173b62); color:white; border-radius:22px; padding:24px 28px; margin-bottom:18px; box-shadow:0 8px 30px rgba(15,39,66,.16); }
.hero h1 { margin:0; font-size:1.65rem; letter-spacing:-.02em; }
.hero p { margin:.35rem 0 0; color:#d9e6f3; font-size:.95rem; }
.pill { display:inline-block; background:#e8f1ff; color:#164b8d; border-radius:999px; padding:6px 10px; font-weight:700; font-size:.78rem; margin-right:6px; }
.card { background:white; border:1px solid #e7edf4; border-radius:18px; padding:18px; box-shadow:0 4px 18px rgba(20,45,70,.05); }
.small-muted { color:#718397; font-size:.83rem; }
div[data-testid="stMetric"] { background:white; border:1px solid #e6edf5; padding:15px 16px; border-radius:16px; box-shadow:0 4px 14px rgba(20,45,70,.04); }
.stButton > button { border-radius:12px; font-weight:700; }
.stDownloadButton > button { border-radius:12px; font-weight:700; }
@media (max-width: 700px) {
  .block-container { padding-left:.8rem; padding-right:.8rem; padding-top:.75rem; }
  .hero { padding:18px; border-radius:18px; }
  .hero h1 { font-size:1.35rem; }
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
    rows=[]
    for area in analytics.area_keys:
        if area is None:
            continue
        for year in analytics.years:
            stats=analytics.cells.get((area,year))
            if stats and stats.average is not None:
                rows.append({"Ano":year,"Área":f"{area} m²","Preço médio":stats.average})
    return pd.DataFrame(rows)


years_count, row_count, last_update, latest_year, latest_month = status_snapshot()
period = fmt_period(latest_year, latest_month)

st.markdown(
    f"""
<div class="hero">
  <h1>ITBI Analytics SP</h1>
  <p>Histórico de transações imobiliárias com recolhimento de ITBI — Município de São Paulo</p>
  <div style="margin-top:13px"><span class="pill">Base: {period}</span><span class="pill">{years_count} anos indexados</span></div>
</div>
""",
    unsafe_allow_html=True,
)

with st.sidebar:
    st.header("Base de dados")
    st.caption("Os arquivos são obtidos diretamente da Prefeitura de São Paulo.")
    st.write(f"**Último período:** {period}")
    st.write(f"**Registros indexados:** {row_count:,}".replace(",", "."))
    force_all = st.checkbox("Rebaixar e reindexar todos os anos", value=False)
    if st.button("Atualizar base da Prefeitura", use_container_width=True):
        progress_bar = st.progress(0.0)
        status_box = st.empty()

        def _progress(message, fraction):
            status_box.info(message)
            if fraction is not None:
                progress_bar.progress(max(0.0, min(1.0, float(fraction))))

        try:
            summary = core.update_database(force_all=force_all, progress=_progress)
            progress_bar.progress(1.0)
            status_box.success(
                f"Atualização concluída. {len(summary.get('indexed', []))} ano(s) reindexado(s)."
            )
            st.rerun()
        except Exception as exc:
            st.error(f"Não foi possível atualizar a base: {exc}")

    st.divider()
    st.caption("Fonte oficial")
    st.link_button("Prefeitura de São Paulo", core.SOURCE_PAGE, use_container_width=True)

if years_count == 0:
    st.warning(
        "A base deste servidor ainda não foi preparada. Abra o menu lateral ☰ e toque em **Atualizar base da Prefeitura**. "
        "Na primeira carga, o servidor baixa e indexa os arquivos históricos; depois as consultas ficam rápidas."
    )

st.subheader("Pesquisar imóvel")
search_col1, search_col2 = st.columns([2.2, 1], gap="large")

with search_col1:
    street_text = st.text_input(
        "Logradouro",
        placeholder="Digite parte do nome da rua, avenida, alameda...",
        key="street_text",
    )
    street_suggestions = core.suggest_streets(street_text, limit=40) if street_text else []
    if street_suggestions:
        street_labels = [item.display for item in street_suggestions]
        chosen_label = st.selectbox(
            "Selecione o logradouro",
            street_labels,
            index=0,
            key="street_choice",
        )
        chosen_street = next(i for i in street_suggestions if i.display == chosen_label)
        st.caption(f"Bairro(s) encontrado(s): {chosen_street.neighborhoods or 'não informado'}")
    else:
        chosen_street = None

with search_col2:
    if chosen_street:
        number_suggestions = core.suggest_numbers(
            chosen_street.value,
            selected_street_norm=chosen_street.key,
            limit=200,
        )
        number_labels = [item.display for item in number_suggestions]
        selected_number_label = st.selectbox(
            "Número",
            number_labels if number_labels else ["Nenhum número encontrado"],
            index=0,
            key="number_choice",
        )
        chosen_number = (
            next((i for i in number_suggestions if i.display == selected_number_label), None)
            if number_suggestions else None
        )
        if chosen_number and chosen_number.neighborhoods:
            st.caption(f"Bairro: {chosen_number.neighborhoods}")
    else:
        chosen_number = None
        st.selectbox("Número", ["Selecione primeiro o logradouro"], disabled=True)

search_disabled = not (chosen_street and chosen_number)
if st.button("🔎 Analisar endereço", type="primary", disabled=search_disabled, use_container_width=True):
    with st.spinner("Buscando todo o histórico do endereço..."):
        try:
            hits = core.find_hits(chosen_street.value, chosen_number.value, mode="exact")
            if not hits:
                st.session_state.pop("result_rows", None)
                st.warning("Nenhuma transação encontrada para esse endereço.")
            else:
                def _silent_progress(_message, _fraction):
                    pass
                result_rows = core.load_source_rows(hits, _silent_progress)
                st.session_state["result_rows"] = result_rows
                st.session_state["result_street"] = chosen_street.value
                st.session_state["result_number"] = chosen_number.value
        except Exception as exc:
            st.error(f"Erro ao consultar o endereço: {exc}")

rows = st.session_state.get("result_rows")
if rows:
    analytics = core.build_analytics(rows)
    street = st.session_state.get("result_street", "")
    number = st.session_state.get("result_number", "")

    st.markdown(f"### {street}, {number}")
    st.caption(
        "Indicadores e matriz consideram exclusivamente Natureza da Transação = **1. Compra e Venda**, "
        "o **Valor de Transação (declarado pelo contribuinte)** e a **Área Construída (m²)**."
    )

    k1,k2,k3,k4,k5 = st.columns(5)
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
            st.dataframe(matrix, use_container_width=True, hide_index=True, height=min(650, 80+35*len(matrix)))
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
            "Ano","Mês","Data de Transação","Nome do Logradouro","Número","Complemento","Bairro",
            "Natureza de Transação","Valor de Transação (declarado pelo contribuinte)","Área Construída (m2)",
            "Descrição do uso (IPTU)","N° do Cadastro (SQL)","Matrícula do Imóvel"
        ]
        remaining = [c for c in df.columns if c not in preferred]
        display_df = df[preferred + remaining]
        st.dataframe(display_df, use_container_width=True, hide_index=True, height=620)
        st.caption("O campo **Complemento** é exibido integralmente na base e no Excel exportado; nenhuma informação é truncada no dado de origem.")

    with tempfile.TemporaryDirectory() as tmpdir:
        outfile = Path(tmpdir) / f"ITBI_{core.normalize_street(street).replace(' ','_')}_{core.normalize_number(number)}.xlsx"
        core.export_results(outfile, street, number, "exact", rows, analytics)
        excel_bytes = outfile.read_bytes()
    st.download_button(
        "⬇️ Baixar relatório Excel",
        data=excel_bytes,
        file_name=f"ITBI_{core.normalize_street(street).replace(' ','_')}_{core.normalize_number(number)}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

st.divider()
st.caption("ITBI Analytics SP • Dados públicos da Secretaria Municipal da Fazenda de São Paulo")
