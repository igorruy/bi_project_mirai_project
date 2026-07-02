"""Aplicação Streamlit para gerar e visualizar os BIs do Projeto Mirai."""

from __future__ import annotations

import tempfile
from datetime import datetime
from pathlib import Path

import streamlit as st
import streamlit.components.v1 as components

from mirai_reports import (
    build_scurve,
    compute_frente_metrics,
    generate_atrasos,
    generate_gerencial,
    load_xml,
    next_sunday,
    read_status_date,
)

st.set_page_config(
    page_title="Mirai BI Reports",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
    <style>
    .main .block-container{padding-top:2rem;max-width:1280px}
    div[data-testid="stMetric"]{background:#fff;border:1px solid #e2e5eb;border-radius:10px;padding:14px}
    .mirai-note{background:#f8fafc;border:1px solid #e2e5eb;border-radius:10px;padding:14px 16px;color:#334155}
    </style>
    """,
    unsafe_allow_html=True,
)

st.title("📊 Mirai BI Reports")
st.caption("Upload do XML do MS Project para gerar automaticamente os relatórios gerencial e de atrasos em HTML.")

with st.sidebar:
    st.header("Entrada")
    uploaded_xml = st.file_uploader(
        "Arquivo XML do MS Project",
        type=["xml"],
        help="Suporta o XML atual (~163 MB). No Streamlit Community Cloud, mantenha o arquivo abaixo do limite configurado de 200 MB.",
    )
    use_status_override = st.checkbox(
        "Informar data de status manualmente",
        value=False,
        help="Se desmarcado, a aplicação lê o StatusDate do XML; se inexistente, usa a data atual.",
    )
    status_override = None
    if use_status_override:
        status_override = st.date_input(
            "Data de status",
            value=datetime.today(),
            format="DD/MM/YYYY",
        )
    render_html = st.toggle("Exibir HTMLs completos na tela", value=True)
    st.markdown("---")
    st.markdown("**Saídas geradas**")
    st.markdown("- `report_gerencial_mirai.html`\n- `report_atrasos_mirai.html`")

if uploaded_xml is None:
    st.markdown(
        """
        <div class="mirai-note">
        <b>Como usar:</b> envie o XML exportado do Microsoft Project na barra lateral. A aplicação calcula os mesmos indicadores do script atual e disponibiliza os dois HTMLs para visualização e download.
        </div>
        """,
        unsafe_allow_html=True,
    )
    st.stop()

with st.status("Processando XML e gerando BIs...", expanded=True) as status:
    with tempfile.TemporaryDirectory() as tmpdir:
        tmp_path = Path(tmpdir)
        xml_path = tmp_path / uploaded_xml.name
        xml_path.write_bytes(uploaded_xml.getbuffer())
        out_dir = tmp_path / "reports"
        out_dir.mkdir(parents=True, exist_ok=True)

        st.write("Lendo data de status...")
        status_date = (
            datetime.combine(status_override, datetime.min.time())
            if status_override is not None
            else read_status_date(xml_path)
        )
        sun = next_sunday(status_date)

        st.write("Lendo tarefas do XML...")
        leaf_tasks, delay_tasks = load_xml(xml_path, status_date)

        st.write("Calculando Curva S e métricas por workstream...")
        sc_labels, sc_bl0, sc_bl5, sc_ev = build_scurve(leaf_tasks, status_date)
        frente_spi, fr_delay, overall = compute_frente_metrics(leaf_tasks, delay_tasks, status_date)

        gerencial_path = out_dir / "report_gerencial_mirai.html"
        atrasos_path = out_dir / "report_atrasos_mirai.html"

        st.write("Montando HTMLs finais...")
        generate_gerencial(
            sc_labels, sc_bl0, sc_bl5, sc_ev,
            frente_spi, fr_delay, overall,
            status_date, uploaded_xml.name, gerencial_path,
        )
        generate_atrasos(
            delay_tasks, fr_delay,
            status_date, sun, uploaded_xml.name, atrasos_path,
        )

        gerencial_html = gerencial_path.read_text(encoding="utf-8")
        atrasos_html = atrasos_path.read_text(encoding="utf-8")

    status.update(label="BIs gerados com sucesso!", state="complete", expanded=False)

st.subheader("Indicadores principais")
col1, col2, col3, col4, col5 = st.columns(5)
col1.metric("SPI Global vs BL5", f"{overall['spi']:.2f}")
col2.metric("EV realizado", f"{overall['ev']:.0f} h", f"{overall['pct']}%")
col3.metric("PV planejado", f"{overall['pv']:.0f} h")
col4.metric("Tarefas atrasadas", f"{len(delay_tasks)}")
col5.metric("Data status", status_date.strftime("%d/%m/%Y"), f"Sem. {status_date.isocalendar()[1]}")

st.subheader("Downloads")
dl1, dl2 = st.columns(2)
dl1.download_button(
    "⬇️ Baixar report_gerencial_mirai.html",
    gerencial_html,
    file_name="report_gerencial_mirai.html",
    mime="text/html",
    use_container_width=True,
)
dl2.download_button(
    "⬇️ Baixar report_atrasos_mirai.html",
    atrasos_html,
    file_name="report_atrasos_mirai.html",
    mime="text/html",
    use_container_width=True,
)

if render_html:
    tab1, tab2 = st.tabs(["BI Gerencial", "BI de Atrasos"])
    with tab1:
        components.html(gerencial_html, height=1200, scrolling=True)
    with tab2:
        components.html(atrasos_html, height=1200, scrolling=True)
else:
    st.info("Ative 'Exibir HTMLs completos na tela' na barra lateral para visualizar os BIs dentro do Streamlit.")
