"""
App Streamlit: preenchimento automático da planilha padrão de análise por
máquina a partir dos arquivos baixados dos analisadores de energia.
"""
import html
from datetime import date

import streamlit as st

from preenchimento import (
    ErroLeitura,
    aplicar_plano,
    aplicar_plano_antes_depois,
    carregar_planilha_padrao,
    data_id1_atual,
    equipamento_esperado,
    ler_arquivo_analisador,
    montar_plano,
    montar_plano_antes_depois,
    salvar_para_bytes,
)

st.set_page_config(
    page_title="Preenchimento · Planilha padrão",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ----------------------------------------------------------------------------
# Estilo
# ----------------------------------------------------------------------------
st.markdown(
    """
    <link rel="preconnect" href="https://fonts.googleapis.com">
    <link href="https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@500;600;700&family=IBM+Plex+Sans:wght@400;500;600&family=IBM+Plex+Mono:wght@400;500&display=swap" rel="stylesheet">
    <style>
        html, body, [class*="css"]  { font-family: 'IBM Plex Sans', sans-serif; }
        h1, h2, h3 { font-family: 'Space Grotesk', sans-serif; letter-spacing: -0.01em; }

        .bloco-topo {
            display: flex; align-items: center; gap: 0.9rem;
            padding-bottom: 0.6rem; border-bottom: 1px solid #DEDCD4;
            margin-bottom: 1.4rem;
        }
        .bloco-topo .icone {
            font-size: 2rem; line-height: 1;
        }
        .bloco-topo h1 {
            font-size: 1.6rem; margin: 0; color: #1B2028;
        }
        .bloco-topo p {
            margin: 0.1rem 0 0 0; color: #5B6270; font-size: 0.95rem;
        }

        .secao-titulo {
            font-family: 'Space Grotesk', sans-serif; font-weight: 600;
            font-size: 1.05rem; color: #1B2028; margin: 1.6rem 0 0.4rem 0;
            display: flex; align-items: center; gap: 0.5rem;
        }
        .secao-titulo .marcador {
            width: 1.6rem; height: 1.6rem; border-radius: 50%;
            background: #1B2028; color: #F5F6F4; font-size: 0.85rem;
            display: inline-flex; align-items: center; justify-content: center;
            font-family: 'IBM Plex Mono', monospace; flex-shrink: 0;
        }
        .secao-sub { color: #5B6270; font-size: 0.88rem; margin-bottom: 0.7rem; }

        .painel {
            background: #FFFFFF; border: 1px solid #E4E2D9; border-radius: 6px;
            padding: 1rem 1.2rem; margin-bottom: 0.8rem;
        }
        .painel h4 {
            margin: 0 0 0.6rem 0; font-family: 'Space Grotesk', sans-serif;
            font-size: 0.95rem; color: #1B2028;
        }

        .pilula {
            display: inline-block; padding: 0.15rem 0.55rem; border-radius: 3px;
            font-size: 0.78rem; font-family: 'IBM Plex Mono', monospace;
            font-weight: 500;
        }
        .pilula-novo { background: #E3F1EF; color: #14524F; }
        .pilula-substituicao { background: #FBEBDA; color: #8A4B0F; }
        .pilula-fora { background: #FBE3E1; color: #8C231C; }
        .pilula-ok { background: #E3F1EF; color: #14524F; }
        .pilula-erro { background: #FBE3E1; color: #8C231C; }
        .pilula-antes { background: #EFE7F8; color: #4B2E83; }
        .pilula-depois { background: #E3F1EF; color: #14524F; }

        table.tabela-plano {
            width: 100%; border-collapse: collapse; font-size: 0.87rem;
        }
        table.tabela-plano th {
            text-align: left; padding: 0.5rem 0.6rem; color: #5B6270;
            font-weight: 500; border-bottom: 1px solid #E4E2D9;
        }
        table.tabela-plano td {
            padding: 0.5rem 0.6rem; border-bottom: 1px solid #F0EFE9;
            vertical-align: top;
        }
        table.tabela-plano td.mono { font-family: 'IBM Plex Mono', monospace; }
        table.tabela-plano tr:last-child td { border-bottom: none; }
        .aviso-texto { color: #8A4B0F; font-size: 0.82rem; }

        div[data-testid="stMetric"] {
            background: #FFFFFF; border: 1px solid #E4E2D9; border-radius: 6px;
            padding: 0.8rem 1rem;
        }

        .stButton>button[kind="primary"] {
            background-color: #BD6B1A; border-color: #BD6B1A;
        }
        .stButton>button[kind="primary"]:hover {
            background-color: #A25A13; border-color: #A25A13;
        }
    </style>
    """,
    unsafe_allow_html=True,
)

st.markdown(
    """
    <div class="bloco-topo">
        <div class="icone">⚡</div>
        <div>
            <h1>Preenchimento da planilha padrão</h1>
            <p>Suba o modelo e os arquivos dos analisadores — a ferramenta descobre sozinha
            em qual aba cada dia entra e evita misturar dados de períodos diferentes.</p>
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)


# ----------------------------------------------------------------------------
# Estado
# ----------------------------------------------------------------------------
def _estado_inicial():
    return dict(
        wb=None,
        nome_planilha=None,
        equipamento=None,
        data_id1=None,
        registros=[],
        erros_leitura=[],
        plano=None,
        resultado_bytes=None,
        nome_saida=None,
        # modo Antes e Depois
        registros_antes=[],
        registros_depois=[],
        erros_antes=[],
        erros_depois=[],
        plano_ad=None,
        resumo_ad=None,
    )


if "s" not in st.session_state:
    st.session_state.s = _estado_inicial()
s = st.session_state.s

with st.sidebar:
    st.markdown("### Sobre a ferramenta")
    st.markdown(
        "Cada aba numerada da planilha padrão (1 a 49) guarda um dia de "
        "medição. No modo **Campanha única**, a data de cada dia vem da "
        "aba **1**. No modo **Antes e Depois**, os dias 'antes' sempre "
        "ocupam as abas a partir da 1 (bloco M1) e os 'depois' entram logo "
        "em seguida (bloco M2), inclusive gerando a aba **Comparativo** "
        "com a economia estimada."
    )
    with st.expander("Detalhes técnicos"):
        st.markdown(
            "- Só escreve nas colunas **A:M** das abas numeradas.\n"
            "- No modo Antes e Depois, também ajusta o Índice de Medições, "
            "os blocos de consumo (M1/M2) e os gráficos de perfil semanal.\n"
            "- Nunca mexe nas abas M3-M7 nem nas demais abas numeradas.\n"
            "- Corrige a célula A1 de cada aba tocada (usada pelos botões "
            "de navegação da planilha).\n"
        )
    st.divider()
    if st.button("Começar de novo", use_container_width=True):
        st.session_state.s = _estado_inicial()
        st.rerun()

# ----------------------------------------------------------------------------
# 1. Planilha padrão
# ----------------------------------------------------------------------------
st.markdown(
    '<div class="secao-titulo"><span class="marcador">1</span>Planilha padrão (.xlsm)</div>',
    unsafe_allow_html=True,
)
st.markdown(
    '<div class="secao-sub">O arquivo modelo que você quer preencher.</div>',
    unsafe_allow_html=True,
)

arquivo_planilha = st.file_uploader(
    "Planilha padrão", type=["xlsm"], label_visibility="collapsed", key="up_planilha"
)

if arquivo_planilha is not None and arquivo_planilha.name != s.get("nome_planilha"):
    try:
        wb = carregar_planilha_padrao(arquivo_planilha.getvalue())
        if "1" not in wb.sheetnames or "Resumo" not in wb.sheetnames:
            st.error(
                "Esse arquivo não parece ter a estrutura esperada da planilha "
                "padrão (abas 'Resumo' e '1' não encontradas)."
            )
        else:
            s["wb"] = wb
            s["nome_planilha"] = arquivo_planilha.name
            s["equipamento"] = equipamento_esperado(wb)
            s["data_id1"] = data_id1_atual(wb)
            s["plano"] = None
            s["resultado_bytes"] = None
            s["plano_ad"] = None
            s["resumo_ad"] = None
    except Exception as e:
        st.error(f"Não consegui abrir essa planilha: {e}")

if s["wb"] is not None:
    col1, col2, col3 = st.columns(3)
    col1.metric("Equipamento cadastrado", s["equipamento"] or "não identificado")
    col2.metric(
        "Data do ID 1",
        s["data_id1"].strftime("%d/%m/%Y") if s["data_id1"] else "aba vazia",
    )
    n_ids_preenchidos = sum(
        1 for i in range(1, 50)
        if str(i) in s["wb"].sheetnames and s["wb"][str(i)]["A6"].value
    )
    col3.metric("Abas já preenchidas", f"{n_ids_preenchidos} / 49")

# ----------------------------------------------------------------------------
# Seletor de modo
# ----------------------------------------------------------------------------
if s["wb"] is not None:
    modo = st.radio(
        "Modo de preenchimento",
        options=["Campanha única", "Antes e Depois"],
        horizontal=True,
        help=(
            "Campanha única: preenche as abas a partir da data já cadastrada "
            "na aba 1. Antes e Depois: monta o comparativo M1 (antes) x M2 "
            "(depois), com a aba Comparativo e a economia estimada."
        ),
    )
else:
    modo = None

# ============================================================================
# MODO: CAMPANHA ÚNICA
# ============================================================================
if modo == "Campanha única":
    if s["data_id1"] is None:
        st.warning(
            "A aba **1** ainda está vazia — informe manualmente a que dia ela "
            "vai corresponder antes de continuar."
        )
        data_manual = st.date_input(
            "Data do ID 1", value=date.today(), format="DD/MM/YYYY", key="data_manual"
        )
        s["data_id1"] = data_manual
    else:
        with st.expander("Usar outra data de referência para o ID 1"):
            usar_manual = st.checkbox("Substituir a data detectada")
            if usar_manual:
                s["data_id1"] = st.date_input(
                    "Nova data do ID 1", value=s["data_id1"], format="DD/MM/YYYY"
                )

    st.markdown(
        '<div class="secao-titulo"><span class="marcador">2</span>Arquivos dos analisadores (.xls)</div>',
        unsafe_allow_html=True,
    )
    st.markdown(
        '<div class="secao-sub">Pode soltar quantos arquivos quiser de uma vez — a ordem não importa.</div>',
        unsafe_allow_html=True,
    )

    arquivos_dados = st.file_uploader(
        "Arquivos dos analisadores",
        type=["xls"],
        accept_multiple_files=True,
        label_visibility="collapsed",
        key="up_dados",
    )

    if arquivos_dados:
        nomes_atuais = {f.name for f in arquivos_dados}
        nomes_lidos = {r.nome_arquivo for r in s["registros"]} | {e[0] for e in s["erros_leitura"]}
        if nomes_atuais != nomes_lidos:
            registros, erros = [], []
            for f in arquivos_dados:
                try:
                    registros.append(ler_arquivo_analisador(f.name, f.getvalue()))
                except ErroLeitura as e:
                    erros.append((f.name, str(e)))
            s["registros"] = registros
            s["erros_leitura"] = erros
            s["plano"] = None
            s["resultado_bytes"] = None

        linhas_html = []
        for r in sorted(s["registros"], key=lambda r: r.data):
            linhas_html.append(
                f"<tr><td>{html.escape(r.nome_arquivo)}</td>"
                f"<td class='mono'>{r.data.strftime('%d/%m/%Y')}</td>"
                f"<td>{html.escape(r.equipamento or '—')}</td>"
                f"<td class='mono'>{r.n_linhas}</td>"
                f"<td><span class='pilula pilula-ok'>ok</span></td></tr>"
            )
        for nome, erro in s["erros_leitura"]:
            linhas_html.append(
                f"<tr><td>{html.escape(nome)}</td><td colspan='3' class='aviso-texto'>"
                f"{html.escape(erro)}</td>"
                f"<td><span class='pilula pilula-erro'>erro</span></td></tr>"
            )

        st.markdown(
            f"""
            <div class="painel">
            <table class="tabela-plano">
            <tr><th>Arquivo</th><th>Data</th><th>Equipamento</th><th>Linhas</th><th>Status</th></tr>
            {''.join(linhas_html)}
            </table>
            </div>
            """,
            unsafe_allow_html=True,
        )

    pronto_para_plano = s["wb"] is not None and s["registros"] and s["data_id1"] is not None

    if pronto_para_plano:
        st.markdown(
            '<div class="secao-titulo"><span class="marcador">3</span>Conferir e gerar</div>',
            unsafe_allow_html=True,
        )

        plano = montar_plano(s["wb"], s["registros"], s["data_id1"])
        s["plano"] = plano

        n_novo = sum(1 for p in plano if p.status == "novo")
        n_subst = sum(1 for p in plano if p.status == "substituicao")
        n_fora = sum(1 for p in plano if p.status == "fora_do_intervalo")
        n_avisos = sum(1 for p in plano if p.aviso)

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Dias novos", n_novo)
        c2.metric("Substituições", n_subst)
        c3.metric("Fora do intervalo", n_fora)
        c4.metric("Avisos", n_avisos)

        classe_status = {
            "novo": ("novo", "novo"),
            "substituicao": ("substituição", "substituicao"),
            "fora_do_intervalo": ("fora do intervalo", "fora"),
        }

        linhas_plano = []
        for item in plano:
            rotulo, classe = classe_status[item.status]
            data_str = item.registro.data.strftime("%d/%m/%Y")
            aviso_html = f"<div class='aviso-texto'>{html.escape(item.aviso)}</div>" if item.aviso else ""
            linhas_plano.append(
                f"<tr><td class='mono'>{item.sheet_id}</td>"
                f"<td class='mono'>{data_str}</td>"
                f"<td>{html.escape(item.registro.nome_arquivo)}</td>"
                f"<td><span class='pilula pilula-{classe}'>{rotulo}</span></td>"
                f"<td>{html.escape(item.detalhe)}{aviso_html}</td></tr>"
            )

        st.markdown(
            f"""
            <div class="painel">
            <table class="tabela-plano">
            <tr><th>ID</th><th>Data</th><th>Arquivo</th><th>Situação</th><th>Detalhe</th></tr>
            {''.join(linhas_plano)}
            </table>
            </div>
            """,
            unsafe_allow_html=True,
        )

        if n_fora:
            st.warning(
                f"{n_fora} arquivo(s) ficariam fora do intervalo de abas (1-49) e "
                "não serão aplicados. Confira a data de referência do ID 1."
            )

        with st.expander("Limpar abas sem dado novo (opcional)"):
            st.caption(
                "Use para esvaziar abas que ainda têm dados de um período não "
                "relacionado a esta campanha, mesmo sem um arquivo novo pra elas."
            )
            texto_limpar = st.text_input(
                "IDs separados por vírgula", placeholder="ex.: 13, 14", key="ids_limpar"
            )

        ids_para_limpar = []
        if texto_limpar.strip():
            for parte in texto_limpar.split(","):
                parte = parte.strip()
                if parte.isdigit():
                    ids_para_limpar.append(int(parte))

        itens_aplicaveis = [p for p in plano if p.status != "fora_do_intervalo"]
        confirmar = st.checkbox(
            f"Conferi o plano acima — aplicar {len(itens_aplicaveis)} dia(s)"
            + (f" e limpar {len(ids_para_limpar)} aba(s)" if ids_para_limpar else "")
        )

        if st.button("Gerar planilha atualizada", type="primary", disabled=not confirmar):
            with st.spinner("Escrevendo os dados nas abas..."):
                aplicados = aplicar_plano(s["wb"], plano, ids_para_limpar)
                s["resultado_bytes"] = salvar_para_bytes(s["wb"])
                base = s["nome_planilha"].rsplit(".", 1)[0]
                s["nome_saida"] = f"{base}_ATUALIZADA.xlsm"
            st.success(f"Pronto — {len(aplicados)} aba(s) atualizada(s).")

    if s["resultado_bytes"] is not None:
        st.markdown(
            '<div class="secao-titulo"><span class="marcador">4</span>Baixar</div>',
            unsafe_allow_html=True,
        )
        st.download_button(
            "Baixar planilha atualizada (.xlsm)",
            data=s["resultado_bytes"],
            file_name=s["nome_saida"],
            mime="application/vnd.ms-excel.sheet.macroEnabled.12",
            type="primary",
        )
        st.caption(
            "As fórmulas (Resumo, M1-M7) recalculam automaticamente quando você "
            "abre o arquivo no Excel — não é preciso fazer nada além de abrir."
        )

# ============================================================================
# MODO: ANTES E DEPOIS
# ============================================================================
elif modo == "Antes e Depois":
    st.markdown(
        '<div class="secao-titulo"><span class="marcador">2</span>Dados "Antes" e "Depois" (.xls)</div>',
        unsafe_allow_html=True,
    )
    st.markdown(
        '<div class="secao-sub">Até 7 dias em cada grupo. O grupo Antes sempre ocupa o bloco M1 '
        '(a partir da aba 1); o Depois ocupa o bloco M2, logo em seguida.</div>',
        unsafe_allow_html=True,
    )

    col_a, col_b = st.columns(2)
    with col_a:
        st.markdown('<div class="painel"><h4>🟣 Antes</h4>', unsafe_allow_html=True)
        arquivos_antes = st.file_uploader(
            "Dados Antes", type=["xls"], accept_multiple_files=True,
            label_visibility="collapsed", key="up_antes",
        )
        st.markdown('</div>', unsafe_allow_html=True)
    with col_b:
        st.markdown('<div class="painel"><h4>🟢 Depois</h4>', unsafe_allow_html=True)
        arquivos_depois = st.file_uploader(
            "Dados Depois", type=["xls"], accept_multiple_files=True,
            label_visibility="collapsed", key="up_depois",
        )
        st.markdown('</div>', unsafe_allow_html=True)

    def _processar_grupo(arquivos, chave_reg, chave_erro):
        if not arquivos:
            return
        nomes_atuais = {f.name for f in arquivos}
        nomes_lidos = {r.nome_arquivo for r in s[chave_reg]} | {e[0] for e in s[chave_erro]}
        if nomes_atuais != nomes_lidos:
            registros, erros = [], []
            for f in arquivos:
                try:
                    registros.append(ler_arquivo_analisador(f.name, f.getvalue()))
                except ErroLeitura as e:
                    erros.append((f.name, str(e)))
            s[chave_reg] = registros
            s[chave_erro] = erros
            s["plano_ad"] = None
            s["resumo_ad"] = None

    _processar_grupo(arquivos_antes, "registros_antes", "erros_antes")
    _processar_grupo(arquivos_depois, "registros_depois", "erros_depois")

    def _tabela_grupo(registros, erros, pilula_classe):
        linhas_html = []
        for r in sorted(registros, key=lambda r: r.data):
            linhas_html.append(
                f"<tr><td>{html.escape(r.nome_arquivo)}</td>"
                f"<td class='mono'>{r.data.strftime('%d/%m/%Y')}</td>"
                f"<td class='mono'>{r.n_linhas}</td>"
                f"<td><span class='pilula pilula-{pilula_classe}'>ok</span></td></tr>"
            )
        for nome, erro in erros:
            linhas_html.append(
                f"<tr><td>{html.escape(nome)}</td><td colspan='2' class='aviso-texto'>"
                f"{html.escape(erro)}</td>"
                f"<td><span class='pilula pilula-erro'>erro</span></td></tr>"
            )
        if not linhas_html:
            return
        st.markdown(
            f"""
            <div class="painel">
            <table class="tabela-plano">
            <tr><th>Arquivo</th><th>Data</th><th>Linhas</th><th>Status</th></tr>
            {''.join(linhas_html)}
            </table>
            </div>
            """,
            unsafe_allow_html=True,
        )

    col_a2, col_b2 = st.columns(2)
    with col_a2:
        _tabela_grupo(s["registros_antes"], s["erros_antes"], "antes")
    with col_b2:
        _tabela_grupo(s["registros_depois"], s["erros_depois"], "depois")

    pronto_ad = (
        s["wb"] is not None and s["registros_antes"] and s["registros_depois"]
    )

    if pronto_ad:
        st.markdown(
            '<div class="secao-titulo"><span class="marcador">3</span>Conferir e gerar</div>',
            unsafe_allow_html=True,
        )

        plano_ad = montar_plano_antes_depois(s["wb"], s["registros_antes"], s["registros_depois"])
        s["plano_ad"] = plano_ad

        for aviso in plano_ad.avisos:
            st.warning(aviso)

        c1, c2 = st.columns(2)
        c1.metric("Dias 'Antes' (abas M1)", len(plano_ad.antes))
        c2.metric("Dias 'Depois' (abas M2)", len(plano_ad.depois))

        def _linhas_plano_ad(itens):
            linhas = []
            for item in itens:
                data_str = item.registro.data.strftime("%d/%m/%Y")
                aviso_html = (
                    f"<div class='aviso-texto'>{html.escape(item.aviso)}</div>"
                    if item.aviso else ""
                )
                linhas.append(
                    f"<tr><td class='mono'>{item.sheet_id}</td>"
                    f"<td class='mono'>{data_str}</td>"
                    f"<td>{html.escape(item.registro.nome_arquivo)}</td>"
                    f"<td>{aviso_html or '—'}</td></tr>"
                )
            return "".join(linhas)

        col_a3, col_b3 = st.columns(2)
        with col_a3:
            st.markdown(
                f"""
                <div class="painel"><h4>🟣 Antes → bloco M1</h4>
                <table class="tabela-plano">
                <tr><th>Aba</th><th>Data</th><th>Arquivo</th><th>Aviso</th></tr>
                {_linhas_plano_ad(plano_ad.antes)}
                </table>
                </div>
                """,
                unsafe_allow_html=True,
            )
        with col_b3:
            st.markdown(
                f"""
                <div class="painel"><h4>🟢 Depois → bloco M2</h4>
                <table class="tabela-plano">
                <tr><th>Aba</th><th>Data</th><th>Arquivo</th><th>Aviso</th></tr>
                {_linhas_plano_ad(plano_ad.depois)}
                </table>
                </div>
                """,
                unsafe_allow_html=True,
            )

        st.caption(
            "Ao gerar, a ferramenta também corrige o Índice de Medições, os "
            "blocos de consumo de energia (M1/M2), o perfil semanal e cria/"
            "atualiza a aba **Comparativo** com a economia estimada."
        )

        confirmar_ad = st.checkbox(
            f"Conferi o plano acima — aplicar {len(plano_ad.antes)} dia(s) de Antes "
            f"e {len(plano_ad.depois)} dia(s) de Depois"
        )

        if st.button("Gerar planilha Antes x Depois", type="primary", disabled=not confirmar_ad):
            with st.spinner("Escrevendo os dados, ajustando o Resumo e montando o Comparativo..."):
                try:
                    resumo_ad = aplicar_plano_antes_depois(s["wb"], plano_ad)
                    s["resumo_ad"] = resumo_ad
                    s["resultado_bytes"] = salvar_para_bytes(s["wb"])
                    base = s["nome_planilha"].rsplit(".", 1)[0]
                    s["nome_saida"] = f"{base}_ANTES_DEPOIS.xlsm"
                    st.success(
                        f"Pronto — {resumo_ad['n_antes']} dia(s) de Antes "
                        f"(abas {resumo_ad['id_antes'][0]}-{resumo_ad['id_antes'][1]}) e "
                        f"{resumo_ad['n_depois']} dia(s) de Depois "
                        f"(abas {resumo_ad['id_depois'][0]}-{resumo_ad['id_depois'][1]})."
                    )
                except ValueError as e:
                    st.error(str(e))

    if s["resultado_bytes"] is not None and s.get("resumo_ad") is not None:
        st.markdown(
            '<div class="secao-titulo"><span class="marcador">4</span>Baixar</div>',
            unsafe_allow_html=True,
        )
        st.download_button(
            "Baixar planilha Antes x Depois (.xlsm)",
            data=s["resultado_bytes"],
            file_name=s["nome_saida"],
            mime="application/vnd.ms-excel.sheet.macroEnabled.12",
            type="primary",
        )
        st.caption(
            "As fórmulas (Resumo, M1/M2, Comparativo) recalculam automaticamente "
            "quando você abre o arquivo no Excel — não é preciso fazer nada além "
            "de abrir."
        )
