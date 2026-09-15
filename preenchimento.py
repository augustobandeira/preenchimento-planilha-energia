"""
Núcleo de lógica do preenchimento automático da planilha padrão de análise
por máquina. Sem nenhuma dependência do Streamlit — só openpyxl/xlrd — para
poder ser testado e reaproveitado fora do app também.

Estrutura da planilha padrão (ver LEIAME do projeto para mais detalhes):
- Abas numeradas "1" a "49": uma por dia de medição, dados minuto a minuto
  colados em A:M a partir da linha 6.
- "Resumo" e "M1".."M7": só leem as abas numeradas via fórmula.
"""
from __future__ import annotations

import io
from dataclasses import dataclass, field
from datetime import date, datetime
from typing import Optional

import openpyxl
import xlrd

N_COLS_DADOS = 13
LINHA_DADOS_INICIO = 6
LINHA_DADOS_FIM = 1445
ID_MIN, ID_MAX = 1, 49
CAPACIDADE_ABA = LINHA_DADOS_FIM - LINHA_DADOS_INICIO + 1

COLUNAS_ESPERADAS = [
    "Horário", "UrmsA", "UrmsB", "UrmsC", "IrmsA", "IrmsB", "IrmsC",
    "PT", "QT", "Temperatura A", "Temperatura B", "PTG", "QTG",
]


@dataclass
class RegistroDiario:
    """Um arquivo .xls do analisador já lido e interpretado."""
    nome_arquivo: str
    linhas: list
    data: date
    equipamento: Optional[str] = None
    primeiro_horario: Optional[str] = None
    ultimo_horario: Optional[str] = None

    @property
    def n_linhas(self) -> int:
        return len(self.linhas)


@dataclass
class ItemPlano:
    """Uma linha do plano de preenchimento: o que vai acontecer com um ID."""
    sheet_id: int
    registro: Optional[RegistroDiario]
    status: str          # "novo" | "substituicao" | "fora_do_intervalo" | "limpeza"
    detalhe: str = ""
    aviso: Optional[str] = None


class ErroLeitura(Exception):
    pass


def ler_arquivo_analisador(nome_arquivo: str, conteudo_bytes: bytes) -> RegistroDiario:
    """Lê um .xls baixado do analisador (bytes em memória) e devolve um
    RegistroDiario. Não depende de o arquivo ter sempre as mesmas linhas de
    metadados no topo — procura a linha de cabeçalho "Horário" dinamicamente.
    """
    try:
        book = xlrd.open_workbook(file_contents=conteudo_bytes)
    except Exception as e:
        raise ErroLeitura(f"Não consegui abrir '{nome_arquivo}' como .xls: {e}")

    sh = book.sheet_by_index(0)

    header_row = None
    for r in range(min(sh.nrows, 20)):
        v = sh.cell(r, 0).value
        if isinstance(v, str) and v.strip().lower() == "horário":
            header_row = r
            break
    if header_row is None:
        raise ErroLeitura(
            f"'{nome_arquivo}' não parece um arquivo do analisador — não encontrei "
            f"a linha de cabeçalho 'Horário'."
        )

    # A linha "Equipamento ..." fica normalmente 3 linhas acima do cabeçalho
    # (acima dela vêm "Início:" e "Fim:"). Confirma pelo prefixo antes de
    # aceitar, e cai para uma busca genérica se o layout for diferente.
    equipamento = None
    candidata = sh.cell(header_row - 3, 0).value if header_row >= 3 else None
    if isinstance(candidata, str) and candidata.strip().lower().startswith("equipamento"):
        equipamento = candidata.strip()
    else:
        for r in range(header_row):
            v = sh.cell(r, 0).value
            if isinstance(v, str) and v.strip().lower().startswith("equipamento"):
                equipamento = v.strip()
                break

    data_rows = []
    for r in range(header_row + 1, sh.nrows):
        primeira_cel = sh.cell(r, 0).value
        if not primeira_cel:
            continue
        linha = [sh.cell(r, c).value for c in range(min(N_COLS_DADOS, sh.ncols))]
        while len(linha) < N_COLS_DADOS:
            linha.append(None)
        data_rows.append(linha)

    if not data_rows:
        raise ErroLeitura(f"'{nome_arquivo}' não tem nenhuma linha de dados.")

    try:
        primeira_dt = datetime.strptime(str(data_rows[0][0]).strip(), "%d/%m/%Y %H:%M")
    except ValueError:
        raise ErroLeitura(
            f"Não consegui interpretar a data '{data_rows[0][0]}' em '{nome_arquivo}'."
        )

    return RegistroDiario(
        nome_arquivo=nome_arquivo,
        linhas=data_rows,
        data=primeira_dt.date(),
        equipamento=equipamento,
        primeiro_horario=str(data_rows[0][0]),
        ultimo_horario=str(data_rows[-1][0]),
    )


def carregar_planilha_padrao(conteudo_bytes: bytes):
    """Abre a planilha padrão (.xlsm) a partir de bytes em memória."""
    return openpyxl.load_workbook(io.BytesIO(conteudo_bytes), keep_vba=True, data_only=False)


def data_id1_atual(wb) -> Optional[date]:
    """Data já cadastrada na aba '1' (A6), se houver."""
    if "1" not in wb.sheetnames:
        return None
    a6 = wb["1"]["A6"].value
    if not a6:
        return None
    try:
        return datetime.strptime(str(a6).strip(), "%d/%m/%Y %H:%M").date()
    except ValueError:
        return None


def equipamento_esperado(wb) -> Optional[str]:
    """Nome do equipamento já cadastrado no Índice de Medições (Resumo!F4)."""
    if "Resumo" not in wb.sheetnames:
        return None
    for r in range(4, 11):
        v = wb["Resumo"].cell(row=r, column=6).value  # coluna F
        if v:
            return str(v).strip()
    return None


def equipamento_compativel(esperado: Optional[str], do_arquivo: Optional[str]) -> bool:
    """Compara de forma tolerante: o código curto do Índice de Medições
    (ex.: 'RE30 436') normalmente aparece como substring dentro da descrição
    completa do arquivo do analisador (ex.: 'Equipamento NS: 69000436 | ...
    Modelo: RE30'). Considera compatível se todos os tokens alfanuméricos do
    código curto aparecem em algum lugar da descrição do arquivo."""
    if not esperado or not do_arquivo:
        return True  # não dá pra comparar — não bloqueia, só não avisa
    alvo = do_arquivo.upper()
    tokens = [t for t in esperado.upper().replace("-", " ").split() if t.isalnum()]
    if not tokens:
        return True
    return all(t in alvo for t in tokens)


def data_do_id(wb, sheet_id: int) -> Optional[date]:
    """Data já preenchida numa aba numerada (A6), se houver."""
    nome = str(sheet_id)
    if nome not in wb.sheetnames:
        return None
    a6 = wb[nome]["A6"].value
    if not a6:
        return None
    try:
        return datetime.strptime(str(a6).strip(), "%d/%m/%Y %H:%M").date()
    except ValueError:
        return None


def montar_plano(wb, registros: list[RegistroDiario], data_id1: date) -> list[ItemPlano]:
    """Calcula, para cada arquivo lido, em qual aba (ID) ele deve entrar, e
    classifica a situação (novo / substituição / fora do intervalo)."""
    esperado = equipamento_esperado(wb)
    plano = []
    for reg in sorted(registros, key=lambda r: r.data):
        offset = (reg.data - data_id1).days
        sheet_id = ID_MIN + offset
        if sheet_id < ID_MIN or sheet_id > ID_MAX:
            plano.append(ItemPlano(
                sheet_id=sheet_id,
                registro=reg,
                status="fora_do_intervalo",
                detalhe=f"Cairia no ID {sheet_id}, fora do intervalo {ID_MIN}-{ID_MAX}.",
            ))
            continue

        data_existente = data_do_id(wb, sheet_id)
        if data_existente is None:
            status = "novo"
            detalhe = "Aba vazia."
        elif data_existente == reg.data:
            status = "substituicao"
            detalhe = "Já tinha dados desse mesmo dia — serão atualizados."
        else:
            status = "substituicao"
            detalhe = (
                f"Aba tinha dados de {data_existente.strftime('%d/%m/%Y')} "
                f"(período diferente) — serão substituídos."
            )

        avisos = []
        if reg.n_linhas > CAPACIDADE_ABA:
            avisos.append(
                f"Arquivo tem {reg.n_linhas} linhas, mais que a capacidade da aba "
                f"({CAPACIDADE_ABA}). O excedente será cortado."
            )
        if not equipamento_compativel(esperado, reg.equipamento):
            avisos.append(
                f"O equipamento deste arquivo ('{reg.equipamento}') não parece bater "
                f"com o esperado ('{esperado}') — confira se não é de outra máquina."
            )
        aviso = " ".join(avisos) if avisos else None

        plano.append(ItemPlano(sheet_id=sheet_id, registro=reg, status=status,
                                detalhe=detalhe, aviso=aviso))
    return plano


def _limpar_linhas_excedentes(ws, primeira_linha_livre: int):
    for r in range(primeira_linha_livre, LINHA_DADOS_FIM + 1):
        for c in range(1, N_COLS_DADOS + 1):
            ws.cell(row=r, column=c).value = None


def aplicar_plano(wb, plano: list[ItemPlano], ids_para_limpar: Optional[list[int]] = None):
    """Escreve os dados no workbook conforme o plano. Só mexe nas colunas
    A:M das abas envolvidas — nunca toca em fórmulas ou em outras abas."""
    aplicados = []

    for item in plano:
        if item.status == "fora_do_intervalo":
            continue
        nome_aba = str(item.sheet_id)
        if nome_aba not in wb.sheetnames:
            continue
        ws = wb[nome_aba]
        linhas = item.registro.linhas[:CAPACIDADE_ABA]

        for i, linha in enumerate(linhas):
            r = LINHA_DADOS_INICIO + i
            for c, valor in enumerate(linha, start=1):
                ws.cell(row=r, column=c).value = valor

        _limpar_linhas_excedentes(ws, LINHA_DADOS_INICIO + len(linhas))
        ws["A1"].value = item.sheet_id
        aplicados.append(item)

    for sheet_id in (ids_para_limpar or []):
        nome_aba = str(sheet_id)
        if nome_aba not in wb.sheetnames:
            continue
        ws = wb[nome_aba]
        _limpar_linhas_excedentes(ws, LINHA_DADOS_INICIO)
        ws["A1"].value = sheet_id

    return aplicados


def salvar_para_bytes(wb) -> bytes:
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ============================================================================
# Modo "Antes e Depois" — planilha padrão v2, com abas M1 (Antes) / M2 (Depois)
# e aba "Comparativo". Cada bloco (M1, M2) suporta até 7 dias, assim como o
# restante do template (M3-M7).
# ============================================================================
MAX_DIAS_BLOCO = 7

_DIA_SEMANA_FORMULA = (
    '=IF(WEEKDAY(B{r})=1,"Domingo",IF(WEEKDAY(B{r})=2,"Segunda-Feira",'
    'IF(WEEKDAY(B{r})=3,"Terça-Feira",IF(WEEKDAY(B{r})=4,"Quarta-Feira",'
    'IF(WEEKDAY(B{r})=5,"Quinta-Feira",IF(WEEKDAY(B{r})=6,"Sexta-Feira",'
    'IF(WEEKDAY(B{r})=7,"Sábado","Erro")))))))'
)


@dataclass
class PlanoAntesDepois:
    antes: list          # list[ItemPlano], sheet_id 1..n_antes
    depois: list         # list[ItemPlano], sheet_id (n_antes+1)..(n_antes+n_depois)
    avisos: list = field(default_factory=list)


def montar_plano_antes_depois(wb, registros_antes: list, registros_depois: list) -> PlanoAntesDepois:
    """Monta o plano para o modo Antes x Depois: Antes sempre ocupa as abas
    1..n_antes (bloco M1), Depois ocupa as abas seguintes, n_antes+1..n_antes+n_depois
    (bloco M2). Cada bloco aceita no máximo 7 dias (mesmo limite dos blocos
    M1-M7 do template)."""
    avisos = []
    antes_ordenados = sorted(registros_antes, key=lambda r: r.data)
    depois_ordenados = sorted(registros_depois, key=lambda r: r.data)

    if len(antes_ordenados) > MAX_DIAS_BLOCO:
        avisos.append(
            f"Você enviou {len(antes_ordenados)} arquivo(s) para 'Antes', mas cada "
            f"bloco (M1) aceita no máximo {MAX_DIAS_BLOCO} dias — os "
            f"{len(antes_ordenados) - MAX_DIAS_BLOCO} mais recentes serão ignorados."
        )
        antes_ordenados = antes_ordenados[:MAX_DIAS_BLOCO]
    if len(depois_ordenados) > MAX_DIAS_BLOCO:
        avisos.append(
            f"Você enviou {len(depois_ordenados)} arquivo(s) para 'Depois', mas cada "
            f"bloco (M2) aceita no máximo {MAX_DIAS_BLOCO} dias — os "
            f"{len(depois_ordenados) - MAX_DIAS_BLOCO} mais recentes serão ignorados."
        )
        depois_ordenados = depois_ordenados[:MAX_DIAS_BLOCO]

    esperado = equipamento_esperado(wb)
    plano_antes, plano_depois = [], []
    for i, reg in enumerate(antes_ordenados, start=1):
        aviso = None
        if not equipamento_compativel(esperado, reg.equipamento):
            aviso = (
                f"O equipamento deste arquivo ('{reg.equipamento}') não parece "
                f"bater com o esperado ('{esperado}')."
            )
        plano_antes.append(ItemPlano(sheet_id=i, registro=reg, status="novo", aviso=aviso))

    id_ini_depois = len(antes_ordenados) + 1
    for offset, reg in enumerate(depois_ordenados):
        i = id_ini_depois + offset
        aviso = None
        if not equipamento_compativel(esperado, reg.equipamento):
            aviso = (
                f"O equipamento deste arquivo ('{reg.equipamento}') não parece "
                f"bater com o esperado ('{esperado}')."
            )
        plano_depois.append(ItemPlano(sheet_id=i, registro=reg, status="novo", aviso=aviso))

    return PlanoAntesDepois(antes=plano_antes, depois=plano_depois, avisos=avisos)


def _escrever_dia(wb, sheet_id: int, reg: RegistroDiario):
    ws = wb[str(sheet_id)]
    linhas = reg.linhas[:CAPACIDADE_ABA]
    for i, linha in enumerate(linhas):
        r = LINHA_DADOS_INICIO + i
        for c, valor in enumerate(linha, start=1):
            ws.cell(row=r, column=c).value = valor
    for r in range(LINHA_DADOS_INICIO + len(linhas), LINHA_DADOS_FIM + 1):
        for c in range(1, N_COLS_DADOS + 1):
            ws.cell(row=r, column=c).value = None
    ws["A1"].value = sheet_id


def _atualizar_bloco_indice_e_consumo(ws, linha_indice: int, linha_bloco_ini: int,
                                       rotulo: str, id_ini: int, n_dias: int,
                                       equipamento: Optional[str] = None):
    """Atualiza uma linha do Índice de Medições (4=M1, 5=M2, ...) e o bloco
    'Consumo de energia elétrica' correspondente (linhas linha_bloco_ini até
    linha_bloco_ini+6), ajustando para o número real de dias (1 a 7)."""
    id_fim = id_ini + n_dias - 1
    col_G, col_H = f"G{linha_indice}", f"H{linha_indice}"
    ws[f"B{linha_indice}"] = rotulo
    ws[f"C{linha_indice}"] = f"=TEXT('{id_ini}'!$A$6,\"dd/mm/aa\")"
    ws[f"D{linha_indice}"] = f"=TEXT('{id_fim}'!$A$6,\"dd/mm/aa\")"
    if equipamento:
        ws[f"F{linha_indice}"] = equipamento
    if isinstance(ws[col_G].value, (int, float)) or ws[col_G].value is None:
        ws[col_G] = id_ini
    if isinstance(ws[col_H].value, (int, float)) or ws[col_H].value is None:
        ws[col_H] = id_fim
    # quando G/H já são fórmulas (referenciando o bloco anterior), preserva-as
    # e só ajusta H para refletir o novo n_dias, mantendo a referência.
    if isinstance(ws[col_H].value, str) and ws[col_H].value.startswith("="):
        base = ws[col_H].value.split("+")[0]
        ws[col_H] = f"{base}+{n_dias}"

    linha_dados_ini = linha_bloco_ini + 2  # cabeçalho ocupa 2 linhas
    for i in range(MAX_DIAS_BLOCO):
        r = linha_dados_ini + i
        if i < n_dias:
            sheet_num = id_ini + i
            ws[f"B{r}"] = f"=TEXT('{sheet_num}'!$A$6,\"dd/mm/aa\")"
            ws[f"C{r}"] = _DIA_SEMANA_FORMULA.format(r=r)
            ws[f"D{r}"] = f"='{sheet_num}'!$U$1"
            ws[f"E{r}"] = f"='{sheet_num}'!$U$3"
            ws[f"F{r}"] = f"='{sheet_num}'!$U$2"
            ws[f"H{r}"] = sheet_num
            ws[f"I{r}"] = f"=WEEKDAY(B{r})"
            ws[f"J{r}"] = f'=IF(AND(I{r}>1,I{r}<7),F{r},"")'
            ws[f"K{r}"] = f'=IF(OR(I{r}>6,I{r}=1),F{r},"")'
        else:
            for col in "BCDEFHIJK":
                ws[f"{col}{r}"] = None

    ultima_linha_dados = linha_dados_ini + n_dias - 1
    linha_resumo = linha_dados_ini + MAX_DIAS_BLOCO + 1  # 2 linhas em branco após o bloco de 7
    j_range = f"J{linha_dados_ini}:J{ultima_linha_dados}"
    k_range = f"K{linha_dados_ini}:K{ultima_linha_dados}"
    ws[f"D{linha_resumo}"] = f"=AVERAGE({j_range})"
    ws[f"D{linha_resumo + 1}"] = f"=SUM({k_range})"
    ws[f"D{linha_resumo + 2}"] = (
        f"=(D{linha_resumo}*5)*4+D{linha_resumo}*2+4*SUM({k_range})"
    )
    # D(+3)=mensal->custo, D(+4)=anual, D(+5)=anual->custo já referenciam
    # D(linha_resumo+2) e não precisam mudar.
    return id_fim


def _formula_dia_perfil(sheet_num: int, row: int) -> str:
    lookup = f"VLOOKUP(ROUND($B{row},8),'{sheet_num}'!$N$6:$Z$1445,8,FALSE)"
    return f'=IF(ISERROR({lookup}),"",{lookup})'


ROWS_POR_DIA = 1440  # cada bloco de 1440 linhas = 1 dia (coluna B reinicia a cada bloco)


def _preencher_perfil_semanal(wb, ws_nome: str, id_ini: int, n_dias: int):
    """Preenche as colunas D..J (perfil minuto-a-minuto) e a coluna K (perfil
    médio dos dias reais) da aba M1/M2.

    A coluna B reinicia a cada 1440 linhas (um bloco de 1440 linhas = um dia).
    Cada coluna de dia (D=dia1, E=dia2, ...) só deve ter valor dentro do SEU
    bloco de 1440 linhas — fora dele, fica em branco. Preencher a coluna
    inteira (10080 linhas) para todo dia, como numa versão anterior, faz o
    Excel sobrepor as 6-7 séries em cada bloco, dando a aparência de dados
    "cruzados" no gráfico.

    A coluna K é uma média separada: para cada um dos 1440 minutos de um dia
    "típico", faz a média do mesmo horário nos n_dias dias reais — não é uma
    média de D:J na mesma linha, já que agora D:J nunca têm mais de um valor
    não-vazio na mesma linha.
    """
    ws = wb[ws_nome]
    colunas = ["D", "E", "F", "G", "H", "I", "J"]

    # Coluna C é um helper legado do template original (usado só pela coluna D
    # na versão de fábrica), com fórmula fixa referenciando abas fixas 8-14.
    # Reescreve sempre apontando para id_ini (aba garantida existir) para não
    # deixar referências a abas fora do intervalo criado.
    for row in range(2, 10082):
        ws[f"C{row}"] = f"=VLOOKUP(ROUND(B{row},8),'{id_ini}'!$N$6:$Z$1445,8,FALSE)"

    for i, col in enumerate(colunas):
        bloco_ini = 2 + i * ROWS_POR_DIA
        bloco_fim = 1 + (i + 1) * ROWS_POR_DIA
        if i < n_dias:
            sheet_num = id_ini + i
            for row in range(bloco_ini, bloco_fim + 1):
                ws[f"{col}{row}"] = _formula_dia_perfil(sheet_num, row)
        for row in range(2, 10082):
            if row < bloco_ini or row > bloco_fim:
                ws[f"{col}{row}"] = None
        if i >= n_dias:
            for row in range(2, 10082):
                ws[f"{col}{row}"] = None

    # Média por SOMA/CONTAGEM em vej de AVERAGE(): passar "" (texto literal)
    # como argumento direto do AVERAGE trava o cálculo com #VALUE! assim que
    # um dos dias não tem dado naquele minuto (o que é a maioria dos casos,
    # já que cada dia real cobre só uma parte das 24h) — SUM/COUNT ignora
    # naturalmente os dias sem dado nesse horário.
    sheets_reais = [str(id_ini + i) for i in range(n_dias)]
    for row in range(2, 1442):
        vlookups = [
            f"VLOOKUP(ROUND($B{row},8),'{s}'!$N$6:$Z$1445,8,FALSE)" for s in sheets_reais
        ]
        valores = "+".join(f"IF(ISERROR({v}),0,{v})" for v in vlookups)
        contagem = "+".join(f"IF(ISERROR({v}),0,1)" for v in vlookups)
        ws[f"K{row}"] = f'=IF(({contagem})=0,"",({valores})/({contagem}))'
    for row in range(1442, 10082):
        ws[f"K{row}"] = None
    ws["K1"] = "Média"


def _ajustar_grafico_perfil(ws_resumo, nome_aba_dados: str, n_dias: int):
    """Encontra o gráfico 'Perfil de consumo semanal' cuja 1a série aponta
    para a aba indicada (M1 ou M2) e corta as séries extras (dias que não
    existem nesta campanha)."""
    alvo = f"'{nome_aba_dados}'!"
    for ch in ws_resumo._charts:
        if not ch.series:
            continue
        val = getattr(ch.series[0].val, "numRef", None)
        f = val.f if val else ""
        if f.startswith(alvo):
            ch.series = ch.series[:max(n_dias, 1)]
            return True
    return False


def _criar_aba_comparativo(wb):
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.chart import BarChart, LineChart, Reference

    if "Comparativo" in wb.sheetnames:
        del wb["Comparativo"]
    idx = 1 if "Média" in wb.sheetnames else 0
    ws = wb.create_sheet("Comparativo", idx)

    TITULO = Font(name="Arial", size=16, bold=True, color="1F1F1F")
    SUB = Font(name="Arial", size=10, italic=True, color="595959")
    HEADER = Font(name="Arial", size=10, bold=True, color="FFFFFF")
    HEADER_FILL = PatternFill("solid", fgColor="BD6B1A")
    LABEL = Font(name="Arial", size=10)
    VALOR = Font(name="Arial", size=10)
    DESTAQUE_FONT = Font(name="Arial", size=20, bold=True, color="1F7A1F")
    DESTAQUE_LABEL = Font(name="Arial", size=11, bold=True)
    THIN = Side(style="thin", color="D9D9D9")
    BORDA = Border(top=THIN, bottom=THIN, left=THIN, right=THIN)

    ws.column_dimensions["A"].width = 2
    ws.column_dimensions["B"].width = 34
    for col in "CDEF":
        ws.column_dimensions[col].width = 16

    ws["B1"] = "Comparativo Antes x Depois"
    ws["B1"].font = TITULO
    ws["B2"] = (
        '=_xlfn.CONCAT("Equipamento: ", Resumo!F4, "   |   Antes: ", Resumo!C4, '
        '" a ", Resumo!D4, "   |   Depois: ", Resumo!C5, " a ", Resumo!D5)'
    )
    ws["B2"].font = SUB

    headers = ["Métrica", "Antes", "Depois", "Redução", "Redução %"]
    for i, h in enumerate(headers):
        c = ws.cell(row=4, column=2 + i, value=h)
        c.font, c.fill = HEADER, HEADER_FILL
        c.alignment = Alignment(horizontal="center")
        c.border = BORDA

    linhas = [
        ("Potência média [W]", "=AVERAGE(Resumo!D15:D21)", "=AVERAGE(Resumo!D32:D38)", "0.0"),
        ("Consumo diário médio [kWh/dia]", "=AVERAGE(Resumo!F15:F21)", "=AVERAGE(Resumo!F32:F38)", "0.00"),
        ("Consumo mensal estimado [kWh/mês]", "=Resumo!D25", "=Resumo!D42", "0"),
        ("Custo mensal estimado [R$]", "=Resumo!D26", "=Resumo!D43", '"R$" #,##0.00'),
        ("Consumo anual estimado [kWh/ano]", "=Resumo!D27", "=Resumo!D44", "0"),
        ("Custo anual estimado [R$]", "=Resumo!D28", "=Resumo!D45", '"R$" #,##0.00'),
    ]
    r = 5
    for nome, f_antes, f_depois, numfmt in linhas:
        ws.cell(row=r, column=2, value=nome).font = LABEL
        c_ant = ws.cell(row=r, column=3, value=f_antes)
        c_dep = ws.cell(row=r, column=4, value=f_depois)
        c_red = ws.cell(row=r, column=5, value=f"=C{r}-D{r}")
        c_pct = ws.cell(row=r, column=6, value=f"=IFERROR(E{r}/C{r},0)")
        for cc in (c_ant, c_dep, c_red):
            cc.number_format = numfmt
            cc.font = VALOR
        c_pct.number_format = "0.0%"
        c_pct.font = VALOR
        for col in "BCDEF":
            ws[f"{col}{r}"].border = BORDA
        r += 1

    ws["B12"] = "Economia anual estimada"
    ws["B12"].font = DESTAQUE_LABEL
    ws["C12"] = "=E10"
    ws["C12"].number_format = '"R$" #,##0.00'
    ws["C12"].font = DESTAQUE_FONT
    ws["D12"] = "=F10"
    ws["D12"].number_format = "0.0%"
    ws["D12"].font = DESTAQUE_FONT
    for cell in ("B12", "C12", "D12"):
        ws[cell].fill = PatternFill("solid", fgColor="E8F5E9")

    ws["B14"] = '=_xlfn.CONCAT("Custo de energia considerado: ", Resumo!O2, " ", Resumo!P2)'
    ws["B14"].font = SUB

    ws["B17"] = "Indicador"
    ws["C17"] = "Antes"
    ws["D17"] = "Depois"
    ws["B18"] = "Consumo mensal [kWh/mês]"
    ws["C18"] = "=C7"
    ws["D18"] = "=D7"
    ws["B19"] = "Custo mensal [R$]"
    ws["C19"] = "=C8"
    ws["D19"] = "=D8"
    ws["B20"] = "Potência média [W]"
    ws["C20"] = "=C5"
    ws["D20"] = "=D5"

    bar1 = BarChart()
    bar1.type, bar1.grouping = "col", "clustered"
    bar1.title = "Consumo e Custo Mensal: Antes x Depois"
    bar1.style = 10
    cats = Reference(ws, min_col=2, min_row=18, max_row=19)
    data = Reference(ws, min_col=3, max_col=4, min_row=17, max_row=19)
    bar1.add_data(data, titles_from_data=True)
    bar1.set_categories(cats)
    bar1.width, bar1.height = 16, 9
    ws.add_chart(bar1, "H4")

    bar2 = BarChart()
    bar2.type, bar2.grouping = "col", "clustered"
    bar2.title = "Potência Média: Antes x Depois"
    bar2.style = 11
    bar2.add_data(Reference(ws, min_col=3, max_col=4, min_row=20, max_row=20), titles_from_data=False)
    bar2.set_categories(Reference(ws, min_col=3, max_col=4, min_row=17, max_row=17))
    bar2.width, bar2.height = 16, 9
    ws.add_chart(bar2, "H20")

    if "M1" in wb.sheetnames and "M2" in wb.sheetnames:
        line = LineChart()
        line.title = "Perfil médio de potência: Antes x Depois"
        line.style = 12
        line.y_axis.title = "Potência [W]"
        line.x_axis.title = "Horário"
        line.x_axis.number_format = "hh:mm"
        cat_ref = Reference(wb["M1"], min_col=2, min_row=2, max_row=1441)
        val_antes = Reference(wb["M1"], min_col=11, min_row=1, max_row=1441)
        val_depois = Reference(wb["M2"], min_col=11, min_row=1, max_row=1441)
        wb["M1"]["K1"] = "Antes"
        wb["M2"]["K1"] = "Depois"
        line.add_data(val_antes, titles_from_data=True)
        line.add_data(val_depois, titles_from_data=True)
        line.set_categories(cat_ref)
        line.width, line.height = 24, 11
        ws.add_chart(line, "B37")


def aplicar_plano_antes_depois(wb, plano: PlanoAntesDepois):
    """Executa o plano completo do modo Antes x Depois: grava os dias nas
    abas numeradas, corrige o Índice de Medições e os blocos de consumo do
    Resumo (M1=Antes, M2=Depois), preenche o perfil semanal (M1/M2 + gráficos)
    e (re)cria a aba Comparativo."""
    n_antes, n_depois = len(plano.antes), len(plano.depois)
    if n_antes == 0 or n_depois == 0:
        raise ValueError("É preciso pelo menos 1 dia de 'Antes' e 1 de 'Depois'.")

    for item in plano.antes:
        _escrever_dia(wb, item.sheet_id, item.registro)
    for item in plano.depois:
        _escrever_dia(wb, item.sheet_id, item.registro)

    ws_resumo = wb["Resumo"]
    equipamento = equipamento_esperado(wb)
    _atualizar_bloco_indice_e_consumo(
        ws_resumo, linha_indice=4, linha_bloco_ini=13,
        rotulo="Compressor antes", id_ini=1, n_dias=n_antes, equipamento=equipamento,
    )
    id_ini_depois = n_antes + 1
    _atualizar_bloco_indice_e_consumo(
        ws_resumo, linha_indice=5, linha_bloco_ini=30,
        rotulo="Compressor depois", id_ini=id_ini_depois, n_dias=n_depois,
        equipamento=equipamento,
    )
    # corrige G5 para ser sempre relativo ao H4 (dependente do n_antes real)
    ws_resumo["G5"] = "=H4+1"

    if "M1" in wb.sheetnames:
        _preencher_perfil_semanal(wb, "M1", id_ini=1, n_dias=n_antes)
        _ajustar_grafico_perfil(ws_resumo, "M1", n_antes)
    if "M2" in wb.sheetnames:
        _preencher_perfil_semanal(wb, "M2", id_ini=id_ini_depois, n_dias=n_depois)
        _ajustar_grafico_perfil(ws_resumo, "M2", n_depois)

    _criar_aba_comparativo(wb)

    return dict(n_antes=n_antes, n_depois=n_depois,
                id_antes=(1, n_antes), id_depois=(id_ini_depois, id_ini_depois + n_depois - 1))
