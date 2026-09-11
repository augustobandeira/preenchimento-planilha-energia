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
