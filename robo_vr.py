from __future__ import annotations

import argparse
import sys
import time
import unicodedata
from dataclasses import dataclass, field
from datetime import date, datetime
from decimal import Decimal
from pathlib import Path
from typing import Any

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

VALOR_VR_LOGISTICA = 35.00
VALOR_VR_PADRAO = 30.00
SETOR_LOGISTICA = "logistica"
PERC_DESCONTO_COLABORADOR = 0.20
DIAS_EXPERIENCIA_LIMITE = 90
LIMITE_ABSENTEISMO_PCT = 10.0
INTERVALO_PROGRESSO_SEGUNDOS = 1.0
FORMATO_MOEDA = 'R$ #,##0.00'
COLUNAS_OBRIGATORIAS = [
    "matricula", "nome", "setor", "dias_trabalhados",
    "dias_uteis_mes", "faltas", "atestados", "admissao",
]
SETORES_CONHECIDOS = {"logistica", "rh", "ti", "financeiro", "comercial"}

HEADER_FILL = PatternFill("solid", fgColor="8C3A46")
HEADER_FONT = Font(color="FFFFFF", bold=True)
HEADER_ALIGNMENT = Alignment(horizontal="center", vertical="center")
ZEBRA_FILL_PAR = PatternFill("solid", fgColor="F5EDEC")
ZEBRA_FILL_IMPAR = PatternFill("solid", fgColor="FFFFFF")
ALERTA_FILL = PatternFill("solid", fgColor="F8C9C4")
ALERTA_FONT = Font(bold=True)
CABECALHO_TERMINAL = "=" * 36


class RoboVRError(Exception):
    pass


class ArquivoNaoEncontradoError(RoboVRError):
    pass


class EstruturaInvalidaError(RoboVRError):
    pass


class PlanilhaVaziaError(RoboVRError):
    pass


@dataclass
class Colaborador:
    linha_excel: int
    matricula: str
    nome: str
    setor: str
    dias_trabalhados: int
    dias_uteis_mes: int
    faltas: int
    atestados: int
    admissao: date


@dataclass
class ResultadoVR:
    matricula: str
    nome: str
    setor: str
    dias_a_pagar: int
    valor_bruto_vr: float
    desconto_colaborador: float
    valor_liquido_empresa: float
    elegibilidade: str
    absenteismo_pct: float
    status: str
    avisos: list[str] = field(default_factory=list)


@dataclass
class Estatisticas:
    total_colaboradores: int = 0
    total_pagar_empresa: Decimal = Decimal("0.00")
    total_alerta_absenteismo: int = 0
    total_em_experiencia: int = 0
    gerado_em: datetime = field(default_factory=datetime.now)


def normalizar_texto(valor: Any) -> str:
    texto = "" if valor is None else str(valor)
    texto = unicodedata.normalize("NFKD", texto)
    return "".join(c for c in texto if not unicodedata.combining(c)).strip().lower()


def texto_obrigatorio(valor: Any, campo: str, linha: int) -> str:
    texto = "" if valor is None else str(valor).strip()
    if not texto:
        raise EstruturaInvalidaError(f"Linha {linha}: campo '{campo}' é obrigatório e não pode estar vazio.")
    return texto


def parse_inteiro_nao_negativo(valor: Any, campo: str, linha: int) -> int:
    try:
        numero = int(valor)
    except (TypeError, ValueError) as exc:
        raise EstruturaInvalidaError(
            f"Linha {linha}: campo '{campo}' inválido ('{valor}'), esperado número inteiro."
        ) from exc
    if numero < 0:
        raise EstruturaInvalidaError(f"Linha {linha}: campo '{campo}' não pode ser negativo ({numero}).")
    return numero


def parse_data_admissao(valor: Any, linha: int) -> date:
    if isinstance(valor, datetime):
        return valor.date()
    if isinstance(valor, date):
        return valor
    if isinstance(valor, str):
        for formato in ("%Y-%m-%d", "%d/%m/%Y"):
            try:
                return datetime.strptime(valor.strip(), formato).date()
            except ValueError:
                pass
    raise EstruturaInvalidaError(
        f"Linha {linha}: data de admissão inválida ('{valor}'). Use AAAA-MM-DD ou DD/MM/AAAA."
    )


def caminho_base_execucao() -> Path:
    return Path(sys.executable).parent if getattr(sys, "frozen", False) else Path(__file__).parent


def localizar_arquivo_entrada(nome_arquivo: str) -> Path:
    for caminho in (Path(nome_arquivo), caminho_base_execucao() / nome_arquivo):
        if caminho.exists():
            return caminho
    raise ArquivoNaoEncontradoError(
        f"Arquivo '{nome_arquivo}' não encontrado no diretório atual nem em '{caminho_base_execucao()}'."
    )


def carregar_planilha(caminho: Path) -> Worksheet:
    try:
        return load_workbook(caminho, data_only=True).active
    except Exception as exc:
        raise EstruturaInvalidaError(
            f"Não foi possível abrir '{caminho.name}'. Verifique se é um .xlsx válido. Detalhe: {exc}"
        ) from exc


def validar_estrutura_colunas(ws: Worksheet) -> dict[str, int]:
    cabecalho = next(ws.iter_rows(min_row=1, max_row=1, values_only=True), None)
    if not cabecalho:
        raise PlanilhaVaziaError("A planilha de entrada está vazia.")

    mapa = {normalizar_texto(valor): idx for idx, valor in enumerate(cabecalho, 1) if valor is not None}
    faltantes = [coluna for coluna in COLUNAS_OBRIGATORIAS if coluna not in mapa]
    if faltantes:
        raise EstruturaInvalidaError(f"Coluna(s) obrigatória(s) ausente(s): {', '.join(faltantes)}.")
    return mapa


def ler_colaboradores(ws: Worksheet, mapa_colunas: dict[str, int]) -> list[Colaborador]:
    colaboradores: list[Colaborador] = []
    matriculas_vistas: dict[str, int] = {}

    for linha_idx, linha in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        if not linha or all(valor is None or str(valor).strip() == "" for valor in linha):
            continue

        def campo(nome: str) -> Any:
            return linha[mapa_colunas[nome] - 1]

        matricula = texto_obrigatorio(campo("matricula"), "matricula", linha_idx)
        nome = texto_obrigatorio(campo("nome"), "nome", linha_idx)
        if matricula in matriculas_vistas:
            raise EstruturaInvalidaError(
                f"Linha {linha_idx}: matrícula duplicada '{matricula}', já encontrada na linha {matriculas_vistas[matricula]}."
            )
        matriculas_vistas[matricula] = linha_idx

        colaborador = Colaborador(
            linha_excel=linha_idx,
            matricula=matricula,
            nome=nome,
            setor=str(campo("setor") or "").strip(),
            dias_trabalhados=parse_inteiro_nao_negativo(campo("dias_trabalhados"), "dias_trabalhados", linha_idx),
            dias_uteis_mes=parse_inteiro_nao_negativo(campo("dias_uteis_mes"), "dias_uteis_mes", linha_idx),
            faltas=parse_inteiro_nao_negativo(campo("faltas"), "faltas", linha_idx),
            atestados=parse_inteiro_nao_negativo(campo("atestados"), "atestados", linha_idx),
            admissao=parse_data_admissao(campo("admissao"), linha_idx),
        )
        if colaborador.faltas > colaborador.dias_uteis_mes:
            raise EstruturaInvalidaError(
                f"Linha {linha_idx}: faltas ({colaborador.faltas}) não podem ser maiores que dias úteis ({colaborador.dias_uteis_mes})."
            )
        colaboradores.append(colaborador)

    if not colaboradores:
        raise PlanilhaVaziaError("Nenhum colaborador encontrado no arquivo de entrada.")
    return colaboradores


def validar_colaborador(c: Colaborador, data_referencia: date) -> list[str]:
    avisos: list[str] = []
    setor = normalizar_texto(c.setor)

    if c.dias_uteis_mes == 0:
        avisos.append("dias_uteis_mes igual a 0; absenteísmo não pôde ser calculado.")
    elif c.dias_trabalhados + c.faltas > c.dias_uteis_mes:
        avisos.append(
            f"dias_trabalhados ({c.dias_trabalhados}) + faltas ({c.faltas}) excede dias_uteis_mes ({c.dias_uteis_mes})."
        )
    if c.admissao > data_referencia:
        avisos.append(f"admissão ({c.admissao.isoformat()}) posterior à data de referência.")
    if not setor:
        avisos.append("setor não informado; aplicado valor padrão.")
    elif setor not in SETORES_CONHECIDOS:
        avisos.append(f"setor '{c.setor}' não reconhecido; aplicado valor padrão.")
    return avisos


def calcular_elegibilidade(c: Colaborador, data_referencia: date) -> tuple[float, str]:
    if (data_referencia - c.admissao).days < DIAS_EXPERIENCIA_LIMITE:
        return 0.5, "Experiência (50%)"
    return 1.0, "Integral (100%)"


def valor_dia_por_setor(setor: str) -> float:
    return VALOR_VR_LOGISTICA if normalizar_texto(setor) == SETOR_LOGISTICA else VALOR_VR_PADRAO


def calcular_valor_vr(c: Colaborador, dias_a_pagar: int, percentual_elegibilidade: float) -> tuple[float, float, float]:
    valor_bruto = dias_a_pagar * valor_dia_por_setor(c.setor) * percentual_elegibilidade
    desconto = valor_bruto * PERC_DESCONTO_COLABORADOR
    return round(valor_bruto, 2), round(desconto, 2), round(valor_bruto - desconto, 2)


def calcular_absenteismo(c: Colaborador) -> tuple[float, bool]:
    if c.dias_uteis_mes == 0:
        return 0.0, False
    percentual = round((c.faltas / c.dias_uteis_mes) * 100, 2)
    return percentual, percentual > LIMITE_ABSENTEISMO_PCT


def processar_colaboradores(colaboradores: list[Colaborador], data_referencia: date) -> tuple[list[ResultadoVR], Estatisticas]:
    resultados: list[ResultadoVR] = []
    stats = Estatisticas()

    for c in colaboradores:
        avisos = validar_colaborador(c, data_referencia)
        for aviso in avisos:
            print(f"  [AVISO] Linha {c.linha_excel} ({c.nome}): {aviso}")

        percentual, elegibilidade = calcular_elegibilidade(c, data_referencia)
        dias_a_pagar = c.dias_trabalhados + c.atestados
        bruto, desconto, liquido = calcular_valor_vr(c, dias_a_pagar, percentual)
        absenteismo, em_alerta = calcular_absenteismo(c)

        resultados.append(ResultadoVR(
            c.matricula, c.nome, c.setor, dias_a_pagar, bruto, desconto, liquido,
            elegibilidade, absenteismo, "ALERTA" if em_alerta else "OK", avisos,
        ))
        stats.total_colaboradores += 1
        stats.total_pagar_empresa += Decimal(str(liquido))
        stats.total_alerta_absenteismo += int(em_alerta)
        stats.total_em_experiencia += int(percentual < 1.0)

    stats.total_pagar_empresa = stats.total_pagar_empresa.quantize(Decimal("0.01"))
    return resultados, stats


def texto_para_largura(valor: Any, number_format: str | None = None) -> str:
    if valor is None:
        return ""
    if isinstance(valor, datetime):
        return valor.strftime("%d/%m/%Y %H:%M:%S")
    if isinstance(valor, date):
        return valor.strftime("%d/%m/%Y")
    if number_format == FORMATO_MOEDA:
        try:
            return f"R$ {float(valor):,.2f}"
        except (TypeError, ValueError):
            pass
    return str(valor)


def ajustar_largura_colunas(ws: Worksheet) -> None:
    for col_cells in ws.columns:
        maior = max(
            max((len(linha) for linha in texto_para_largura(c.value, c.number_format).splitlines()), default=0)
            for c in col_cells
        )
        ws.column_dimensions[get_column_letter(col_cells[0].column)].width = min(max(maior + 4, 10), 60)


def aplicar_estilo_cabecalho(ws: Worksheet) -> None:
    for celula in ws[1]:
        celula.fill = HEADER_FILL
        celula.font = HEADER_FONT
        celula.alignment = HEADER_ALIGNMENT


def montar_aba_detalhamento(wb: Workbook, resultados: list[ResultadoVR]) -> None:
    ws = wb.active
    ws.title = "Detalhamento"
    ws.append([
        "Matrícula", "Nome", "Setor", "Dias a Pagar", "Valor Bruto VR",
        "Desconto Colaborador", "Valor Líquido Empresa", "Elegibilidade", "Absenteísmo (%)", "Status",
    ])
    aplicar_estilo_cabecalho(ws)

    for linha_idx, r in enumerate(resultados, start=2):
        ws.append([
            r.matricula, r.nome, r.setor, r.dias_a_pagar, r.valor_bruto_vr,
            r.desconto_colaborador, r.valor_liquido_empresa, r.elegibilidade, r.absenteismo_pct, r.status,
        ])
        fill = ZEBRA_FILL_PAR if linha_idx % 2 == 0 else ZEBRA_FILL_IMPAR
        for celula in ws[linha_idx]:
            celula.fill = fill
        for coluna in ("E", "F", "G"):
            ws[f"{coluna}{linha_idx}"].number_format = FORMATO_MOEDA
        if r.status == "ALERTA":
            ws[f"J{linha_idx}"].fill = ALERTA_FILL
            ws[f"J{linha_idx}"].font = ALERTA_FONT

    ajustar_largura_colunas(ws)
    ws.freeze_panes = "A2"


def montar_aba_resumo(wb: Workbook, stats: Estatisticas) -> None:
    ws = wb.create_sheet("Resumo")
    ws.append(["Indicador", "Valor"])
    aplicar_estilo_cabecalho(ws)
    for rotulo, valor in [
        ("Total de colaboradores processados", stats.total_colaboradores),
        ("Total a pagar pela empresa", stats.total_pagar_empresa),
        ("Total em alerta de absenteísmo", stats.total_alerta_absenteismo),
        ("Total em período de experiência", stats.total_em_experiencia),
        ("Relatório gerado em", stats.gerado_em.strftime("%d/%m/%Y %H:%M:%S")),
    ]:
        ws.append([rotulo, valor])
    ws["B3"].number_format = FORMATO_MOEDA
    ajustar_largura_colunas(ws)


def gerar_relatorio_saida(resultados: list[ResultadoVR], stats: Estatisticas, caminho_saida: Path) -> None:
    wb = Workbook()
    montar_aba_detalhamento(wb, resultados)
    montar_aba_resumo(wb, stats)
    try:
        wb.save(caminho_saida)
    except PermissionError as exc:
        raise RoboVRError(f"Não foi possível salvar '{caminho_saida.name}'. Feche o arquivo no Excel e tente novamente.") from exc


def parse_argumentos(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Robô de cálculo do Vale-Refeição (VR)")
    parser.add_argument("--entrada", default="colaboradores.xlsx", help="Arquivo de entrada .xlsx")
    parser.add_argument("--saida", default=None, help="Arquivo de saída .xlsx")
    parser.add_argument("--data-ref", default=None, help="Data de referência no formato AAAA-MM-DD")
    return parser.parse_args(argv)


def resolver_data_referencia(data_ref: str | None) -> date:
    if data_ref is None:
        return date.today()
    try:
        return datetime.strptime(data_ref, "%Y-%m-%d").date()
    except ValueError as exc:
        raise RoboVRError(f"--data-ref inválida ('{data_ref}'). Use o formato AAAA-MM-DD.") from exc


def exibir_progresso(mensagem: str) -> None:
    print(mensagem)
    time.sleep(INTERVALO_PROGRESSO_SEGUNDOS)


def exibir_resumo(stats: Estatisticas, caminho_saida: Path) -> None:
    print(CABECALHO_TERMINAL)
    print(f"Arquivo: {caminho_saida.name}")
    print(f"Colaboradores processados: {stats.total_colaboradores}")
    print(f"Total a pagar pela empresa: R$ {stats.total_pagar_empresa:,.2f}")
    print(f"Em alerta de absenteísmo: {stats.total_alerta_absenteismo}")
    print(f"Em período de experiência: {stats.total_em_experiencia}")


def aguardar_antes_de_sair() -> None:
    if getattr(sys, "frozen", False):
        input("\nPressione ENTER para sair...")


def main(argv: list[str] | None = None) -> int:
    args = parse_argumentos(argv)
    print(CABECALHO_TERMINAL)
    print("ROBÔ DE CÁLCULO DE VALE REFEIÇÃO")
    print(CABECALHO_TERMINAL)

    try:
        data_referencia = resolver_data_referencia(args.data_ref)
        nome_saida = args.saida or f"relatorio_vr_{data_referencia.strftime('%m%Y')}.xlsx"
        caminho_saida = caminho_base_execucao() / nome_saida

        exibir_progresso("Lendo arquivo...")
        ws = carregar_planilha(localizar_arquivo_entrada(args.entrada))
        exibir_progresso("Validando colunas...")
        colaboradores = ler_colaboradores(ws, validar_estrutura_colunas(ws))
        exibir_progresso("Aplicando regras...")
        resultados, stats = processar_colaboradores(colaboradores, data_referencia)
        exibir_progresso("Gerando relatório...")
        exibir_progresso("Formatando planilha...")
        gerar_relatorio_saida(resultados, stats, caminho_saida)
        exibir_progresso("Relatório salvo com sucesso.")
        exibir_resumo(stats, caminho_saida)
        aguardar_antes_de_sair()
        return 0
    except RoboVRError as exc:
        print(f"\n[ERRO] {exc}\n{CABECALHO_TERMINAL}")
        aguardar_antes_de_sair()
        return 1
    except Exception as exc:
        print(f"\n[ERRO INESPERADO] {exc}\n{CABECALHO_TERMINAL}")
        aguardar_antes_de_sair()
        return 1


if __name__ == "__main__":
    sys.exit(main())
