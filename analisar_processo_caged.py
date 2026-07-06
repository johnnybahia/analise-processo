#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Análise de inconsistências: Processo Trabalhista (PJe-Calc) x CAGED
====================================================================

Lê as planilhas de cálculo do processo (PDF, intervalo de páginas configurável,
padrão 351 a 1142), extrai os dados de cada reclamante, cruza com os vínculos
do extrato CAGED (PDF) e gera um relatório minucioso de inconsistências.

Verificações realizadas
-----------------------
Por reclamante (processo x CAGED):
  1. Reclamante SEM vínculo localizado no CAGED (crítica)
  2. Data de admissão divergente entre processo e CAGED
  3. Data de demissão/desligamento divergente entre processo e CAGED
  4. Período do cálculo INICIA ANTES da admissão registrada no CAGED
  5. Período do cálculo TERMINA APÓS o desligamento registrado no CAGED
  6. Salário base do histórico salarial divergente da remuneração do CAGED
  7. Meses cobrados no cálculo em que o CAGED registra remuneração ZERADA
  8. Reclamante com MAIS DE UMA planilha de cálculo (possível duplicidade)
  9. Correspondência apenas aproximada de nome (grafia divergente)

Consistência interna do processo:
 10. Admissão/Demissão da planilha incompatíveis com o Período do Cálculo
 11. Soma das verbas do demonstrativo divergente do total do resumo
 12. Páginas sem texto extraível (provável imagem escaneada -> requer OCR)

Uso
---
    python3 analisar_processo_caged.py PROCESSO.pdf CAGED.pdf \
        --pagina-inicial 351 --pagina-final 1142 \
        --saida relatorio_inconsistencias.xlsx

Dependências: pdfplumber, pandas, openpyxl  (pip install -r requirements.txt)
"""

import argparse
import re
import sys
import unicodedata
from dataclasses import dataclass, field
from datetime import datetime
from difflib import SequenceMatcher

import pdfplumber

# ---------------------------------------------------------------------------
# Utilidades
# ---------------------------------------------------------------------------

MESES_PT = {
    "Janeiro": 1, "Fevereiro": 2, "Março": 3, "Abril": 4, "Maio": 5,
    "Junho": 6, "Julho": 7, "Agosto": 8, "Setembro": 9, "Outubro": 10,
    "Novembro": 11, "Dezembro": 12,
}

RE_DATA = r"\d{2}/\d{2}/\d{4}"
RE_VALOR = r"[\d.]*\d,\d{2}"


def normalizar_nome(nome: str) -> str:
    """Remove acentos, espaços duplicados e caixa para comparação de nomes."""
    nome = unicodedata.normalize("NFKD", nome)
    nome = "".join(c for c in nome if not unicodedata.combining(c))
    return re.sub(r"\s+", " ", nome).strip().upper()


def parse_valor(txt: str) -> float:
    """Converte '1.234,56' -> 1234.56."""
    return float(txt.replace(".", "").replace(",", "."))


def parse_data(txt: str):
    try:
        return datetime.strptime(txt.strip(), "%d/%m/%Y").date()
    except (ValueError, AttributeError):
        return None


def fmt_data(d) -> str:
    return d.strftime("%d/%m/%Y") if d else "-"


def similaridade(a: str, b: str) -> float:
    return SequenceMatcher(None, a, b).ratio()


# ---------------------------------------------------------------------------
# Estruturas de dados
# ---------------------------------------------------------------------------

@dataclass
class Reclamante:
    """Dados extraídos de uma planilha de cálculo do PJe-Calc."""
    nome: str = ""
    numero_calculo: str = ""
    processo: str = ""
    paginas_pdf: list = field(default_factory=list)   # páginas do PDF
    fls: list = field(default_factory=list)           # numeração de folhas (Fls.)
    periodo_inicio: object = None
    periodo_fim: object = None
    data_ajuizamento: object = None
    admissao: object = None
    demissao: object = None
    historico_salarial: dict = field(default_factory=dict)  # {'MM/AAAA': valor}
    verbas: dict = field(default_factory=dict)               # {nome: total}
    meses_cobrados: set = field(default_factory=set)         # {(ano, mes)}
    total_bruto: float = None
    total_devido: float = None

    @property
    def nome_norm(self) -> str:
        return normalizar_nome(self.nome)


@dataclass
class VinculoCaged:
    """Dados extraídos de um vínculo do extrato CAGED/RAIS."""
    pis: str = ""
    cpf: str = ""
    nome: str = ""
    nascimento: object = None
    admissao: object = None
    desligamento: object = None
    causa_desligamento: str = ""
    tipo_salario: str = ""
    horas_semanais: str = ""
    salario_contratual: float = None
    tipo_vinculo: str = ""
    cbo: str = ""
    remuneracoes: dict = field(default_factory=dict)  # {mes(1-12): valor}
    ano_referencia: int = None
    aviso_previo: float = None
    ferias_indenizadas: float = None
    pagina_pdf: int = None

    @property
    def nome_norm(self) -> str:
        return normalizar_nome(self.nome)

    @property
    def salario_mensal_estimado(self):
        """Se o salário é por hora, converte para mensal (h/semana * 220/44)."""
        if self.salario_contratual is None:
            return None
        if self.tipo_salario.startswith("5"):  # 5 - Horário
            try:
                horas = float(self.horas_semanais)
            except (TypeError, ValueError):
                horas = 44.0
            return round(self.salario_contratual * horas * 5.0, 2)
        return self.salario_contratual


@dataclass
class Inconsistencia:
    reclamante: str
    tipo: str
    gravidade: str      # CRÍTICA / ALTA / MÉDIA / INFORMATIVA
    descricao: str
    valor_processo: str = ""
    valor_caged: str = ""
    paginas: str = ""


# ---------------------------------------------------------------------------
# Parser do PROCESSO (planilhas PJe-Calc)
# ---------------------------------------------------------------------------

def extrair_reclamantes(caminho_pdf, pag_inicial, pag_final, log=print):
    """Percorre o intervalo de páginas e agrupa as planilhas por nº de cálculo."""
    reclamantes = []
    paginas_sem_texto = []
    atual = None

    with pdfplumber.open(caminho_pdf) as pdf:
        total = len(pdf.pages)
        ini = max(1, pag_inicial)
        fim = min(pag_final, total)
        if ini > total:
            raise SystemExit(
                f"ERRO: página inicial {pag_inicial} maior que o total de "
                f"páginas do PDF ({total})."
            )
        log(f"  PDF do processo tem {total} páginas; lendo páginas {ini} a {fim}...")

        for num in range(ini, fim + 1):
            pagina = pdf.pages[num - 1]
            texto = pagina.extract_text() or ""
            if not texto.strip():
                paginas_sem_texto.append(num)
                continue

            m_calc = re.search(r"C[áa]lculo:\s*(\d+)", texto)
            # Novo número de cálculo => nova planilha / novo reclamante
            if m_calc and (atual is None or m_calc.group(1) != atual.numero_calculo):
                if atual:
                    reclamantes.append(atual)
                atual = Reclamante(numero_calculo=m_calc.group(1))

            if atual is None:
                # Página anterior ao primeiro cabeçalho de cálculo — ignora,
                # mas registra para auditoria.
                continue

            atual.paginas_pdf.append(num)
            m_fls = re.search(r"Fls\.?:\s*(\d+)", texto)
            if m_fls:
                atual.fls.append(int(m_fls.group(1)))
            _preencher_reclamante(atual, texto)

            if num % 100 == 0:
                log(f"    ... página {num} processada "
                    f"({len(reclamantes)} planilhas concluídas)")

    if atual:
        reclamantes.append(atual)

    return reclamantes, paginas_sem_texto


def _preencher_reclamante(r: Reclamante, texto: str):
    """Extrai os campos de interesse do texto de uma página da planilha."""
    if not r.processo:
        m = re.search(r"Processo:\s*([\d.\-]+)", texto)
        if m:
            r.processo = m.group(1)

    if not r.nome:
        m = re.search(r"Reclamante:\s*(.+)", texto)
        if m:
            r.nome = m.group(1).strip()

    if not r.periodo_inicio:
        m = re.search(
            rf"Per[íi]odo do C[áa]lculo:\s*({RE_DATA})\s*a\s*({RE_DATA})", texto)
        if m:
            r.periodo_inicio = parse_data(m.group(1))
            r.periodo_fim = parse_data(m.group(2))

    if not r.data_ajuizamento:
        m = re.search(rf"Data Ajuizamento:\s*({RE_DATA})", texto)
        if m:
            r.data_ajuizamento = parse_data(m.group(1))

    if not r.admissao:
        m = re.search(rf"Admiss[ãa]o:\s*({RE_DATA})", texto)
        if m:
            r.admissao = parse_data(m.group(1))

    if not r.demissao:
        m = re.search(rf"Demiss[ãa]o:\s*({RE_DATA})", texto)
        if m:
            r.demissao = parse_data(m.group(1))

    # Histórico salarial: linhas "MM/AAAA 1.100,00"
    em_historico = False
    for linha in texto.splitlines():
        if "HISTÓRICO SALARIAL" in linha.upper():
            em_historico = True
            continue
        if em_historico:
            m = re.match(rf"^(\d{{2}}/\d{{4}})\s+({RE_VALOR})\s*$", linha.strip())
            if m:
                r.historico_salarial[m.group(1)] = parse_valor(m.group(2))
            elif linha.strip() and not re.match(
                    r"^(MÊS/ANO|OCORR)", linha.strip().upper()):
                # saiu da tabela
                if not re.match(rf"^\d{{2}}/\d{{4}}", linha.strip()):
                    em_historico = False

    # Meses cobrados nas verbas: "22 a 28/02/2021 ..." / "01 a 27/04/2021 ..."
    for m in re.finditer(rf"^\d{{2}} a \d{{2}}/(\d{{2}})/(\d{{4}})\s",
                         texto, re.MULTILINE):
        r.meses_cobrados.add((int(m.group(2)), int(m.group(1))))

    # Totais das verbas no Resumo do Cálculo:
    # "MULTA CONVENCIONAL 604,15 109,16 713,31"
    em_resumo = "Resumo do Cálculo" in texto
    if em_resumo:
        for m in re.finditer(
                rf"^([A-ZÀ-Ü0-9 ()\-\.º/&]+?)\s+({RE_VALOR})\s+({RE_VALOR})\s+({RE_VALOR})\s*$",
                texto, re.MULTILINE):
            nome_verba = m.group(1).strip()
            if not any(p in nome_verba.upper() for p in
                       ("TOTAL", "DESCRIÇÃO", "PÁG", "OCORRÊNCIA")):
                r.verbas[nome_verba] = parse_valor(m.group(4))
        m = re.search(
            rf"^Total\s+({RE_VALOR})\s+({RE_VALOR})\s+({RE_VALOR})\s*$",
            texto, re.MULTILINE)
        if m:
            r.total_bruto = parse_valor(m.group(3))

    m = re.search(rf"Total Devido pelo Reclamado\s+({RE_VALOR})", texto)
    if m:
        r.total_devido = parse_valor(m.group(1))


# ---------------------------------------------------------------------------
# Parser do CAGED
# ---------------------------------------------------------------------------

def extrair_vinculos_caged(caminho_pdf, log=print):
    """Extrai todos os vínculos do extrato CAGED."""
    vinculos = []
    texto_total = []

    with pdfplumber.open(caminho_pdf) as pdf:
        log(f"  PDF do CAGED tem {len(pdf.pages)} páginas; lendo todas...")
        for num, pagina in enumerate(pdf.pages, start=1):
            t = pagina.extract_text() or ""
            texto_total.append((num, t))

    # Ano de referência do extrato, se declarado (ex.: "Ano Base 2019")
    ano_ref_global = None
    for _, t in texto_total:
        m = re.search(r"Ano[- ]?Base[:\s]+(\d{4})", t, re.IGNORECASE)
        if m:
            ano_ref_global = int(m.group(1))
            break

    texto_junto = "\n".join(f"\x0c{num}\n{t}" for num, t in texto_total)

    # Cada vínculo inicia em "PIS: xxx Nome: FULANO"
    blocos = re.split(r"(?=PIS:\s*[\d.\-]+\s*Nome:)", texto_junto)
    for bloco in blocos:
        m_pis = re.match(r"PIS:\s*([\d.\-]+)\s*Nome:\s*(.+)", bloco)
        if not m_pis:
            continue
        v = VinculoCaged(pis=m_pis.group(1).strip(),
                         nome=m_pis.group(2).strip(),
                         ano_referencia=ano_ref_global)

        m_pg = re.search(r"\x0c(\d+)\n", bloco)
        if m_pg:
            v.pagina_pdf = int(m_pg.group(1))

        m = re.search(r"(\d{3}\.\d{3}\.\d{3}-\d{2})", bloco)
        if m:
            v.cpf = m.group(1)

        m = re.search(rf"Nascimento\s*\n[^\n]*?({RE_DATA})", bloco)
        if m:
            v.nascimento = parse_data(m.group(1))

        m = re.search(rf"Data de Admiss[ãa]o[^\n]*\n({RE_DATA})", bloco)
        if m:
            v.admissao = parse_data(m.group(1))

        # Salário: "5 - Horário 44 4,64 784205 - ALIMENTADOR..."
        m = re.search(
            rf"Tipo de Sal[áa]rio[^\n]*\n(\d+ - [^\d\n]+?)\s+(\d+)\s+({RE_VALOR})\s+(\d+)\s*-\s*(.+)",
            bloco)
        if m:
            v.tipo_salario = m.group(1).strip()
            v.horas_semanais = m.group(2)
            v.salario_contratual = parse_valor(m.group(3))
            v.cbo = f"{m.group(4)} - {m.group(5).strip()}"

        m = re.search(r"Tipo de V[íi]nculo\s*\n(\d+ - .+)", bloco)
        if m:
            v.tipo_vinculo = m.group(1).strip()

        # Desligamento: "Data 18/04/2019" (na seção de desligamento)
        m = re.search(rf"\bData\s+({RE_DATA})", bloco)
        if m:
            v.desligamento = parse_data(m.group(1))

        m = re.search(r"Causa\s+(\d+\s*-\s*.+)", bloco)
        if m:
            v.causa_desligamento = m.group(1).strip()

        m = re.search(rf"Valor Aviso Pr[ée]vio\s+({RE_VALOR})", bloco)
        if m:
            v.aviso_previo = parse_valor(m.group(1))

        m = re.search(rf"F[ée]rias Indenizadas\s+({RE_VALOR})", bloco)
        if m:
            v.ferias_indenizadas = parse_valor(m.group(1))

        # Remunerações mensais: "Janeiro 898,91 0 Julho 0,00 0"
        for m in re.finditer(
                rf"(Janeiro|Fevereiro|Mar[çc]o|Abril|Maio|Junho|Julho|Agosto|"
                rf"Setembro|Outubro|Novembro|Dezembro)\s+({RE_VALOR}|0,00)\s+\d+",
                bloco):
            nome_mes = m.group(1).replace("Marco", "Março")
            v.remuneracoes[MESES_PT[nome_mes]] = parse_valor(m.group(2))

        # Sem ano-base declarado: infere pelo desligamento (extratos anuais
        # RAIS/CAGED normalmente referem-se ao ano do desligamento)
        if v.ano_referencia is None and v.desligamento:
            v.ano_referencia = v.desligamento.year

        vinculos.append(v)

    return vinculos


# ---------------------------------------------------------------------------
# Cruzamento e verificação de inconsistências
# ---------------------------------------------------------------------------

def localizar_vinculos(reclamante, vinculos, limiar_fuzzy=0.88):
    """Retorna (vínculos do reclamante, exato?) — casa por nome normalizado."""
    alvo = reclamante.nome_norm
    exatos = [v for v in vinculos if v.nome_norm == alvo]
    if exatos:
        return exatos, True
    aproximados = [(similaridade(alvo, v.nome_norm), v) for v in vinculos]
    aproximados = [(s, v) for s, v in aproximados if s >= limiar_fuzzy]
    aproximados.sort(key=lambda x: -x[0])
    return [v for _, v in aproximados], False


def escolher_vinculo(reclamante, candidatos):
    """Entre vários vínculos do mesmo nome, escolhe o que melhor cobre o
    período do cálculo (readmissões geram mais de um vínculo)."""
    if len(candidatos) == 1 or not reclamante.periodo_inicio:
        return candidatos[0]

    def sobreposicao(v):
        ini = max(filter(None, [v.admissao, reclamante.periodo_inicio]))
        fim_v = v.desligamento or reclamante.periodo_fim
        fim = min(filter(None, [fim_v, reclamante.periodo_fim]))
        return (fim - ini).days if ini and fim else -10**6

    return max(candidatos, key=sobreposicao)


def verificar(reclamantes, vinculos, tolerancia_salario=0.05, log=print):
    inconsistencias = []

    def add(rec, tipo, gravidade, descricao, vp="", vc=""):
        pags = ""
        if rec.fls:
            pags = f"Fls. {min(rec.fls)}-{max(rec.fls)}"
        elif rec.paginas_pdf:
            pags = f"pág. PDF {min(rec.paginas_pdf)}-{max(rec.paginas_pdf)}"
        inconsistencias.append(Inconsistencia(
            reclamante=rec.nome or f"(cálculo {rec.numero_calculo})",
            tipo=tipo, gravidade=gravidade, descricao=descricao,
            valor_processo=str(vp), valor_caged=str(vc), paginas=pags))

    # 8) Duplicidade de planilhas
    vistos = {}
    for rec in reclamantes:
        vistos.setdefault(rec.nome_norm, []).append(rec)
    for nome, lst in vistos.items():
        if nome and len(lst) > 1:
            calcs = ", ".join(r.numero_calculo for r in lst)
            for rec in lst:
                add(rec, "DUPLICIDADE DE PLANILHA", "ALTA",
                    f"Reclamante aparece em {len(lst)} planilhas de cálculo "
                    f"(cálculos nº {calcs}). Verificar possível cobrança em dobro.")

    for rec in reclamantes:
        # 10) Consistência interna: admissão/demissão x período do cálculo
        if rec.admissao and rec.periodo_inicio and rec.periodo_inicio < rec.admissao:
            add(rec, "PERÍODO ANTERIOR À ADMISSÃO (INTERNO)", "ALTA",
                "O período do cálculo inicia antes da própria data de admissão "
                "informada na planilha.",
                f"início {fmt_data(rec.periodo_inicio)}",
                f"admissão (planilha) {fmt_data(rec.admissao)}")
        if rec.demissao and rec.periodo_fim and rec.periodo_fim > rec.demissao:
            add(rec, "PERÍODO POSTERIOR À DEMISSÃO (INTERNO)", "ALTA",
                "O período do cálculo termina depois da própria data de demissão "
                "informada na planilha.",
                f"fim {fmt_data(rec.periodo_fim)}",
                f"demissão (planilha) {fmt_data(rec.demissao)}")

        # 11) Soma das verbas x total do resumo
        if rec.verbas and rec.total_bruto is not None:
            soma = round(sum(rec.verbas.values()), 2)
            if abs(soma - rec.total_bruto) > 0.02:
                add(rec, "SOMA DAS VERBAS DIVERGE DO TOTAL", "MÉDIA",
                    "A soma das verbas do resumo não confere com o total bruto "
                    "apresentado na planilha.",
                    f"soma verbas {soma:.2f}", f"total planilha {rec.total_bruto:.2f}")

        # Cruzamento com CAGED
        candidatos, exato = localizar_vinculos(rec, vinculos)
        if not candidatos:
            add(rec, "SEM VÍNCULO NO CAGED", "CRÍTICA",
                "Reclamante não localizado em nenhum vínculo do extrato CAGED. "
                "Sem registro de vínculo empregatício no período, a cobrança "
                "pode ser indevida — conferir manualmente por CPF/PIS.",
                f"período cobrado {fmt_data(rec.periodo_inicio)} a "
                f"{fmt_data(rec.periodo_fim)}", "nenhum vínculo encontrado")
            continue

        if not exato:
            v0 = candidatos[0]
            add(rec, "NOME APENAS APROXIMADO", "MÉDIA",
                "O nome no processo não é idêntico ao do CAGED (possível erro "
                "de grafia ou homônimo). Conferir CPF/PIS manualmente.",
                rec.nome, f"{v0.nome} (CPF {v0.cpf or '-'})")

        v = escolher_vinculo(rec, candidatos)

        # 2) Admissão divergente
        if rec.admissao and v.admissao and rec.admissao != v.admissao:
            dias = abs((rec.admissao - v.admissao).days)
            add(rec, "ADMISSÃO DIVERGENTE", "ALTA" if dias > 5 else "MÉDIA",
                f"Data de admissão do processo difere do CAGED em {dias} dia(s).",
                fmt_data(rec.admissao), fmt_data(v.admissao))

        # 3) Demissão divergente
        if rec.demissao and v.desligamento and rec.demissao != v.desligamento:
            dias = abs((rec.demissao - v.desligamento).days)
            add(rec, "DEMISSÃO DIVERGENTE", "ALTA" if dias > 5 else "MÉDIA",
                f"Data de demissão do processo difere do desligamento no CAGED "
                f"em {dias} dia(s).",
                fmt_data(rec.demissao), fmt_data(v.desligamento))

        # 4) Cálculo antes da admissão real
        if rec.periodo_inicio and v.admissao and rec.periodo_inicio < v.admissao:
            dias = (v.admissao - rec.periodo_inicio).days
            add(rec, "CÁLCULO ANTES DA ADMISSÃO (CAGED)", "CRÍTICA",
                f"O cálculo cobra {dias} dia(s) ANTERIORES à admissão registrada "
                "no CAGED — período sem vínculo empregatício.",
                f"início do cálculo {fmt_data(rec.periodo_inicio)}",
                f"admissão CAGED {fmt_data(v.admissao)}")

        # 5) Cálculo depois do desligamento real
        if rec.periodo_fim and v.desligamento and rec.periodo_fim > v.desligamento:
            dias = (rec.periodo_fim - v.desligamento).days
            add(rec, "CÁLCULO APÓS O DESLIGAMENTO (CAGED)", "CRÍTICA",
                f"O cálculo cobra {dias} dia(s) POSTERIORES ao desligamento "
                "registrado no CAGED — período sem vínculo empregatício.",
                f"fim do cálculo {fmt_data(rec.periodo_fim)}",
                f"desligamento CAGED {fmt_data(v.desligamento)}")

        # 6) Salário base x remuneração CAGED
        for mes_ano, salario in rec.historico_salarial.items():
            mm, aaaa = int(mes_ano[:2]), int(mes_ano[3:])
            if v.ano_referencia == aaaa and mm in v.remuneracoes:
                remun = v.remuneracoes[mm]
                if remun > 0 and salario > 0:
                    desvio = abs(salario - remun) / salario
                    if desvio > tolerancia_salario:
                        add(rec, "SALÁRIO DIVERGENTE", "MÉDIA",
                            f"Salário base usado no cálculo em {mes_ano} difere "
                            f"{desvio*100:.1f}% da remuneração do CAGED. (A "
                            "remuneração CAGED pode incluir extras — conferir.)",
                            f"{salario:.2f}", f"{remun:.2f}")

        # 7) Meses cobrados com remuneração zerada no CAGED
        for (aaaa, mm) in sorted(rec.meses_cobrados):
            if v.ano_referencia == aaaa and v.remuneracoes.get(mm, None) == 0.0:
                add(rec, "MÊS COBRADO SEM REMUNERAÇÃO NO CAGED", "ALTA",
                    f"O cálculo cobra verbas em {mm:02d}/{aaaa}, mas o CAGED "
                    "registra remuneração 0,00 nesse mês (sem trabalho "
                    "registrado).",
                    f"mês cobrado {mm:02d}/{aaaa}", "remuneração CAGED 0,00")

    return inconsistencias


# ---------------------------------------------------------------------------
# Relatórios
# ---------------------------------------------------------------------------

def gerar_relatorios(reclamantes, vinculos, inconsistencias, paginas_sem_texto,
                     saida_xlsx, saida_txt, log=print):
    import pandas as pd

    ordem_grav = {"CRÍTICA": 0, "ALTA": 1, "MÉDIA": 2, "INFORMATIVA": 3}
    inconsistencias = sorted(
        inconsistencias, key=lambda i: (ordem_grav.get(i.gravidade, 9),
                                        i.reclamante, i.tipo))

    df_inc = pd.DataFrame([{
        "Reclamante": i.reclamante,
        "Gravidade": i.gravidade,
        "Tipo": i.tipo,
        "Descrição": i.descricao,
        "Valor no Processo": i.valor_processo,
        "Valor no CAGED": i.valor_caged,
        "Localização no Processo": i.paginas,
    } for i in inconsistencias])

    df_rec = pd.DataFrame([{
        "Reclamante": r.nome,
        "Nº Cálculo": r.numero_calculo,
        "Fls.": f"{min(r.fls)}-{max(r.fls)}" if r.fls else "",
        "Págs. PDF": f"{min(r.paginas_pdf)}-{max(r.paginas_pdf)}"
                     if r.paginas_pdf else "",
        "Período Início": fmt_data(r.periodo_inicio),
        "Período Fim": fmt_data(r.periodo_fim),
        "Admissão (planilha)": fmt_data(r.admissao),
        "Demissão (planilha)": fmt_data(r.demissao),
        "Ajuizamento": fmt_data(r.data_ajuizamento),
        "Meses no Histórico": len(r.historico_salarial),
        "Total Bruto": r.total_bruto,
        "Total Devido pelo Reclamado": r.total_devido,
        "Verbas": "; ".join(f"{k}={v:.2f}" for k, v in r.verbas.items()),
    } for r in reclamantes])

    df_cag = pd.DataFrame([{
        "Nome": v.nome,
        "CPF": v.cpf,
        "PIS": v.pis,
        "Admissão": fmt_data(v.admissao),
        "Desligamento": fmt_data(v.desligamento),
        "Causa Desligamento": v.causa_desligamento,
        "Tipo Salário": v.tipo_salario,
        "Salário Contratual": v.salario_contratual,
        "Salário Mensal Estimado": v.salario_mensal_estimado,
        "Ano Referência": v.ano_referencia,
        "Remunerações (mês=valor)": "; ".join(
            f"{m:02d}={val:.2f}" for m, val in sorted(v.remuneracoes.items())),
        "CBO": v.cbo,
        "Pág. PDF": v.pagina_pdf,
    } for v in vinculos])

    with pd.ExcelWriter(saida_xlsx, engine="openpyxl") as writer:
        df_inc.to_excel(writer, sheet_name="Inconsistências", index=False)
        df_rec.to_excel(writer, sheet_name="Reclamantes (Processo)", index=False)
        df_cag.to_excel(writer, sheet_name="Vínculos (CAGED)", index=False)
        if paginas_sem_texto:
            pd.DataFrame({"Página sem texto (requer OCR)": paginas_sem_texto}) \
                .to_excel(writer, sheet_name="Páginas sem texto", index=False)
        # Ajuste de largura das colunas
        for aba in writer.sheets.values():
            for col in aba.columns:
                largura = max((len(str(c.value)) for c in col
                               if c.value is not None), default=10)
                aba.column_dimensions[col[0].column_letter].width = \
                    min(largura + 2, 80)

    # Relatório em texto
    linhas = []
    linhas.append("=" * 78)
    linhas.append("RELATÓRIO DE ANÁLISE — PROCESSO x CAGED")
    linhas.append(f"Gerado em: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}")
    linhas.append("=" * 78)
    linhas.append(f"Planilhas de cálculo lidas (reclamantes): {len(reclamantes)}")
    linhas.append(f"Vínculos no CAGED: {len(vinculos)}")
    linhas.append(f"Inconsistências encontradas: {len(inconsistencias)}")
    if paginas_sem_texto:
        linhas.append(f"ATENÇÃO: {len(paginas_sem_texto)} página(s) sem texto "
                      f"extraível (PDF escaneado?) — necessitam OCR: "
                      f"{paginas_sem_texto[:20]}"
                      f"{'...' if len(paginas_sem_texto) > 20 else ''}")
    linhas.append("")
    por_grav = {}
    for i in inconsistencias:
        por_grav[i.gravidade] = por_grav.get(i.gravidade, 0) + 1
    for g in ("CRÍTICA", "ALTA", "MÉDIA", "INFORMATIVA"):
        if g in por_grav:
            linhas.append(f"  {g:12s}: {por_grav[g]}")
    linhas.append("")
    atual = None
    for i in inconsistencias:
        if i.reclamante != atual:
            atual = i.reclamante
            linhas.append("-" * 78)
            linhas.append(f"RECLAMANTE: {i.reclamante}   ({i.paginas})")
        linhas.append(f"  [{i.gravidade}] {i.tipo}")
        linhas.append(f"      {i.descricao}")
        if i.valor_processo or i.valor_caged:
            linhas.append(f"      Processo: {i.valor_processo} | "
                          f"CAGED: {i.valor_caged}")
    if not inconsistencias:
        linhas.append("Nenhuma inconsistência encontrada com os critérios atuais.")
    with open(saida_txt, "w", encoding="utf-8") as f:
        f.write("\n".join(linhas) + "\n")

    return df_inc


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    ap = argparse.ArgumentParser(
        description="Cruza planilhas de cálculo do processo (PJe-Calc) com o "
                    "extrato CAGED e aponta inconsistências.",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter)
    ap.add_argument("processo", help="PDF do processo (planilhas de cálculo)")
    ap.add_argument("caged", help="PDF do extrato CAGED")
    ap.add_argument("--pagina-inicial", type=int, default=351,
                    help="Primeira página do PDF do processo a ler")
    ap.add_argument("--pagina-final", type=int, default=1142,
                    help="Última página do PDF do processo a ler")
    ap.add_argument("--saida", default="relatorio_inconsistencias.xlsx",
                    help="Arquivo Excel de saída")
    ap.add_argument("--tolerancia-salario", type=float, default=0.05,
                    help="Tolerância (fração) para divergência salarial")
    args = ap.parse_args()

    saida_txt = re.sub(r"\.xlsx?$", "", args.saida) + ".txt"

    print("[1/4] Lendo planilhas de cálculo do processo...")
    reclamantes, paginas_sem_texto = extrair_reclamantes(
        args.processo, args.pagina_inicial, args.pagina_final)
    print(f"  -> {len(reclamantes)} planilha(s) de cálculo extraída(s).")
    if paginas_sem_texto:
        print(f"  ATENÇÃO: {len(paginas_sem_texto)} página(s) sem texto "
              "extraível — provavelmente escaneadas; rode OCR nelas.")

    print("[2/4] Lendo vínculos do CAGED...")
    vinculos = extrair_vinculos_caged(args.caged)
    print(f"  -> {len(vinculos)} vínculo(s) extraído(s).")

    print("[3/4] Cruzando dados e verificando inconsistências...")
    inconsistencias = verificar(reclamantes, vinculos,
                                tolerancia_salario=args.tolerancia_salario)
    print(f"  -> {len(inconsistencias)} inconsistência(s) encontrada(s).")

    print("[4/4] Gerando relatórios...")
    gerar_relatorios(reclamantes, vinculos, inconsistencias,
                     paginas_sem_texto, args.saida, saida_txt)
    print(f"  -> Excel : {args.saida}")
    print(f"  -> Texto : {saida_txt}")

    criticas = sum(1 for i in inconsistencias if i.gravidade == "CRÍTICA")
    altas = sum(1 for i in inconsistencias if i.gravidade == "ALTA")
    print(f"\nRESUMO: {criticas} crítica(s), {altas} alta(s), "
          f"{len(inconsistencias) - criticas - altas} demais.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
