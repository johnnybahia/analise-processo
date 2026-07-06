# Análise de Inconsistências: Processo Trabalhista x CAGED

Script Python que lê as **planilhas de cálculo do PJe-Calc** contidas no PDF do
processo (páginas 351 a 1142, configurável), extrai os dados de cada
reclamante, cruza com os **vínculos do extrato CAGED** (PDF) e gera um
relatório minucioso de inconsistências em **Excel** e **texto**.

## Instalação

```bash
pip install -r requirements.txt
```

## Uso

```bash
python3 analisar_processo_caged.py PROCESSO.pdf CAGED.pdf \
    --pagina-inicial 351 --pagina-final 1142 \
    --saida relatorio_inconsistencias.xlsx
```

Os valores padrão de `--pagina-inicial` e `--pagina-final` já são 351 e 1142,
então para o processo completo basta:

```bash
python3 analisar_processo_caged.py processo_completo.pdf caged_completo.pdf
```

Saídas geradas:

- `relatorio_inconsistencias.xlsx` — abas: **Inconsistências** (ordenadas por
  gravidade), **Reclamantes (Processo)**, **Vínculos (CAGED)** e, se houver,
  **Páginas sem texto** (páginas escaneadas que exigem OCR).
- `relatorio_inconsistencias.txt` — mesmo conteúdo em texto, agrupado por
  reclamante, para leitura rápida.

## Verificações realizadas

| # | Verificação | Gravidade |
|---|-------------|-----------|
| 1 | Reclamante sem vínculo localizado no CAGED | CRÍTICA |
| 2 | Período do cálculo inicia **antes** da admissão registrada no CAGED | CRÍTICA |
| 3 | Período do cálculo termina **após** o desligamento registrado no CAGED | CRÍTICA |
| 4 | Data de admissão divergente entre processo e CAGED | ALTA/MÉDIA |
| 5 | Data de demissão divergente entre processo e CAGED | ALTA/MÉDIA |
| 6 | Mês cobrado no cálculo com remuneração **zerada** no CAGED | ALTA |
| 7 | Reclamante com mais de uma planilha de cálculo (duplicidade) | ALTA |
| 8 | Período do cálculo incompatível com admissão/demissão da própria planilha | ALTA |
| 9 | Salário base do cálculo divergente da remuneração do CAGED (> tolerância) | MÉDIA |
| 10 | Soma das verbas divergente do total do resumo da planilha | MÉDIA |
| 11 | Nome com correspondência apenas aproximada (grafia divergente) | MÉDIA |
| 12 | Páginas sem texto extraível (escaneadas — exigem OCR) | aviso |

## Opções

| Opção | Padrão | Descrição |
|-------|--------|-----------|
| `--pagina-inicial` | 351 | Primeira página do PDF do processo a ler |
| `--pagina-final` | 1142 | Última página do PDF do processo a ler |
| `--saida` | `relatorio_inconsistencias.xlsx` | Arquivo Excel de saída |
| `--tolerancia-salario` | 0.05 | Tolerância (5%) para apontar divergência salarial |

## Observações importantes

- A correspondência entre reclamante e vínculo é feita **por nome**
  (normalizado, sem acentos). Nomes parecidos mas não idênticos são apontados
  como "NOME APENAS APROXIMADO" para conferência manual por CPF/PIS.
- Quando há **mais de um vínculo** para o mesmo nome (readmissão), o script
  escolhe o vínculo com maior sobreposição com o período do cálculo.
- A remuneração do CAGED pode incluir horas extras e adicionais; por isso a
  divergência salarial usa tolerância configurável e é classificada como MÉDIA.
- Se o relatório apontar páginas sem texto extraível, essas páginas são
  imagens escaneadas: rode OCR (ex.: `ocrmypdf -l por entrada.pdf saida.pdf`)
  e reprocesse.
- O relatório é uma **triagem automática**: confira as inconsistências
  apontadas no documento original antes de usá-las na defesa (a coluna
  "Localização no Processo" indica as folhas/páginas de cada planilha).
