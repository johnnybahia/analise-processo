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

### Modo automático (recomendado)

Coloque o PDF do processo e **todos** os PDFs de CAGED na mesma pasta do
script e rode **sem argumentos**:

```bash
python3 analisar_processo_caged.py
```

O script identifica cada PDF **pelo conteúdo** (não pelo nome): o que contém
planilhas do PJe-Calc é tratado como processo e os que contêm vínculos
(PIS/Nome) são tratados como CAGED — todos os CAGEDs são lidos e somados.
Para procurar em outra pasta: `--pasta /caminho/da/pasta`.

### Modo manual

```bash
python3 analisar_processo_caged.py PROCESSO.pdf CAGED1.pdf CAGED2.pdf ... \
    --pagina-inicial 351 --pagina-final 1142 \
    --saida relatorio_inconsistencias.xlsx
```

Os valores padrão de `--pagina-inicial` e `--pagina-final` já são 351 e 1142,
então normalmente basta rodar sem esses parâmetros.

Saídas geradas:

- `relatorio_inconsistencias.xlsx` — abas: **Inconsistências** (ordenadas por
  gravidade), **Reclamantes (Processo)**, **Vínculos (CAGED)** e, se houver,
  **Páginas sem texto** (páginas escaneadas que exigem OCR).
- `relatorio_inconsistencias.txt` — mesmo conteúdo em texto, agrupado por
  reclamante, para leitura rápida.

## Verificações realizadas

### Cruzamento com o CAGED

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

### Auditoria dos cálculos de liquidação (itens jurídicos)

Verificações espelhadas da auditoria da liquidação de sentença, para
conferência independente dos números:

| Item | Verificação | Gravidade |
|------|-------------|-----------|
| 1 | Cobrança **fora do título executivo** — competências anteriores a `--inicio-titulo` (ex.: CCT 2018/2019 quando a sentença cobre só 2019/2020 em diante) | CRÍTICA |
| 1/2 | **Competências prescritas** — lançamentos anteriores ao marco quinquenal (ajuizamento − 5 anos, calculado automaticamente da Data de Ajuizamento de cada planilha) | CRÍTICA |
| 2 | **Crédito integralmente prescrito** — demissão anterior ao marco quinquenal | CRÍTICA |
| 3a | **PLR sem proporcionalidade (avos)** — valor cheio (`--valor-plr-cheio`, padrão R$ 400,00) lançado em contrato parcial; calcula os avos (meses com ≥ 15 dias, Súmula 451/TST) e o valor proporcional | ALTA |
| 3b | **Multa acima do teto** — multa convencional corrigida superior à obrigação principal recomposta (art. 412 CC; OJ 54 SBDI-1/TST) | ALTA |
| 3c | **Fragmentação** — mesma competência da mesma verba lançada mais de uma vez na mesma planilha | ALTA |
| 4 | **Sem memória de cálculo individualizada** — nomes de `--lista-reclamantes` sem planilha no intervalo analisado | CRÍTICA |
| 5 | **Custas recalculadas por cálculo** — quando a sentença fixou custas em valor único (`--custas-fixas`); aponta cada planilha e quantifica o excesso total | ALTA/MÉDIA |

Cada inconsistência quantificável traz a coluna **Impacto Estimado (R$)**
(no valor corrigido usado pelo próprio cálculo, sem juros/honorários), e o
relatório em texto apresenta o **impacto total estimado** — compare com o
"total mínimo impugnável" de outras auditorias.

## Opções

| Opção | Padrão | Descrição |
|-------|--------|-----------|
| `--pasta` | `.` | Pasta onde procurar os PDFs no modo automático |
| `--pagina-inicial` | 351 | Primeira página do PDF do processo a ler |
| `--pagina-final` | 1142 | Última página do PDF do processo a ler |
| `--saida` | `relatorio_inconsistencias.xlsx` | Arquivo Excel de saída |
| `--tolerancia-salario` | 0.05 | Tolerância (5%) para apontar divergência salarial |
| `--valor-plr-cheio` | 400.00 | Valor cheio da PLR previsto na norma coletiva |
| `--custas-fixas` | (desligado) | Valor de custas fixado na sentença (ex.: `--custas-fixas 2000`) |
| `--inicio-titulo` | (desligado) | Data inicial coberta pelo título executivo, `DD/MM/AAAA` (ex.: início da vigência da CCT 2019/2020) |
| `--lista-reclamantes` | (desligado) | Arquivo texto com um nome por linha, para apontar quem não tem planilha individualizada |

Exemplo completo, espelhando a auditoria:

```bash
python3 analisar_processo_caged.py processo.pdf caged.pdf \
    --custas-fixas 2000 \
    --inicio-titulo 01/11/2019 \
    --lista-reclamantes reclamantes.txt
```

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
- **Premissa da PLR proporcional**: o ciclo é assumido como o **ano civil da
  competência** em que a PLR foi lançada. Se a norma coletiva definir ciclo
  diverso (ex.: novembro a outubro), os avos podem mudar — confira o ACT/CCT.
- Os impactos estimados podem **se sobrepor** (ex.: uma competência prescrita
  também pode estar fora do título): revise antes de somar na impugnação.
