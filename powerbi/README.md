# Modelo do Power BI

Versão em Power BI do projeto, com os **mesmos dados fictícios** da planilha e do BI em Python. Aqui ficam as tabelas prontas para importar, os relacionamentos e as medidas DAX comentadas.

```bash
python bi_dados_industriais/gerar_bi.py   # gera as ordens e os estoques
python powerbi/gerar_modelo.py            # gera powerbi/dados/*.csv
```

Os CSVs usam `;` como separador e vírgula decimal, para abrir direto no Power BI em português.

## Modelo de dados (esquema estrela)

Tabelas de **fato** guardam os números (uma linha por mês, etapa ou ordem). Tabelas de **dimensão** são as réguas usadas para filtrar. Cada fato se liga às dimensões por uma chave, sempre de **um** (dimensão) para **muitos** (fato).

```
                     dCalendario (Mes)
          ┌──────────┬─────┴─────┬───────────┬──────────┐
     fFaturamento  fPerdas    fMeta       fOrdens    fEstoque
          │          │          │            │
       dSetor ──► dEtapa ───────┘        dCliente
```

| Tabela | Tipo | Grão (uma linha por) | Chave |
|---|---|---|---|
| `dCalendario` | dimensão | mês | `Mes` |
| `dSetor` | dimensão | setor (4) | `Setor` |
| `dEtapa` | dimensão | etapa (7) | `IdEtapa`; ligada a `dSetor` por `Setor` |
| `dCliente` | dimensão | cliente (60) | `Cliente` |
| `fFaturamento` | fato | mês × setor | `Mes`, `Setor` |
| `fPerdas` | fato | mês × etapa | `Mes`, `IdEtapa` |
| `fMeta` | fato | mês × etapa (meta vigente no mês) | `Mes`, `IdEtapa` |
| `fOrdens` | fato | ordem de produção (2026, G1/G2) | `Mes`, `Cliente` |
| `fEstoque` | fato | mês (estoque no fim do mês) | `Mes` |

**Por que existe `dSetor`?** O faturamento vem por setor e a perda vem por etapa. `dSetor` é a ponte: filtrar um setor filtra as etapas dele (`dSetor → dEtapa → fPerdas`) e o faturamento dele (`dSetor → fFaturamento`).

**Por que a meta já vem resolvida por mês?** A meta pode mudar no meio do período. O script já grava qual meta valia em cada mês, então a DAX só precisa somar.

## Montagem (passo a passo)

1. **Obter dados → Texto/CSV**, importar os 9 arquivos de `powerbi/dados/`. Conferir os tipos: `Mes` = Data; colunas `_t` = Número decimal.
2. **Exibição de modelo:** criar os relacionamentos (todos *um para muitos*, filtro em direção única):
   - `dCalendario[Mes]` → `fFaturamento`, `fPerdas`, `fMeta`, `fOrdens`, `fEstoque`
   - `dSetor[Setor]` → `fFaturamento[Setor]` e `dEtapa[Setor]`
   - `dEtapa[IdEtapa]` → `fPerdas[IdEtapa]` e `fMeta[IdEtapa]`
   - `dCliente[Cliente]` → `fOrdens[Cliente]`
3. Em `dCalendario`, **Classificar `MesAno` por `IdxMes`** (senão os meses saem em ordem alfabética).
4. Criar uma tabela vazia `_Medidas` e colar as medidas abaixo.

## Medidas DAX

### Base

```dax
Faturamento (t) = SUM ( fFaturamento[Faturamento_t] )

Perda (t) = SUM ( fPerdas[Perda_t] )

Meta (t) = SUM ( fMeta[Meta_t] )

Produção (t) = SUM ( fOrdens[Peso_t] )
```

### % de perda (indicador principal)

```dax
-- Faturamento do setor da etapa que está no filtro.
-- Ao olhar uma etapa (ex.: G1 Acabamento), dEtapa não filtra o faturamento sozinho;
-- TREATAS "copia" os setores das etapas visíveis para o filtro de dSetor.
Faturamento do setor (t) =
CALCULATE ( [Faturamento (t)], TREATAS ( VALUES ( dEtapa[Setor] ), dSetor[Setor] ) )

% de perda = DIVIDE ( [Perda (t)], [Faturamento do setor (t)] )
```

`DIVIDE` em vez de `/` evita erro quando o faturamento é zero ou vazio (devolve vazio).

### Comparação com o mês anterior

```dax
-- Usa IdxMes (1, 2, 3...) para achar o mês anterior sem depender de tabela de datas diárias.
% de perda mês anterior =
VAR MesAtual = MAX ( dCalendario[IdxMes] )
RETURN
    CALCULATE ( [% de perda], REMOVEFILTERS ( dCalendario ), dCalendario[IdxMes] = MesAtual - 1 )

Variação (pontos) = ( [% de perda] - [% de perda mês anterior] ) * 100

Seta vs mês anterior =
SWITCH ( TRUE (),
    ISBLANK ( [% de perda mês anterior] ), BLANK (),
    [% de perda] < [% de perda mês anterior], "▼ melhorou",
    [% de perda] > [% de perda mês anterior], "▲ piorou",
    "= igual" )
```

### Comparação com a meta

```dax
-- A meta é o limite máximo de perda. Âmbar = abaixo da meta, mas a menos de 1% do limite.
Situação na meta =
VAR Razao = DIVIDE ( [Perda (t)], [Meta (t)] )
RETURN
    SWITCH ( TRUE (),
        ISBLANK ( Razao ), BLANK (),
        Razao > 1, "Acima da meta",
        Razao >= 0.99, "No limite",
        "Abaixo da meta" )

-- Cor para formatação condicional (Formatar → Cor → Valor do campo).
Cor da meta =
SWITCH ( [Situação na meta],
    "Acima da meta", "#B3261E",
    "No limite", "#E3A008",
    "Abaixo da meta", "#1B7F3B" )
```

### Onde focar: Pareto e tendência

```dax
-- Quanto cada etapa pesa na perda total (ignora o filtro de etapa, mantém o de mês).
% do total da perda = DIVIDE ( [Perda (t)], CALCULATE ( [Perda (t)], ALL ( dEtapa ) ) )

-- % acumulado do Pareto: soma das etapas com perda maior ou igual à da etapa atual.
% acumulado =
VAR PerdaAtual = [Perda (t)]
VAR Acumulado =
    CALCULATE ( [Perda (t)],
        FILTER ( ALL ( dEtapa ), [Perda (t)] >= PerdaAtual ) )
RETURN DIVIDE ( Acumulado, CALCULATE ( [Perda (t)], ALL ( dEtapa ) ) )

-- Tendência: % de perda dos últimos 3 meses x 3 meses anteriores (até o último mês do filtro).
Tendência 3 meses =
VAR Ultimo = MAX ( dCalendario[IdxMes] )
VAR Ult3 = CALCULATE ( [% de perda], REMOVEFILTERS ( dCalendario ),
                       dCalendario[IdxMes] >= Ultimo - 2, dCalendario[IdxMes] <= Ultimo )
VAR Ant3 = CALCULATE ( [% de perda], REMOVEFILTERS ( dCalendario ),
                       dCalendario[IdxMes] >= Ultimo - 5, dCalendario[IdxMes] <= Ultimo - 3 )
RETURN DIVIDE ( Ult3, Ant3 ) - 1

Prioridade =
VAR PesaMuito = [% do total da perda] >= 0.10
VAR Piorando = [Tendência 3 meses] > 0
RETURN
    SWITCH ( TRUE (),
        ISBLANK ( [Tendência 3 meses] ), BLANK (),
        PesaMuito && Piorando, "Agir agora",
        PesaMuito, "Padronizar",
        Piorando, "Acompanhar",
        "Manter" )
```

### Balanço de peso (só G1 + G2)

```dax
Faturamento G1+G2 (t) = CALCULATE ( [Faturamento (t)], dSetor[Grupo] = "Matriz G1/G2" )

Perda G1+G2 (t) = CALCULATE ( [Perda (t)], dSetor[Grupo] = "Matriz G1/G2" )

Estoque produto pronto (t) = SUM ( fEstoque[ProdutoPronto_t] )

Variação do estoque (t) =
VAR MesAtual = MAX ( dCalendario[IdxMes] )
VAR Anterior = CALCULATE ( [Estoque produto pronto (t)], REMOVEFILTERS ( dCalendario ), dCalendario[IdxMes] = MesAtual - 1 )
RETURN IF ( ISBLANK ( Anterior ), BLANK (), [Estoque produto pronto (t)] - Anterior )

-- Produção = faturado + perdas + variação do estoque. O que sobra é "diferença sem explicação".
Diferença do balanço % =
DIVIDE ( [Produção (t)] - ( [Faturamento G1+G2 (t)] + [Perda G1+G2 (t)] + [Variação do estoque (t)] ), [Produção (t)] )

Alerta do balanço = IF ( ABS ( [Diferença do balanço %] ) > 0.05, "Investigar", "OK" )
```

*O estoque de janeiro não tem mês anterior no modelo, então a variação e o balanço começam em fevereiro.*

### Curva ABC (colunas calculadas em `dCliente`)

Colunas calculadas são fixas (calculadas uma vez, na atualização). Servem bem para a classe ABC do período inteiro.

```dax
Produção cliente (t) = CALCULATE ( SUM ( fOrdens[Peso_t] ) )

% acumulado antes =
VAR Atual = dCliente[Produção cliente (t)]
VAR Maiores = CALCULATE ( SUM ( fOrdens[Peso_t] ), FILTER ( ALL ( dCliente ), dCliente[Produção cliente (t)] > Atual ) )
RETURN DIVIDE ( Maiores, CALCULATE ( SUM ( fOrdens[Peso_t] ), ALL ( dCliente ) ) )

Classe ABC = IF ( dCliente[% acumulado antes] < 0.80, "A", IF ( dCliente[% acumulado antes] < 0.95, "B", "C" ) )
```

### Estoque de papel

```dax
-- Simplificação: consumo de papel = peso produzido; mês de 30 dias. É ESTIMATIVA.
Dias de papel = DIVIDE ( SUM ( fEstoque[PapelTotal_t] ), [Produção (t)] / 30 )

Dias de papel na fábrica = DIVIDE ( SUM ( fEstoque[PapelFabrica_t] ), [Produção (t)] / 30 )
```

## Páginas sugeridas

| Página | Visuais |
|---|---|
| **Visão geral** | Cartões: Faturamento, Perda, % de perda, Seta vs mês anterior. Colunas: faturamento por mês + linha de % de perda. Matriz: setor × mês com `% de perda` e cor por `Cor da meta`. Segmentação de mês. |
| **Onde focar** | Gráfico de colunas e linha (Pareto): `Perda (t)` por `dEtapa[NomeCurto]` (ordenado do maior para o menor) + linha `% acumulado`. Tabela: etapa, % do total, Tendência 3 meses, Prioridade. |
| **Balanço e estoques** | Colunas: Produção × Faturamento G1+G2. Colunas: Diferença do balanço % com linha constante em 5%. Linha: Dias de papel. |
| **Clientes** | Barras: 12 maiores clientes por produção, cor = Classe ABC. Cartões: número de clientes A, B e C. |

## Conferências

Depois de montar, os totais precisam bater com a planilha e o BI:

- Faturamento de G1 + G2 de jan a ago/2026 = **3.776,1 t**.
- Perdas de G1 + G2 de jan a ago/2026 = **305,4 t**.
- Diferença do balanço acima de 5% só em **jun/2026** (proposital nos dados fictícios).
