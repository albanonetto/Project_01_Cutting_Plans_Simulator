# Especificação Técnica e Análise do Algoritmo

Este documento detalha a implementação técnica do **Simulador de Planos de Corte**. A lógica foi desenvolvida em VBA e se baseia em cálculos geométricos e loops aninhados para explorar o maior número possível de combinações de corte.

### 1. Configuração do Ambiente e Entradas

A simulação é controlada a partir da planilha `DADOS DE ENTRADA`.

#### Entradas do Sistema

| Célula | Descrição | Variável VBA |
| :--- | :--- | :--- |
| `D4` | Diâmetro da tora (mm) | `DP` |
| `D5` | Diâmetro do grupo (não detalhado) | `DG` |
| `D8` | Comprimento da tora (mm) | `CT` |
| `D11` | Espessura da serra múltipla (mm) | `ES` |
| `D12` | Espessura da serra fita (mm) | `EF` |
| `C16:C40`| Largura verde do taboado (mm) | `intervalE` |
| `D16:D40`| Largura seca do taboado (mm) | `intervalLSEC` |
| `A16:A26`| Espessura verde do taboado (mm) | `intervalET` |
| `B16:B26`| Espessura seca do taboado (mm) | `intervalESEC` |

### 2. Análise Detalhada do Código (`CalcularValores_Otimizado`)

O código segue uma estrutura lógica para garantir desempenho e precisão.

#### 2.1. Otimizações e Tratamento de Erros

Antes de iniciar os cálculos, o ambiente do Excel é otimizado para alta performance.

```vba
' Melhorar o desempenho durante a execução
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False

' Garante que, mesmo em caso de erro, as configurações serão restauradas
On Error GoTo RestaurarConfiguracoes
```

#### 2.2. Leitura de Parâmetros e Setup

Os valores da planilha `DADOS DE ENTRADA` são carregados em variáveis e um novo worksheet para os resultados é criado com um timestamp único.

```vba
' Lendo parâmetros da tora
With ThisWorkbook.Sheets("DADOS DE ENTRADA")
    DP = .Range("D4").Value
    CT = .Range("D8").Value
    ES = .Range("D11").Value
    EF = .Range("D12").Value
    R = .Range("D4").Value / 2 ' Raio da tora
    DM = .Range("D4").Value / 5 ' Diâmetro da medula (estimado)
End With

' Criando a planilha de resultados
timestamp = Format(Now, "YYYYMMDD_HHMMSS")
Set resultWs = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
resultWs.Name = "RESULTADOS_" & timestamp
```

#### 2.3. Lógica Central de Simulação (Loops Aninhados)

O núcleo da simulação consiste em múltiplos loops aninhados que testam diferentes combinações de dimensões.

**Loop 1: Largura do Bloco Central (`E`)**

O primeiro loop itera sobre as larguras possíveis para os taboados do bloco central. A partir dessa largura (`E`), calcula-se a altura máxima (`L`) do bloco que pode ser inscrito no círculo da tora usando o Teorema de Pitágoras.

![Trigonometria Bloco Central](https://i.imgur.com/2Y4P1vR.png)

```vba
' L é a altura máxima do bloco central
L = 2 * Sqr(R ^ 2 - (E / 2) ^ 2)
```

**Loop 2: Espessura do Bloco Central (`ET`)**

Para cada largura, o segundo loop testa as diferentes espessuras de taboado (`ET`). Com isso, calcula-se o número de taboas (`N`) que cabem na altura `L`.

```vba
' Calcula o número de taboas (arredondado para baixo)
N = WorksheetFunction.Floor(L / ET, 1)

' Limita o número de peças por restrições da máquina
If N > 8 Then GoTo ProximoJ
If N = 0 Then GoTo ProximoI

' Recalcula a largura real do bloco (LBES), considerando a espessura das serras
LBES = ((N - 1) * ES) + (N * ET)

' Ajusta N se a largura com as serras exceder a altura L
If LBES > L Then
    N = N - 1
    LBES = ((N - 1) * ES) + (N * ET)
End If
```

**Loop 3: Cálculo da Costaneira (`LC`)**

A costaneira é a peça retirada da "lateral" do bloco central. Sua espessura máxima (`EC`) é calculada com base na largura (`LC`) escolhida, novamente usando trigonometria.

![Trigonometria Costaneira](https://i.imgur.com/r8k3pXG.png)

```vba
' Calcula a espessura máxima da costaneira
EC = Sqr(R ^ 2 - (LC / 2) ^ 2) - (E / 2) - EF

' Ajusta EC para a medida padrão mais próxima (para baixo)
Select Case EC
    Case Is < 17.85: EC = 0
    Case Is >= 80.85: EC = 80.85
    ' ... outras faixas
End Select
```

**Loop 4: Cálculo da Fresa (`LFR`)**

De forma similar, a fresa é a peça retirada da parte superior/inferior do bloco central. Sua espessura (`EFR`) depende da largura (`LFR`) e da dimensão já ocupada pelo bloco (`LBES`).

![Trigonometria Fresa](https://i.imgur.com/020fM8Q.png)

```vba
' Calcula a espessura máxima da fresa
EFR = Sqr(R ^ 2 - (LFR / 2) ^ 2) - (LBES / 2)

' Ajusta EFR para a medida padrão mais próxima
Select Case EFR
    Case Is < 18: EFR = 0
    ' ... outras faixas
End Select
```

#### 2.4. Cálculo de Volume e Aproveitamento

Após determinar as dimensões de todos os produtos possíveis, o volume total e o aproveitamento são calculados.

```vba
' Volumes são calculados com base nas dimensões SECAS e comprimento da tora
VB = N * (ESEC * LSEC * CT) ' Volume do Bloco
VC = 2 * (ECSEC * LCSEC * CT) ' Volume das 2 Costaneiras
VF = 2 * (EFRSEC * LFRSEC * CT) ' Volume das 2 Fresas

' Volume total de produtos
VT = VC + VB + VF 

' Volume da tora (cilindro)
VTORA = WorksheetFunction.Pi() * (RT ^ 2) * CT

' Aproveitamento percentual
A = (VT / VTORA) * 100
```

#### 2.5. Gravação dos Resultados

Finalmente, todos os resultados calculados para a combinação atual são gravados na planilha `RESULTADOS_xxx` e na `Base visualização`.

```vba
' Escreve os dados na linha correspondente da planilha de resultados
With resultWs
    .Cells(resultRow, 1).Value = resultRow
    .Cells(resultRow, 2).Value = DP
    .Cells(resultRow, 3).Value = VTORA / 1000000000 ' em m³
    .Cells(resultRow, 4).Value = ESEC & "X" & LSEC
    ' ... e assim por diante para todas as colunas
End With

resultRow = resultRow + 1
```

### 3. Geração da Visualização Gráfica

A planilha `Base visualização` contém fórmulas que geram coordenadas X e Y para um gráfico de dispersão, desenhando o plano de corte.

#### 3.1. Desenho da Tora (Círculo)

O diâmetro da tora é desenhado calculando-se 22 pontos em um círculo usando seno e cosseno.

$$
x_i = \frac{D}{2} \times \cos(\theta_i) \quad | \quad y_i = \frac{D}{2} \times \sin(\theta_i) \quad \text{onde} \quad \theta_i = i \times \frac{2\pi}{21}
$$

#### 3.2. Desenho Dinâmico dos Taboados

Os retângulos (taboados) são desenhados dinamicamente. Cada retângulo requer 5 coordenadas (X, Y) para fechar seu contorno. Fórmulas de Excel com a função `SE` (`IF`) são usadas para desenhar um taboado apenas se o plano de corte o incluir.

**Exemplo de fórmula para a coordenada X do primeiro ponto do segundo taboado:**

```excel
=SE($J$1>1; $H$1/2; 0)
```
- `$J$1`: Número total de taboados no bloco central (`N`).
- `$H$1`: Largura do taboado (`E`).

Se `N` for maior que 1, a coordenada X é calculada; senão, é zero, e o retângulo não aparece no gráfico. Lógicas similares são aplicadas para as coordenadas Y e para as costaneiras e fresas.

### 4. Conclusão Estratégica

A estrutura do algoritmo permite uma análise completa e rápida de milhares de cenários. Ao combinar a simulação com a visualização gráfica e a exportação de dados, a ferramenta se torna um poderoso suporte à decisão, permitindo que o planejamento da serraria otimize o uso de matéria-prima, reduza custos e responda de forma ágil às demandas de produção.
