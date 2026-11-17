# 📊 Dashboard de Controle de Ordens de Produção

Projeto de **Business Intelligence** em **Power BI** para acompanhar ordens de produção industriais: perdas por fase, balanceamento de massa, aging de ordens, pendências e assertividade de encerramento.

> **Todos os dados deste repositório são sintéticos.** Foram gerados pelo script [`gerar_planilhas_sinteticas.py`](gerar_planilhas_sinteticas.py) e não pertencem a nenhuma empresa, cliente, produto ou sistema real.

---

## 🎯 Objetivo

Transformar planilhas operacionais extensas e difíceis de interpretar em **indicadores claros** para engenharia de processos, PCP, custos e supervisão.

## 🧠 Contexto

Em indústrias, o controle de ordens de produção costuma viver em planilhas complexas que só quem as construiu consegue ler. Este projeto recria esse cenário com dados fictícios para demonstrar a modelagem e os indicadores que resolvem o problema.

---

## 📈 Indicadores

- **% de waste por fase** produtiva e **KG de perda por fase**
- **Top 5 perdas por máquina**
- **Balanceamento de ordens:** Perda BM (balanço de massa) x Perda apontada, com tolerância de ±10% (OK / NOK), ordens balanceadas x desbalanceadas e taxa de balanceamento
- **Assertividade mensal de encerramento** e ordens encerradas por mês
- **Aging:** ordens abertas, sem movimento e acima de 7 dias
- **Pendências** por tipo e status (embalar, revisão, refugar ou paletizar)
- **Quantidade de ordens:** semiacabado x acabado

## 🧮 Regras de negócio (simplificadas)

| Campo | Regra |
|---|---|
| Perda BM | Consumo − Produção |
| BM (%) | Perda BM ÷ Consumo |
| Diferença (%) | (Perda BM − Perda apontada) ÷ Produção |
| Status de balanceamento | **OK** se \|Diferença\| ≤ 10%; **NOK** caso contrário |

## 🔗 Modelo de dados

```text
Máquinas (OP)  ──1:1──┐
BI (GRUPO_ORDEM) ─1:1─┼──  COOIS (OP)   ← tabela central de ordens
Lista de pendências (OP) ─1:1─┘
Consolidado de encerramento (série mensal) e kg total pendente (resumo)
```

---

## 🗂️ Planilhas sintéticas

| Arquivo | Aba lida pelo Power BI | Conteúdo |
|---|---|---|
| `Controle de encerramento de OPs.xlsm` | `COOIS` | 1.580 ordens de produção |
| | `BI` | 416 linhas de produção, consumo e perdas por ordem |
| `Controle de produção_Maq.xlsx` | `Máquinas ` | 5.852 linhas de produção por máquina e turno |
| `Ordens pendentes.xlsx` | `Pendencia` | 745 pendências |
| | `kg total pendente ` | Resumo de consumo e produção pendentes |
| `Controle de encerramento Consolidado.xlsx` | `Ence. OPs (Consolidado)` | 18 meses de histórico |

A geração usa *seed* fixa, portanto é reproduzível:

```bash
pip install pandas numpy openpyxl
python gerar_planilhas_sinteticas.py
```

---

## ▶️ Como abrir o relatório

1. Baixe o repositório e abra `Indicador de Controle de OPs.pbix` no **Power BI Desktop**.
2. Se aparecer erro de arquivo não encontrado, vá em **Transformar dados → Configurações da fonte de dados → Alterar origem** e aponte as 4 planilhas para a pasta onde você baixou o projeto.
3. Clique em **Atualizar**.

---

## 🛠️ Ferramentas

- **Power BI** (modelagem, DAX, Power Query)
- **Excel** (fonte dos dados)
- **Python** (NumPy e openpyxl, para gerar a base sintética)

## 🖼️ Dashboard

### Visão geral
<img width="379" height="398" alt="Controle de OP&#39;s" src="https://github.com/user-attachments/assets/965c3169-c16f-4fbd-82d1-004a64ea2bb6" />


### Ordens pendentes
<img width="455" height="352" alt="Lista de ordens pendentes" src="https://github.com/user-attachments/assets/1755816f-e327-4025-abb7-115a8e039644" />


### Painel de balanço de massa
<img width="710" height="398" alt="Painel de balanço de massa" src="https://github.com/user-attachments/assets/4f62940c-4a4e-4cac-b18e-c347be592f1b" />


### Ordens encerradas
<img width="635" height="393" alt="indicadores de encerramento de ordens" src="https://github.com/user-attachments/assets/42d88415-36d9-4a5e-8ad2-2b42b6a5db85" />


---

## 📁 Estrutura

```text
dashboard-controle-ordens-producao/
├── images/
├── Controle de encerramento de OPs.xlsm
├── Controle de produção_Maq.xlsx
├── Ordens pendentes.xlsx
├── Controle de encerramento Consolidado.xlsx
├── Indicador de Controle de OPs.pbix
├── gerar_planilhas_sinteticas.py
└── README.md
```

Projeto desenvolvido para **portfólio**. Os dados são fictícios e as regras de negócio foram simplificadas.
