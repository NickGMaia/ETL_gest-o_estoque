# 📦 ETL de Gestão de Estoque com Geração de Relatório em PowerPoint

Pipeline automatizado de análise de movimentação de estoque que lê dados de um CSV, processa indicadores de negócio, gera gráficos e exporta um relatório executivo em `.pptx`.

---

## 🗂️ Visão Geral

Este projeto implementa um fluxo ETL completo (Extract → Transform → Load) voltado para gestão de estoque de múltiplas filiais. A partir de um arquivo CSV de movimentações, o script:

1. **Extrai** os dados brutos de entradas e saídas de estoque
2. **Transforma** os dados em indicadores de negócio (faturamento, ticket médio, saldo de estoque, alertas críticos)
3. **Carrega** os resultados em gráficos `.png` e consolida tudo em um relatório PowerPoint `.pptx`

---

## ⚙️ Funcionalidades

- ✅ Cálculo de **faturamento total** por filial e por categoria
- ✅ Geração de **saldo atual de estoque** por produto e filial
- ✅ Identificação de **erros de registro** (saldos negativos)
- ✅ **Alertas automáticos** de reposição para itens com estoque ≤ 10 unidades
- ✅ Cálculo de **ticket médio** por operação e por filial
- ✅ Análise de **evolução de vendas diárias**
- ✅ Geração de **3 gráficos** (barras, linha e pizza)
- ✅ Exportação de **relatório executivo em PowerPoint**

---

## 📊 Gráficos Gerados

| Arquivo | Tipo | Conteúdo |
|---|---|---|
| `grafico_ticket_medio.png` | Barras | Ticket médio por operação em cada filial |
| `vendas_diarias.png` | Linha | Evolução do faturamento diário — Abril 2026 |
| `participacao_categorias.png` | Pizza | Participação de cada categoria no faturamento total |

---

## 🗃️ Estrutura Esperada do CSV

O arquivo `movimentacao_estoque.csv` deve conter as seguintes colunas:

| Coluna | Tipo | Descrição |
|---|---|---|
| `ID_Empresa` | string | Identificador da filial |
| `Data` | date | Data da movimentação |
| `Tipo` | string | `"Entrada"` ou `"Saída"` |
| `Produto` | string | Nome do produto |
| `Categoria` | string | Categoria do produto |
| `Quantidade` | int | Quantidade movimentada |
| `Valor_Unitario` | float | Valor unitário do produto |

---

## 🚀 Como Executar

### 1. Clone o repositório

```bash
git clone https://github.com/NickGMaia/ETL_gest-o_estoque.git
cd seu-repositorio
```

### 2. Instale as dependências

```bash
pip install pandas numpy matplotlib python-pptx
```

### 3. Adicione o arquivo de dados

Coloque o arquivo `movimentacao_estoque.csv` na raiz do projeto.

### 4. Execute o script

```bash
python etl_estoque.py
```

### 5. Saídas geradas

Após a execução, os seguintes arquivos serão criados na raiz do projeto:

```
grafico_ticket_medio.png
vendas_diarias.png
participacao_categorias.png
Relatorio_Executivo_Estoque.pptx
```

---

## 📁 Estrutura do Projeto

```
.
├── etl_estoque.py                  # Script principal do ETL
├── movimentacao_estoque.csv        # Base de dados de movimentações (não versionado)
├── grafico_ticket_medio.png        # Gráfico gerado
├── vendas_diarias.png              # Gráfico gerado
├── participacao_categorias.png     # Gráfico gerado
├── Relatorio_Executivo_Estoque.pptx  # Relatório final gerado
└── README.md
```

> ⚠️ Recomenda-se adicionar `*.csv`, `*.png` e `*.pptx` ao `.gitignore` para não versionar dados e artefatos gerados.

---

## 📋 Slides do Relatório PowerPoint

O arquivo `.pptx` gerado contém os seguintes slides:

1. **Capa** — Título e subtítulo do relatório
2. **Faturamento por Filial** — Gráfico de ticket médio por filial
3. **Participação por Venda** — Gráfico de pizza por categoria
4. **Vendas Diárias** — Gráfico de evolução temporal do faturamento
5. **Alertas de Estoque Crítico** — Tabela com produtos que precisam de reposição

---

## 🛠️ Tecnologias Utilizadas

| Biblioteca | Finalidade |
|---|---|
| `pandas` | Manipulação e análise dos dados |
| `numpy` | Operações vetorizadas (cálculo de variação de estoque) |
| `matplotlib` | Geração dos gráficos |
| `python-pptx` | Criação do relatório PowerPoint |

---

## 💡 Possíveis Melhorias Futuras

- [ ] Parametrizar o período de análise via argumentos de linha de comando
- [ ] Adicionar suporte a múltiplos arquivos CSV (processamento em lote)
- [ ] Implementar dashboard interativo com Plotly ou Streamlit
- [ ] Conectar a um banco de dados relacional como fonte de dados
- [ ] Envio automático do relatório por e-mail ao final da execução
- [ ] Adicionar testes unitários para as transformações

---

## 📄 Licença

Este projeto está sob a licença MIT. Consulte o arquivo [LICENSE](LICENSE) para mais detalhes.
