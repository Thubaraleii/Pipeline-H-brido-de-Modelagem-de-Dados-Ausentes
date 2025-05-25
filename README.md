# 🧠🔬 Pipeline Híbrido de Modelagem de Dados Ausentes — Paleoecologia Multiproxy (Lontras Shale Lagerstätte)

---

## 📄 Descrição do Projeto

Este repositório implementa um **pipeline híbrido e automatizado** para **modelagem, validação e preenchimento de dados ausentes** em estudos paleoecológicos e geoquímicos multiproxy.

O pipeline combina de forma inteligente:

- ✅ **Machine Learning supervisionado** (Random Forest, Gradient Boosting, KNN, Regressões Linear e Ridge).
- ✅ **Deep Learning** (Redes Neurais Multi-Layer Perceptron - MLP).
- ✅ **Modelagem Probabilística Bayesiana** (`Bayesian Ridge` e `PyMC`).
- ✅ **Bootstrap e Data Augmentation** (estatística robusta).
- ✅ Análise de correlação, matriz de dependência e comportamento estratigráfico.

✔️ O sistema realiza **seleção automática do melhor modelo para cada variável**, utilizando validação cruzada (LOOCV ou Holdout) e gera:

- 📊 **Tabelas preenchidas completas.**
- 📈 **Perfis estratigráficos, scatter, matriz de correlação e gráficos interativos.**
- 📦 Exportação em Excel, CSV e dashboards locais.

---
## 🗺️ Fluxo do Pipeline
📥 Dados → 🧽 Limpeza → 🔗 Data Augmentation

↓
🧠 Modelagem:

→ Random Forest

→ Gradient Boosting

→ KNN

→ Bootstrap (baseline)

→ Deep Learning (MLP)

→ Bayesian Ridge

→ PyMC (Probabilístico)&#x20;

↓
📊 Validação (LOOCV ou Holdout)

↓
🏆 Seleção do melhor modelo por variável

↓
📝 Preenchimento dos dados vazios

↓
📦 Exportação (Excel, CSV, Gráficos, Dashboard Local)

---

## 🔥 Funcionalidades

- ✔️ Preenchimento de dados geoquímicos e paleoecológicos multiproxy.
- ✔️ Escolha automática do melhor modelo para cada variável.
- ✔️ Modelagem híbrida: **ML + Deep Learning + Bayesian.**
- ✔️ Análise de incerteza estatística (modelos probabilísticos).
- ✔️ Dashboards locais interativos via **Plotly + Dash**.
- ✔️ Visualização tabular e wrangling via **Data Wrangler local**.

---

## 🛠️ Instalação

pipeline-paleoecologia/
│
├── data/                  # Dados de entrada
├── outputs/               # Resultados gerados
├── pipeline_hibrido.py    # Código principal do pipeline
├── dashboard.py           # Dashboard interativo
├── fluxograma.dot         # Arquivo do fluxograma Graphviz
├── pipeline.png           # Imagem do pipeline
├── requirements.txt       # Dependências
├── README.md              # Este arquivo

### ✅ 1. Clone este repositório

```bash
git clone https://github.com/seuusuario/pipeline-paleoecologia.git
cd pipeline-paleoecologia.
