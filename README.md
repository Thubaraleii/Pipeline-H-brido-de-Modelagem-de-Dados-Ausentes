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

