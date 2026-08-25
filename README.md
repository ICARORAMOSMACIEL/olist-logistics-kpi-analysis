# 📊 Olist E-Commerce - Análise de Performance Logística & SLA

Projeto end-to-end de análise de dados e inteligência logística utilizando o banco de dados de e-commerce da Olist. O objetivo é mapear os principais gargalos na cadeia de distribuição, calcular SLAs de entrega e mensurar o impacto direto dos atrasos na satisfação do cliente (NPS Proxy).

---

## 🛠️ Tecnologias Utilizadas
* **Banco de Dados:** MySQL
* **Linguagem & Bibliotecas:** Python (Pandas, SQLAlchemy, Seaborn, Matplotlib, OpenPyXL)
* **Visualização & Relatórios:** Planilhas executivas automatizadas (.xlsx)

---

## 📌 Principais Insights de Negócio
* **Impacto no Nível de Serviço:** Pedidos entregues no prazo sustentam nota média de **4.26 / 5.0**. Quando ocorre atraso, a avaliação desaba para **2.55 / 5.0** (queda de ~40%), demonstrando que o cumprimento de prazo é o principal fator de retenção do cliente.
* **Disparidade Regional:** Concentração de mais de 75% da demanda no Sul e Sudeste (Lead Time médio entre 8 e 14 dias). Regiões Norte e Nordeste apresentam Lead Times superiores a 25 dias e custos de frete até 2x maiores.
* **Gargalos Operacionais:** Estados com maior taxa de estouro de SLA (>15% de atraso): RR, AL, AP, SE, MA, PB, PA e PI.

---

## 🏗️ Estrutura do Pipeline
1. **Ingestão:** Carga automatizada dos conjuntos de dados transacionais via Python/SQLAlchemy para contornar limitações de importação em larga escala.
2. **Modelagem SQL:** Criação da view analítica `vw_kpis_logisticos` calculando Lead Time real, desvio em dias e status categórico de SLA.
3. **Análise & Visualização:** Análise exploratória com Pandas e geração de gráficos de dispersão frete × prazo.
4. **Automação de Relatórios:** Geração de planilha gerencial multicamadas com OpenPyXL contendo KPIs em cards, formatação condicional e cálculos nativos.

---

## 🚀 Como Executar o Projeto

1. Clone o repositório:
```bash
git clone [https://github.com/SEU_USUARIO/NOME_DO_REPOSITORIO.git](https://github.com/SEU_USUARIO/NOME_DO_REPOSITORIO.git)