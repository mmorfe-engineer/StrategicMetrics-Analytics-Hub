# 📊 StrategicMetrics Analytics Hub

![Status](https://img.shields.io/badge/status-active-brightgreen)
![License](https://img.shields.io/badge/license-MIT-blue)
![Built%20with](https://img.shields.io/badge/built%20with-Excel%20%7C%20Power%20BI%20%7C%20Python-orange)
![Focus](https://img.shields.io/badge/focus-corporate%20analytics-%2300b894)

A professional analytics workspace for designing, auditing, and visualizing corporate statistical dashboards. 
This repository showcases end-to-end data analysis workflows, from raw transactional data to executive-level KPI dashboards.

---

## 🚀 Project Overview

**StrategicMetrics Analytics Hub** is a curated collection of corporate analytics assets, focused on:

- Transforming raw transactional data into clean, reusable analytical datasets.
- Applying descriptive statistics and exploratory analysis to understand business dynamics.
- Building interactive dashboards that support *real* decision-making (not just pretty charts).
- Embedding **auditing logic** to validate KPIs and detect inconsistencies between tools or models.

The flagship case study in this hub is **"La Tiendita UVM"** - a complete analytics pipeline for a real cooperative business at Universidad Valle del Momboy, including sales by branch, product portfolio (Pareto ABC analysis), payment channels, temporal behavior, inventory rotation, pricing structure, and customer satisfaction analysis.

---

## 🧱 Repository Structure

```bash
StrategicMetrics-Analytics-Hub/
│
├── data/                 # Sample or anonymized datasets (no raw PII)
├── notebooks/            # Exploratory analysis & prototyping
├── dashboards/           # Excel / Power BI / Tableau dashboards
│   └── La_Tiendita_UVM/  # Complete case study: 8 integrated modules
│       ├── Modulo1_VentasGlobales.xlsx
│       ├── Modulo2_AnalisisSucursales.xlsx
│       ├── Modulo3_PortafolioProductos.xlsx
│       ├── Modulo4_ComportamientoTemporal.xlsx
│       ├── Modulo5_InventarioRotacion.xlsx
│       ├── Modulo6_PreciosMargenes.xlsx
│       ├── Modulo7_SatisfaccionClientes.xlsx
│       └── Modulo8_Dashboard_TienditaUVM.xlsx
├── src/                  # Reusable analysis & audit scripts
│   └── La_Tiendita_UVM/
│       ├── generar_modulo1.py
│       ├── generar_modulo2.py
│       ├── generar_modulo3.py
│       ├── generar_modulo4.py
│       ├── generar_modulo5.py
│       ├── generar_modulo6.py
│       ├── generar_modulo7.py
│       └── generar_modulo8.py
├── docs/                 # Methodology notes, playbooks, and guides
│   └── PROMPT  VIBE CODE MÓDULO.txt
├── tests/                # Audit & consistency checks
├── assets/               # Images (charts, dashboard screenshots, logos)
├── README.md             # You are here
└── LICENSE
```

Each subfolder is designed to be **plug-and-play** for new corporate analytics projects: just drop in your data, wire the KPIs, and reuse the same audit and visualization patterns.

---

## 📈 Key Features

- **Corporate KPI Dashboards**  
  Ready-to-adapt templates for sales, branches, product categories, inventory, and payment channels.

- **Statistics-First Approach**  
  Descriptive statistics (mean, median, mode, dispersion, frequency tables, time series) baked into every dashboard, not added as an afterthought.

- **Audit-Friendly Design**  
  Every visual has a backing table, clear formulas, and optional "audit scripts" to cross-check results from different tools (e.g., Excel vs. Python vs. BI engine).

- **Modular Architecture**  
  One sheet / module per business dimension (e.g., global sales, branches comparison, payment mix), plus an integrated dashboard layer that consolidates insights.

- **Education-Ready**  
  Ideal for teaching or learning corporate analytics, statistics, dashboard design, and documentation best practices. The **La Tiendita UVM** case study serves as a comprehensive example.

---

## 🧪 Tech Stack

- **Data & Analysis**  
  - Microsoft Excel (Pivot Tables, Advanced Formulas, Data Validation) - **Primary tool**
  - Power BI (Interactive Dashboards, DAX, Power Query) - **Primary tool**
  - Python (Pandas, openpyxl) - For audit scripts and automation
  - CSV / TXT for transactional data import

- **Visualization & Dashboards**  
  - Microsoft Excel (Built-in Charts, Conditional Formatting)
  - Power BI (Interactive Visualizations, Drill-through)
  - Tableau (Optional for advanced visualizations)

- **Version Control & Collaboration**  
  - Git & GitHub for version control
  - GitHub Actions for CI/CD (Optional QA checks)

---

## 🧠 Core Concepts

This hub emphasizes:

- **Traceability:** every KPI can be traced back to a precise dataset, filter, and formula.
- **Reproducibility:** scripts can fully rebuild dashboards from raw data.
- **Auditability:** independent "audit modules" re-compute critical metrics and highlight any mismatches.
- **Decision Orientation:** visuals are annotated with concise insights, not left to interpretation.

The **La Tiendita UVM** case study demonstrates all these principles through 8 integrated modules that consolidate into a comprehensive executive dashboard.

---

## 📊 Flagship Case Study: La Tiendita UVM

**Project:** Radiografía Estadística de La Tiendita UVM  
**Institution:** Universidad Valle del Momboy - Facultad de Ciencias Económicas y Sociales  
**Course:** Estadística Descriptiva - Período 2026B  
**Professor:** Marilyn Briceño  
**Technical Lead:** Martin Morfe

### Project Overview

La Tiendita UVM is a student cooperative that operates across two branches (Estovacuy and Valera), offering a diverse product portfolio including stationery, laser-engraved items, vinyl products, and sublimation merchandise. This comprehensive analytics project integrates 8 modules of statistical analysis into a unified executive dashboard.

### Modules Structure

| Module | Focus Area | Key Deliverables |
|--------|-----------|------------------|
| **Módulo 1** | Global Sales | Total revenue, transaction counts, currency mix (USD/BS) |
| **Módulo 2** | Branch Analysis | Comparative analysis: Estovacuy vs Valera, ticket averages, dispersion |
| **Módulo 3** | Product Portfolio | Pareto ABC classification, product ranking, sales concentration |
| **Módulo 4** | Temporal Behavior | Weekly/monthly trends, seasonality, peak identification |
| **Módulo 5** | Inventory Rotation | Stock classification (High/Medium/Low), turnover rates |
| **Módulo 6** | Pricing Structure | Price ranges, distribution, median analysis, existence correlation |
| **Módulo 7** | Customer Satisfaction | Likert scale survey (32 respondents, 16 items, 8 dimensions) |
| **Módulo 8** | **Integrated Dashboard** | Executive synthesis with 6 analysis blocks, KPIs, recommendations |

### Integrated Dashboard (Módulo 8) Structure

The executive dashboard (Modulo8_Dashboard_TienditaUVM.xlsx) contains **13 sheets**:

1. **Portada_M8** - Project title, period, methodology, executive summary
2. **Modulo_1 to Modulo_7** - Individual module summaries with KPIs and backing tables
3. **Dashboard_Integrado_M8** - 6 analysis blocks:
   - Global Business Vision
   - Sales by Branch and Category
   - Payment Channels & Revenue
   - Temporal Sales Behavior
   - Integrated KPIs & Objectives
   - Audit Notes (vibe code validation)
4. **Conclusiones_M8** - Key findings and actionable recommendations
5. **Lecciones_M8** - Methodological and management learnings
6. **Glosario_M8** - Technical terms and acronyms dictionary
7. **Validacion_M8** - Internal audit notes and discrepancies

### Key Findings (La Tiendita UVM)

- **Total Revenue:** 5,200.62 USD across 285 transactions
- **Branch Performance:** Valera leads in revenue (57.8%) while Estovacuy has higher transaction volume (51.6%)
- **Payment Mix:** Transfers (49.8%), Cash (21.9%), AirTM (18.8%), Binance (5.1%), Plantilla (4.4%)
- **Product Concentration:** 5 Class A products generate 80.15% of sales (Pareto principle confirmed)
- **Price Distribution:** Median price of 15.00 USD, high dispersion (CV: 124.44%)
- **Customer Satisfaction:** Overall score of 4.72/5.0, with Identity (4.94) as strongest dimension

---

## 📥 Getting Started

### 1️⃣ Clone the Repository

```bash
git clone https://github.com/mmorfe-engineer/StrategicMetrics-Analytics-Hub.git
cd StrategicMetrics-Analytics-Hub
```

### 2️⃣ Explore the Case Study

Start with the **La Tiendita UVM** case study in `/dashboards/La_Tiendita_UVM/`:

1. Open `Modulo8_Dashboard_TienditaUVM.xlsx` for the integrated executive dashboard
2. Review individual modules (Modulo1-7) for detailed analysis
3. Check `/src/La_Tiendita_UVM/` for Python audit scripts
4. Read `/docs/PROMPT VIBE CODE MÓDULO.txt` for the original prompts and methodology

### 3️⃣ Adapt to Your Project

To use this hub for your own analytics project:

```bash
# Create your project folder structure
mkdir -p dashboards/Your_Project
mkdir -p src/Your_Project
mkdir -p data/Your_Project

# Add your files
# - Excel dashboards in dashboards/Your_Project/
# - Python scripts in src/Your_Project/
# - Raw data in data/Your_Project/
```

---

## ✅ Best Practices & Contributions

This hub follows industry-inspired best practices for data science and analytics projects:

- Clear separation between **data**, **code**, **dashboards**, and **documentation**
- Small, frequent commits with meaningful messages
- Issues and pull requests used to track improvements and bug fixes
- Every KPI backed by a supporting table and audit trail
- Standardized naming conventions across all modules

**Contributions are welcome!**  
If you want to:

- Add a new dashboard template,
- Contribute an audit module,
- Propose improvements to documentation or visual design,
- Add another case study,

please open an issue or submit a pull request.

---

## 📜 License

This project is released under the **MIT License**.  
You are free to use, modify, and adapt the templates for educational or corporate purposes, subject to the terms in `LICENSE`.

---

## 👤 Author & Maintainer

**Martin Morfe**  
📧 GitHub: [@morfemartin](https://github.com/morfemartin)  
🎓 Universidad Valle del Momboy - Contaduría Pública  
📍 Caracas, Venezuela

All contributions, code, documentation, and analysis in this repository are authored and maintained by **Martin Morfe**. The **La Tiendita UVM** case study was developed as part of the Statistical Project for the Descriptive Statistics course (2026B) at Universidad Valle del Momboy, under the guidance of Professor Marilyn Briceño.

---

## 🌟 Support & Inspiration

If you find this repository useful:

- ⭐ **Star** the repo on GitHub to help others discover it
- 🧑‍💻 **Fork** it and adapt the templates to your own corporate analytics projects
- 💬 **Open an issue** to share feedback or request new templates

The **La Tiendita UVM** case study demonstrates how academic projects can achieve professional-grade analytics standards. Use it as inspiration for your own data-driven decision support systems.

Let **StrategicMetrics Analytics Hub** be the backbone of your next data-driven decision.
