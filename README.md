# DIAS – Design Informatics Analysis System  
  
[![Python 3.12](https://img.shields.io/badge/Python-3.12-blue.svg)](https://www.python.org/downloads/release/python-3120/)  
[![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)  
  
**DIAS** is an open-source, Python-based statistical analysis system built for design research. It lets design scholars run rigorous quantitative analyses without formal training in statistics or programming by turning statistical reasoning into an executable, rule-based decision system.  
  
> **77 statistical methods** · Automated parameter configuration · Bilingual UI (EN / 中文) · Word reports & publication-ready charts  
  
---  
  
## Table of Contents  
  
- [Why DIAS?](#why-dias)  
- [Key Features](#key-features)  
- [Supported Methods](#supported-methods)  
- [Architecture](#architecture)  
- [Getting Started](#getting-started)  
- [Usage](#usage)  
- [Resources](#resources)  
- [Tech Stack](#tech-stack)  
- [Contributing](#contributing)  
- [License](#license)  
- [Contact](#contact)  
  
---  
  
## Why DIAS?  
  
Design research increasingly demands structured quantitative methods, yet three barriers persist:  
  
1. **Method selection** — choosing the right statistical test is hard without formal training.  
2. **Parameter configuration** — setting up tests correctly is error-prone.  
3. **Workflow integration** — most statistics tools are disconnected from design research workflows.  
  
DIAS solves all three by embedding domain-specific heuristics, automatic diagnostics, and end-to-end report generation into a single desktop application.  
  
---  
  
## Key Features  
  
### Design-Oriented Statistical Taxonomy  
  
Methods are organized by **design research logic** rather than mathematical taxonomy:  
  
| Category | Examples |  
|---|---|  
| Data Description & Validation | Descriptive Statistics, Reliability, Validity |  
| Questionnaire Analysis | KANO Model, NPS, Content Validity, Delphi Method |  
| Correlation Analysis | Pearson, Spearman, Partial Correlation, Canonical Correlation |  
| Difference Analysis | T-Tests, ANOVA, MANOVA, Wilcoxon, Friedman |  
| Multi-Criteria Decision-Making | AHP, FAHP, TOPSIS, CRITIC, Entropy, DEMATEL |  
| Regression & Modeling | OLS, Ridge, Lasso, Logit, Mediation, Moderation |  
| Clustering | K-Means, DBSCAN, Two-Step Clustering, MDS |  
| Time Series & Forecasting | Grey Prediction, Exponential Smoothing |  
  
### Automated Parameter Decision Engine  
  
Instead of manual configuration, DIAS:  
  
- Detects variable types automatically  
- Runs assumption diagnostics (normality, homogeneity)  
- Selects the appropriate test or falls back when assumptions are violated (e.g., Fisher's Exact Test when Chi-square assumptions fail)  
- Adjusts hyperparameters dynamically (e.g., epsilon estimation for DBSCAN)  
  
### End-to-End Data Pipeline

Data Input (CSV/Excel) → Data Cleaning → Statistical Engine → Report Generation → Visualization

- Accepts `.csv`, `.xlsx`, and `.xls` files  
- Handles missing values and column-name normalization  
- Generates structured **Word reports** (`.docx`) and **Matplotlib charts**  
- No scripting required  
  
---  
  
## Supported Methods  
  
<details>  
<summary>Click to expand the full list of 77 methods</summary>  
  
| # | Method | Category |  
|---|--------|----------|  
| 1 | Analysis of Covariance (ANCOVA) | Difference |  
| 2 | Analytic Hierarchy Process (AHP) | MCDM |  
| 3 | Anderson-Darling Test | Normality |  
| 4 | Bartlett Test | Homogeneity |  
| 5 | Binary Logit Regression | Regression |  
| 6 | Canonical Correlation Analysis | Correlation |  
| 7 | Chi-Square Goodness-of-Fit Test | Categorical |  
| 8 | Chi-Squared Test | Categorical |  
| 9 | Cochran's Q Test | Difference |  
| 10 | Collinearity Analysis (VIF) | Regression |  
| 11 | Composite Index | MCDM |  
| 12 | Content Validity | Questionnaire |  
| 13 | Coupling Coordination Degree Model | MCDM |  
| 14 | CRITIC Weighting Method | MCDM |  
| 15 | D'Agostino's K² Test | Normality |  
| 16 | Delphi Method | Questionnaire |  
| 17 | DEMATEL Analysis | MCDM |  
| 18 | Density-Based Clustering (DBSCAN) | Clustering |  
| 19 | Descriptive Statistics | Description |  
| 20 | Efficacy Coefficient | MCDM |  
| 21 | Entropy Method | MCDM |  
| 22 | Exponential Smoothing | Time Series |  
| 23 | Factor Analysis | Questionnaire |  
| 24 | Friedman Test | Difference |  
| 25 | Fuzzy AHP (FAHP) | MCDM |  
| 26 | Gray Prediction Model | Time Series |  
| 27 | Grey Relational Analysis | MCDM |  
| 28 | Independence Weighting Method | MCDM |  
| 29 | Independent Samples T-Test | Difference |  
| 30 | Information Entropy Weight Method | MCDM |  
| 31 | Jarque-Bera Test | Normality |  
| 32 | KANO Model | Questionnaire |  
| 33 | Kappa Consistency Test | Questionnaire |  
| 34 | Kendall's Coordination Coefficient | Correlation |  
| 35 | K-Means Clustering | Clustering |  
| 36 | KS Test | Normality |  
| 37 | Lasso Regression | Regression |  
| 38 | Levene Test | Homogeneity |  
| 39 | Lilliefors Test | Normality |  
| 40 | Mediation Analysis | Regression |  
| 41 | Moderated Mediation | Regression |  
| 42 | Moderation Analysis | Regression |  
| 43 | Multi-Sample ANOVA | Difference |  
| 44 | Multidimensional Scaling (MDS) | Clustering |  
| 45 | Multinomial Logit Regression | Regression |  
| 46 | Multiple-Choice Question Analysis | Questionnaire |  
| 47 | MANOVA | Difference |  
| 48 | NPS (Net Promoter Score) | Questionnaire |  
| 49 | Nonlinear Regression | Regression |  
| 50 | Obstacle Degree Model | MCDM |  
| 51 | One-Sample ANOVA | Difference |  
| 52 | One-Sample T-Test | Difference |  
| 53 | One-Sample Wilcoxon Test | Difference |  
| 54 | Ordered Logit Regression | Regression |  
| 55 | OLS Linear Regression | Regression |  
| 56 | Paired T-Test | Difference |  
| 57 | Paired-Sample Wilcoxon Test | Difference |  
| 58 | Partial Correlation Analysis | Correlation |  
| 59 | Partial Least Squares (PLS) | Regression |  
| 60 | Pearson Correlation | Correlation |  
| 61 | Polynomial Regression | Regression |  
| 62 | Post-hoc Multiple Comparisons | Difference |  
| 63 | Price Sensitivity Meter (PSM) | Questionnaire |  
| 64 | Range Analysis | Difference |  
| 65 | Reliability Analysis | Questionnaire |  
| 66 | Ridge Regression | Regression |  
| 67 | Robust Linear Regression | Regression |  
| 68 | Runs Test | Normality |  
| 69 | Second-Order (Two-Step) Clustering | Clustering |  
| 70 | Shapiro-Wilk Test | Normality |  
| 71 | Spearman Correlation | Correlation |  
| 72 | Test-Retest Reliability | Questionnaire |  
| 73 | TOPSIS | MCDM |  
| 74 | TURF Combination Model | Questionnaire |  
| 75 | Undesirable SBM Model | MCDM |  
| 76 | Validity Analysis | Questionnaire |  
| 77 | Within-Group Inter-Rater Reliability (rwg) | Questionnaire |  
  
</details>  
  
---  
  
## Architecture

DIAS/
├── main.py # Application entry point & data-cleaning UI
├── requirements.txt # Python dependencies
└── Source/
├── Toolkit.py # Hub that registers all 77 analysis modules
├── Descriptive_Statistics.py
├── Pearson_Correlation_Analysis.py
├── Independent_Samples_T_Test_Analysis.py
└── ... (77 modules total)

```mermaid  
graph LR  
    A["main.py"] -->|"launches"| B["Toolkit.py"]  
    B -->|"opens"| C["Analysis Module"]  
    C -->|"reads"| D["CSV / Excel"]  
    C -->|"writes"| E[".docx Report"]  
    C -->|"plots"| F["Matplotlib Chart"]

main.py — launches the main window, handles data cleaning, and provides bilingual language switching.
Source/Toolkit.py — the hub that registers every analysis module in and renders them as searchable buttons in the GUI.MODULE_MAP
Source/*.py — each file is a self-contained analysis module with its own ttkbootstrap UI, statistical logic, and report generation.

Getting Started
Option 1 — Run from source (Python)
Prerequisites: Python 3.12

# Clone the repository  
git clone https://github.com/cyr950331-create/Design-Informatics-Analysis-System.git  
cd Design-Informatics-Analysis-System  
  
# Install dependencies  
pip install -r requirements.txt  
  
# Launch  
python main.py

Option 2 — Desktop executable (Windows)
No Python installation required.

Download from the OneDrive link.DIAS.exe
Double-click to launch.DIAS.exe
Follow the built-in user manual.
Supported OS: Windows 10 / 11

Usage
Open a dataset — load a or file from the main window..csv.xlsx
Clean data — use the built-in data-cleaning dialog to handle missing values and normalize column names.
Select a method — open the Toolkit and pick from 77 analysis methods (searchable, bilingual).
Run the analysis — DIAS automatically configures parameters, checks assumptions, and runs the test.
Export results — a Word report and charts are saved to your working directory.

Resources
GitHub Repository	https://github.com/cyr950331-create/Design-Informatics-Analysis-System
Desktop App Download	[OneDrive](https://1drv.ms/u/c/56791b21f8f8c84a/IQAmT6K5BrNUQoB1l56YLDz-ASQ3cBWppZEmBcyexUCMz_U?e=18fNCL)
Demo Video	[YouTube](https://www.youtube.com/watch?v=cFtxfOEURPE)

Tech Stack
Layer	               Libraries
GUI	                 ttkbootstrap 1.12
Data	               Pandas 2.2, NumPy 2.2, openpyxl, xlrd
Statistics	         SciPy 1.16, Statsmodels 0.14, pingouin 0.5, factor-analyzer
Machine Learning	   Scikit-learn 1.7
Visualization	       Matplotlib 3.10, Seaborn 0.13
Reports	             python-docx 1.2
Optimization	       PuLP 3.1, pyDEA 1.6
Packaging	           PyInstaller 6.16, Nuitka 2.7

Contributing
Contributions are welcome! To add a new analysis module:

Create following the pattern of existing modules (ttkbootstrap UI + analysis logic + report generation).Source/Your_Method_Analysis.py
Register it in inside .MODULE_MAPSource/Toolkit.py
Open a pull request.

License
This project is licensed under the MIT License.

Contact
Yingrui Chi
University of Camerino
Email: yingrui.chi@unicam.it
Email: cyr950331@unicam.it


