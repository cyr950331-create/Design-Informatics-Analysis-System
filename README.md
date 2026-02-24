DIAS – Design Informatics Analysis System

DIAS (Design Informatics Analysis System) is an open-source Python-based statistical analysis infrastructure tailored specifically for design research.

It enables design scholars to perform rigorous quantitative analysis without requiring formal training in statistics or programming, transforming statistical reasoning into an executable, rule-based decision system.

📌 Overview

Design research has long relied on intuition and empiricism. However, interdisciplinary integration and data-driven design demand structured quantitative methods.

DIAS addresses three persistent challenges:

Difficulty in selecting appropriate statistical methods

High barriers to parameter configuration

Limited integration between statistics and design research workflows

To solve these issues, DIAS:

Encapsulates 77 statistical methods

Provides automatic parameter configuration

Embeds a design-oriented case library

Generates standardized analytical reports and visualizations

Offers both a Python package and a desktop executable (EXE)

🚀 Key Features
1️⃣ Design-Oriented Statistical Classification

DIAS reorganizes statistical methods according to design research logic rather than traditional mathematical taxonomy.

Categories include:

Data Description & Validation

Questionnaire Analysis

Correlation Analysis

Difference Analysis

Multi-Criteria Decision-Making

Regression & Theory Modeling

Clustering

Each method is renamed in intuitive design language to improve accessibility.

2️⃣ Automated Parameter Decision Engine

A core innovation of DIAS is its rule-based heuristic parameter configuration mechanism.

Instead of requiring manual configuration, DIAS:

Detects variable types automatically

Performs assumption diagnostics

Selects appropriate statistical tests

Adjusts hyperparameters dynamically

Switches methods when assumptions are violated

Examples:

Automatic Fisher’s Exact Test fallback in Chi-square modules

Automated epsilon estimation for DBSCAN

Covariate extraction in ANCOVA

Objective weight calculation in CRITIC

This reduces configuration errors and improves reproducibility.

3️⃣ End-to-End Data Pipeline

DIAS implements a modular but linear workflow:

User Input → Data Cleaning → Statistical Engine → Parameter Decision → Report Generation → Visualization

The system:

Accepts CSV and Excel files

Performs missing value checks

Normalizes data types

Executes statistical modules

Generates structured text reports

Produces publication-ready visualizations

No scripting required.

🏗 Architecture

Main components:

DIAS/
│
├── icon/
├── Sample_data/
├── Source/
├── main/
└── user_manual/

The execution flow ensures:

Modular extensibility

Unified input-output interfaces

Standardized result objects

Reproducible workflows

🧠 Illustrative Use Case

A UX researcher analyzing:

Menu hierarchy

Icon style

Feedback method

User operation time

Using DIAS, the researcher:

Identified range analysis as appropriate

Cleaned data automatically

Performed statistical analysis

Generated visual and textual reports

Determined the most influential design factor

All without coding.

💻 Installation
Option 1 – Python Package

Requirements

Python 3.12

Install dependencies:

pandas

numpy

scipy

scikit-learn

statsmodels

matplotlib

Clone repository:

git clone https://github.com/cyr950331-create/Design-Informatics-Analysis-System

Run main module:

python main.py
Option 2 – Desktop Application (Recommended for Non-Programmers)

Supported OS: Windows 10 / 11

Launch DIAS.exe

Follow built-in user manual

📂 Resources

GitHub Repository
https://github.com/cyr950331-create/Design-Informatics-Analysis-System

Cloud Storage
https://1drv.ms/u/c/56791b21f8f8c84a/IQAmT6K5BrNUQoB1l56YLDz-ASQ3cBWppZEmBcyexUCMz_U?e=18fNCL

Demonstration Video
https://youtu.be/cFtxfOEURPE

📊 Technical Stack

DIAS is built upon:

Pandas

NumPy

SciPy

Scikit-learn

Statsmodels

Matplotlib

ttkbootstrap

🔬 Research Contribution

DIAS contributes to design informatics by:

Transforming statistical expertise into infrastructure

Formalizing decision pathways for method selection

Reducing modeling ambiguity

Supporting reproducible research

Bridging design practice and quantitative science

Rather than functioning merely as software, DIAS operates as a methodological infrastructure.

📧 Contact

Yingrui Chi
University of Camerino
Email: yingrui.chi@unicam.it

📜 License

MIT License