\# Automated CSR Community Education Programme Reporting



An automated Python-based reporting system that converts programme data and qualitative reports into structured, funder-specific Community Education Programme progress reports.



\## 📌 Project Overview



This project automates the reporting workflow for a Community Education Programme supported by multiple CSR funding partners.



The system brings together quantitative programme data and qualitative field-level information to generate professional, funder-specific reports with KPIs, insights, charts, learning-level analysis, and translated narratives.



\## 🎯 Objectives



\- Reduce manual effort involved in periodic CSR reporting

\- Standardize report generation across multiple funders

\- Generate accurate funder-specific KPIs and insights

\- Combine quantitative and qualitative programme information

\- Automate charts and visualizations

\- Translate Marathi programme reports into English

\- Generate professionally formatted Word reports



\## ✨ Key Features



\### 📊 KPI \& Data Analysis



\- Programme-level KPI calculations

\- Funder-wise beneficiary analysis

\- Activity-wise reach analysis

\- Attendance analysis

\- Gender analysis

\- Continuity/dropout/migration analysis

\- Language learning-level analysis

\- Mathematics learning-level analysis

\- Location-based insights



\### 🏫 Activity Analysis



The system supports analysis of programme activities including:



\- Balwadi

\- Study Class

\- Library Class

\- Home Lending



\### 🤖 AI-Powered Translation



Qualitative Marathi programme reports are translated into English using the Google Gemini API.



The translated content is then incorporated into the reporting workflow to preserve field-level context and stories.



\### 📄 Automated Word Reports



The reporting engine generates structured Word reports containing:



\- Cover page

\- Executive summary

\- Programme highlights

\- Activity-wise objectives and observations

\- Attendance analysis

\- Learning-level analysis

\- Stakeholder engagement

\- Challenges and next-quarter plans

\- Special stories

\- Annexures

\- Charts and photographs



\### 👥 Funder-Specific Reporting



Reports can be generated separately for different funding partners using the same reporting framework while dynamically adapting the relevant KPIs and insights.



\## 🔄 End-to-End Workflow



```text

&#x20;               ┌─────────────────────┐

&#x20;               │  Excel Programme    │

&#x20;               │       Data          │

&#x20;               └──────────┬──────────┘

&#x20;                          │

&#x20;                          ▼

&#x20;               ┌─────────────────────┐

&#x20;               │     KPI Engine      │

&#x20;               └──────────┬──────────┘

&#x20;                          │

&#x20;                          ▼

&#x20;               ┌─────────────────────┐

&#x20;               │   Insight Engine    │

&#x20;               └──────────┬──────────┘

&#x20;                          │

&#x20;         ┌────────────────┼────────────────┐

&#x20;         ▼                ▼                ▼

&#x20;    Activity         Attendance       Learning Levels

&#x20;     Insights          Insights          Insights

&#x20;         │                │                │

&#x20;         └────────────────┼────────────────┘

&#x20;                          ▼

&#x20;               ┌─────────────────────┐

&#x20;               │  Funder-Specific    │

&#x20;               │      Analysis       │

&#x20;               └──────────┬──────────┘

&#x20;                          │

&#x20;                          ▼

&#x20;               ┌─────────────────────┐

&#x20;               │ Marathi Programme   │

&#x20;               │       Report        │

&#x20;               └──────────┬──────────┘

&#x20;                          │

&#x20;                          ▼

&#x20;               ┌─────────────────────┐

&#x20;               │   Gemini Translation│

&#x20;               │     Marathi → EN    │

&#x20;               └──────────┬──────────┘

&#x20;                          │

&#x20;                          ▼

&#x20;               ┌─────────────────────┐

&#x20;               │ Funder Report       │

&#x20;               │     Generator       │

&#x20;               └──────────┬──────────┘

&#x20;                          │

&#x20;                          ▼

&#x20;               ┌─────────────────────┐

&#x20;               │ Professional Word  │

&#x20;               │       Report        │

&#x20;               └─────────────────────┘





📁 Project Structure

Auto\_Report\_Project/

│

├── insights/

│   ├── activity.py

│   ├── attendance.py

│   ├── funder.py

│   ├── funder\_insights.py

│   ├── insight\_engine.py

│   ├── language.py

│   ├── location.py

│   ├── mathematics.py

│   └── overall.py

│

├── kpi/

│   └── kpi\_engine.py

│

├── report/

│   ├── funder\_report.py

│   ├── funder\_report\_generator.py

│   ├── funder\_report\_generator\_v4.py

│   ├── funder\_report\_generator\_v5.py

│   ├── report\_content.json

│   └── report\_generator.py

│

├── translation/

│   └── marathi\_to\_english.py

│

├── main.py

├── test\_gemini.py

├── requirements.txt

└── .gitignore

