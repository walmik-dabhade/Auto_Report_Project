\# Automated CSR Community Education Programme Reporting



An automated reporting system that transforms programme data and qualitative reports into structured, funder-specific Community Education Programme progress reports.



\## Overview



This project automates the generation of CSR/community education programme reports for multiple funders.



The system combines:



\- Programme and learner data from Excel

\- KPI calculations

\- Activity-wise analysis

\- Attendance analysis

\- Language and mathematics learning-level analysis

\- Funder-specific insights

\- Marathi-to-English translation using Google Gemini

\- Automated charts and visualizations

\- Professionally formatted Word reports



\## Project Structure



```text

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

