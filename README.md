# ⚡ DataPilot — AI Analytics Engine

An AI-powered analytics platform with natural language SQL generation, Python code generation, multi-file support, and live MySQL/PostgreSQL connection — built with React and Claude API.

---

## ✨ Features

| Mode | What it does |
|---|---|
| 💡 AI Insights | Plain English → Business insight + chart + recommendation |
| 🗃️ SQL Query | Plain English → MySQL & PostgreSQL queries with results |
| ⬡ Python Code | Plain English → Runnable Pandas code + Open in Colab |
| 🗄 Live Database | Connect MySQL or PostgreSQL directly |
| 📁 Multi-file | CSV, Excel, JSON, TSV — switch between datasets |
| ↓ Export | CSV export on every result |

---

## 🖥️ Screenshots

### Landing Page
<img width="1873" height="893" alt="Landing" src="https://github.com/user-attachments/assets/64edbc8f-5f71-4158-b49f-d28496bbda89" />


### SQL Query Mode
<img width="1842" height="917" alt="Sql" src="https://github.com/user-attachments/assets/7409069f-c6ec-40de-917c-6a2c8ab94c54" />

### Python Code Mode
<img width="1018" height="782" alt="Python" src="https://github.com/user-attachments/assets/152264e0-5d95-4e90-ae2a-1fd17982de3b" />

### AI Insights Mode
<img width="1883" height="920" alt="Insights" src="https://github.com/user-attachments/assets/a535bb71-bac6-4d0b-a573-969f5e11af62" />

---

## 🛠️ Tech Stack

- **React 18** — Frontend
- **Anthropic Claude Sonnet** — AI engine
- **Recharts** — Charts
- **SheetJS (xlsx)** — Excel parsing
- **Node.js + Express** — Backend proxy for DB connections

---

## 📦 Frontend Setup

```bash
git clone https://github.com/Ajaypatel06/datapilot.git
cd datapilot
npm install
npm start
```

Opens at `http://localhost:3000`

---

## 🗄️ Backend Setup (MySQL / PostgreSQL only)

```bash
cd backend
npm install
node server.js
```

Runs at `http://localhost:3001`

Then in DataPilot → Connect DB → enter `http://localhost:3001`

---

## 🗂️ Project Structure

datapilot/
├── src/
│ └── DataPilot.jsx
├── backend/
│ ├── server.js
│ └── package.json
├── Screenshot/
│ ├── Landing.png
│ ├── Sql.png
│ ├── Python.png
│ └── Insights.png
├── .env.example
├── backend/.env.example
├── .gitignore
├── package.json
└── README.md


---

## 🏗️ Architecture

React frontend (DataPilot.jsx)
│
├──/analyze──────► Node/Express backend (backend/server.js) ──► Claude API
│ holds ANTHROPIC_API_KEY server-side,
│ never bundled into the client build
│
└──/connect,/table,/query──► same backend ──► MySQL / PostgreSQL


Both the AI analysis calls and the live database calls go through the same
backend proxy. This keeps the Anthropic key and any DB credentials out of
the browser entirely — the frontend never talks to `api.anthropic.com` or
to a database directly.

---

## 💡 How to Use

1. Click **Launch App →** on the landing page
2. Upload a file or connect a database
3. Choose a mode — AI Insights / SQL Query / Python Code
4. Type your question in plain English
5. Get instant results

### Example questions:
- *"Which month had the highest revenue?"*
- *"Show total orders grouped by month"*
- *"Calculate month-over-month growth rate"*
- *"Find months where returns were above average"*

---

## 🔐 File Types Supported

| Format | Extension |
|---|---|
| Excel | `.xlsx` `.xls` |
| CSV | `.csv` |
| JSON | `.json` |
| TSV | `.tsv` |
| Text | `.txt` |

---

## 🧪 Eval Results

The SQL-generation prompt is tested against a fixed set of 15 questions run
on the app's real sample dataset, scored on structural correctness
(correct aggregate function, `GROUP BY`, `WHERE`, `ORDER BY`, `LIMIT`
presence for each question's intent).

**v1 (baseline): 11/15** (2/6 on the specific "total X by Y" pattern below)
Most failures traced to one root cause: on data that's already one row per
category (e.g. one row per month), "total X by Y" questions were answered
with `ORDER BY` instead of `SUM()` + `GROUP BY` — structurally correct-looking
SQL that would silently return wrong, unaggregated numbers against a real
multi-row-per-month table.

**v2 (patched): 14/15** (5/6 on the same pattern)
Fix: one line added to the system prompt requiring the matching aggregate
function whenever a question asks for a total/sum/average broken down by a
category, even if the sample rows look pre-aggregated. The one remaining
failure in both versions is a genuinely ambiguous question ("average X per
Y" on data that's already one row per Y) rather than a regression.

---

## 🧑‍💻 Author

**Ajay Patel** — Data Analyst
- GitHub: [github.com/Ajaypatel06](https://github.com/Ajaypatel06)
- LinkedIn: [linkedin.com/in/ajay-patel](https://linkedin.com/in/ajay-patel)

---

## 📄 License

MIT
