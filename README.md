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
![alt text](Landing.png)

### SQL Query Mode
![alt text](Sql.png)

### Python Code Mode
![alt text](Python.png)

### AI Insights Mode
![alt text](Insights.png)

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

```
datapilot/
├── src/
│   └── DataPilot.jsx
├── backend/
│   ├── server.js
│   └── package.json
├── screenshots/
│   ├── landing.png
│   ├── sql.png
│   ├── python.png
│   └── insights.png
├── screenshots/
│   └── architecture.svg
├── .env.example
├── .gitignore
├── package.json
└── README.md
```

---

## 🏗️ Architecture

![Architecture](screenshots/architecture.svg)

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

## 🧑‍💻 Author

**Ajay Patel** — Data & Analytics Developer
- GitHub: [github.com/Ajaypatel06](https://github.com/Ajaypatel06)
- LinkedIn: [linkedin.com/in/ajay-patel](https://linkedin.com/in/ajay-patel)

---

## 📄 License

MIT