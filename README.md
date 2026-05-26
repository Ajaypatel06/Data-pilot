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
