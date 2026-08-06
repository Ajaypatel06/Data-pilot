// DataPilot Backend Proxy
// Connects DataPilot frontend to MySQL or PostgreSQL, and proxies the
// Claude API call so the Anthropic key never reaches the browser bundle.
//
// SETUP:
//   npm install express cors pg mysql2 dotenv
//   cp .env.example .env   → set ANTHROPIC_API_KEY
//   node server.js
//
// Then open DataPilot → "Connect DB" → enter http://localhost:3001
// (the same proxy URL is used automatically for the AI analysis calls)

require("dotenv").config();
const express = require("express");
const cors = require("cors");

const app = express();
app.use(cors());
app.use(express.json());

// ─── Helpers ────────────────────────────────────────────────────────────────

async function getPostgresClient(config) {
  const { Client } = require("pg");
  const client = new Client({
    host: config.host,
    port: parseInt(config.port) || 5432,
    database: config.database,
    user: config.user,
    password: config.password,
    connectionTimeoutMillis: 5000,
    ssl: config.host !== "localhost" && config.host !== "127.0.0.1"
      ? { rejectUnauthorized: false }
      : false
  });
  await client.connect();
  return client;
}

async function getMysqlConnection(config) {
  const mysql = require("mysql2/promise");
  return await mysql.createConnection({
    host: config.host,
    port: parseInt(config.port) || 3306,
    database: config.database,
    user: config.user,
    password: config.password,
    connectTimeout: 5000
  });
}

// ─── Routes ─────────────────────────────────────────────────────────────────

// Health check
app.get("/", (req, res) => {
  res.json({ status: "DataPilot proxy running", version: "1.0" });
});

// ─── AI analysis proxy ───────────────────────────────────────────────────────
// Forwards the SQL/Python/Insights prompts to Claude. The key lives only in
// this process's environment (ANTHROPIC_API_KEY) and is never sent to,
// or bundled into, the frontend.
app.post("/analyze", async (req, res) => {
  const { system, question, model, max_tokens } = req.body;

  if (!system || !question) {
    return res.status(400).json({ error: "Missing required fields: system, question" });
  }

  const apiKey = process.env.ANTHROPIC_API_KEY;
  if (!apiKey) {
    return res.status(500).json({ error: "ANTHROPIC_API_KEY is not set on the server (see .env.example)" });
  }

  try {
    const r = await fetch("https://api.anthropic.com/v1/messages", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "x-api-key": apiKey,
        "anthropic-version": "2023-06-01"
      },
      body: JSON.stringify({
        model: model || "claude-sonnet-4-20250514",
        max_tokens: max_tokens || 2500,
        system,
        messages: [{ role: "user", content: question }]
      })
    });

    const data = await r.json();
    if (!r.ok) {
      return res.status(r.status).json({ error: data.error?.message || "Anthropic API error" });
    }
    res.json(data);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// Connect and list tables
app.post("/connect", async (req, res) => {
  const { dbType, host, port, database, user, password } = req.body;

  if (!dbType || !host || !database || !user) {
    return res.status(400).json({ error: "Missing required fields: dbType, host, database, user" });
  }

  try {
    let tables = [];

    if (dbType === "postgresql") {
      const client = await getPostgresClient({ host, port, database, user, password });
      const result = await client.query(
        `SELECT table_name FROM information_schema.tables
         WHERE table_schema = 'public' AND table_type = 'BASE TABLE'
         ORDER BY table_name`
      );
      tables = result.rows.map(r => r.table_name);
      await client.end();

    } else if (dbType === "mysql") {
      const conn = await getMysqlConnection({ host, port, database, user, password });
      const [rows] = await conn.execute(
        `SELECT table_name FROM information_schema.tables
         WHERE table_schema = ? AND table_type = 'BASE TABLE'
         ORDER BY table_name`,
        [database]
      );
      tables = rows.map(r => r.table_name || r.TABLE_NAME);
      await conn.end();

    } else {
      return res.status(400).json({ error: "dbType must be 'postgresql' or 'mysql'" });
    }

    res.json({ success: true, tables });

  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// Load full table (up to 500 rows)
app.post("/table", async (req, res) => {
  const { dbType, host, port, database, user, password, table } = req.body;

  if (!table) return res.status(400).json({ error: "table is required" });

  // Basic SQL injection protection — only allow valid table names
  if (!/^[a-zA-Z0-9_]+$/.test(table)) {
    return res.status(400).json({ error: "Invalid table name" });
  }

  try {
    let rows = [];

    if (dbType === "postgresql") {
      const client = await getPostgresClient({ host, port, database, user, password });
      const result = await client.query(`SELECT * FROM "${table}" LIMIT 500`);
      rows = result.rows;
      await client.end();

    } else if (dbType === "mysql") {
      const conn = await getMysqlConnection({ host, port, database, user, password });
      const [result] = await conn.execute(`SELECT * FROM \`${table}\` LIMIT 500`);
      rows = result;
      await conn.end();
    }

    res.json({ success: true, rows, count: rows.length });

  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// Run a SQL query
app.post("/query", async (req, res) => {
  const { dbType, host, port, database, user, password, sql } = req.body;

  if (!sql) return res.status(400).json({ error: "sql is required" });

  // Safety: only allow SELECT queries
  const trimmed = sql.trim().toUpperCase();
  if (!trimmed.startsWith("SELECT") && !trimmed.startsWith("WITH")) {
    return res.status(400).json({ error: "Only SELECT queries are allowed" });
  }

  // Add LIMIT if not present (safety cap)
  const safeSql = sql.trim().replace(/;$/, "");
  const limitedSql = /\bLIMIT\b/i.test(safeSql) ? safeSql : `${safeSql} LIMIT 1000`;

  try {
    let rows = [];

    if (dbType === "postgresql") {
      const client = await getPostgresClient({ host, port, database, user, password });
      const result = await client.query(limitedSql);
      rows = result.rows;
      await client.end();

    } else if (dbType === "mysql") {
      const conn = await getMysqlConnection({ host, port, database, user, password });
      const [result] = await conn.execute(limitedSql);
      rows = result;
      await conn.end();
    }

    res.json({ success: true, rows, count: rows.length });

  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// ─── Start ───────────────────────────────────────────────────────────────────

const PORT = process.env.PORT || 3001;
app.listen(PORT, () => {
  console.log(`
  ⚡ DataPilot Proxy running on http://localhost:${PORT}
  
  Endpoints:
    GET  /           → health check
    POST /connect    → connect & list tables
    POST /table      → load table data
    POST /query      → run SQL query
  
  Supports: MySQL · PostgreSQL
  `);
});
