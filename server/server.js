require('dotenv').config();
const express = require("express");
const { Pool } = require("pg");
const cors = require("cors");

const app = express();

app.use(cors());
app.use(express.json()); // ให้รับ JSON จาก AutoIt
app.use((req, res, next) => {
  res.setHeader("Content-Type", "application/json; charset=utf-8");
  next();
});

const appTitle = 'AutoIt Server';
// Set the command line title
if (process.platform === 'win32') {
  process.title = appTitle;
} else {
  process.stdout.write(`\x1b]2;${appTitle}\x1b\x5c`);
}

const host = process.env.PG_HOST
const port = process.env.PG_PORT 
const user = process.env.PG_USER 
const password = process.env.PG_PASSWORD 
const database = process.env.PG_DATABASE 
// ตั้งค่าการเชื่อมต่อ PostgreSQL
const pool = new Pool({
  host: host,
  port: port,
  user: user,
  password: password,
  database: database
});

// ตัวอย่าง API: query SQL
app.post("/query", async (req, res) => {
  try {
    const { sql } = req.body;
    const result = await pool.query(sql);
    console.log(result.rows); // เช็คตรงนี้ว่าเป็นไทยปกติไหม
    res.json(result.rows);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

const localhost_port = 3074

app.listen(localhost_port, () => {
  console.log(`API server running on http://localhost:${localhost_port}`);
});
