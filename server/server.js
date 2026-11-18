require("dotenv").config();

const express = require("express");
const { Pool } = require("pg");

const app = express();

app.use(express.json({ limit: "1kb" }));

app.use((req, res, next) => {
  res.setHeader("Content-Type", "application/json; charset=utf-8");
  next();
});

// ตั้งชื่อหน้าต่างให้ตรงกับที่ bot ใช้ตรวจสอบ
const appTitle = "AutoIt Server";

if (process.platform === "win32") {
  process.title = appTitle;
} else {
  process.stdout.write(`\x1b]2;${appTitle}\x1b\x5c`);
}

// ใช้ pool ร่วมกันทุก endpoint
// บัญชีต้องกำหนดสิทธิ์ read-only ที่ PostgreSQL
const pool = new Pool({
  host: process.env.PG_HOST,
  port: process.env.PG_PORT,
  user: process.env.PG_USER,
  password: process.env.PG_PASSWORD,
  database: process.env.PG_DATABASE,
});

// ==================== TELEMED ====================
// รับ body: {"hn":"001234567"}
app.post("/query_telemed", async (req, res) => {
  const body = req.body;

  // รับเฉพาะ hn ซึ่งเป็น string ตัวเลข 9 หลัก
  if (
    !body ||
    Array.isArray(body) ||
    typeof body !== "object" ||
    Object.keys(body).length !== 1 ||
    typeof body.hn !== "string" ||
    body.hn.length !== 9 ||
    !/^[0-9]{9}$/.test(body.hn)
  ) {
    return res.status(400).json({
      error: "INVALID_HN_BODY",
    });
  }

  // รายการสิทธิ์สำหรับ endpoint telemed เท่านั้น
  const CHECK_PTTYPE = new Set([
    "17","19","21","28","30","32",
    "37","44","50","59","99",
  ]);

  // SQL สำหรับ endpoint telemed เท่านั้น
  const SIT_SQL = `
    SELECT
        ovst.vn,
        COALESCE(ovst.oqueue, 0) AS oqueue,
        ovst.pttype,
        CASE
            WHEN COALESCE(ofs.nhso_fee_schedule_type_id, 999) = 101 THEN 'd'
            WHEN COALESCE(ofs.nhso_fee_schedule_type_id, 999) = 102 THEN 'o'
            ELSE 'n'
        END AS lst,
        COALESCE(pbw.bw, 0)::int AS bw
    FROM ovst
    LEFT JOIN ovst_fee_schedule ofs ON ofs.vn = ovst.vn
    LEFT JOIN (
        SELECT
            opdscreen.vn,
            opdscreen.hn,
            opdscreen.bw,
            CASE
                WHEN opdscreen.vn = MAX(opdscreen.vn)
                    OVER (PARTITION BY opdscreen.hn)
                THEN '1'
                ELSE '0'
            END AS flag_last_vn
        FROM opdscreen
        WHERE opdscreen.hn = $1
          AND opdscreen.bw IS NOT NULL
        ORDER BY vn
    ) AS pbw ON pbw.hn = ovst.hn AND pbw.flag_last_vn = '1'
    WHERE ovst.hn = $1
      AND ovst.vstdate = CURRENT_DATE
    ORDER BY ovst.vn DESC
    LIMIT 1;
  `;

  try {
    // ส่ง HN แยกจาก SQL ด้วย parameter binding
    const result = await pool.query(SIT_SQL, [body.hn]);
    const row = result.rows[0];

    // ไม่พบ visit วันนี้
    if (!row) {
      console.log("not foound visit: ", body.hn)
      return res.json(["888", "0", "0"]);
    }

    let pttype = row.pttype == null || String(row.pttype) === "" ? "888" : String(row.pttype);

    // ต่อท้าย d/n/o ตามรายการสิทธิ์ของ telemed
    if (CHECK_PTTYPE.has(pttype) && ["d", "n", "o"].includes(row.lst)) {
        pttype += row.lst;
    }
   
    // ส่งกลับ [pttype, bw, oqueue] เป็น string ทั้งสามค่า
    const responseData = [
      pttype,
      String(row.bw ?? 0),
      String(row.oqueue ?? 0),
    ];

    console.log("query_telemed result:", responseData);

    return res.json(responseData);
  } catch (err) {
    console.error("Telemed database query failed", {
      code: err.code || "UNKNOWN",
    });

    return res.status(500).json({
      error: "DATABASE_QUERY_FAILED",
    });
  }
});

// ==================== STUDENT ====================
app.post("/query_student", async (req, res) => {
  const body = req.body;

  // รับเฉพาะ {"hn":"001234567"}
  if (
    !body ||
    Array.isArray(body) ||
    typeof body !== "object" ||
    Object.keys(body).length !== 1 ||
    typeof body.hn !== "string" ||
    body.hn.length !== 9 ||
    !/^[0-9]{9}$/.test(body.hn)
  ) {
    return res.status(400).json({
      error: "INVALID_HN_BODY",
    });
  }

  // ตามเงื่อนไข GetSitFromDb2 ของ student
  const CHECK_PTTYPE = new Set(["25", "47"]);

  const SIT_SQL = `
    SELECT
        ovst.vn,
        COALESCE(ovst.oqueue, 0) AS oqueue,
        ovst.pttype,
        CASE
            WHEN COALESCE(ofs.nhso_fee_schedule_type_id, 999) = 101 THEN 'd'
            WHEN COALESCE(ofs.nhso_fee_schedule_type_id, 999) = 102 THEN 'o'
            ELSE 'n'
        END AS lst,
        COALESCE(pbw.bw, 0)::int AS bw
    FROM ovst
    LEFT JOIN ovst_fee_schedule ofs ON ofs.vn = ovst.vn
    LEFT JOIN (
        SELECT
            opdscreen.vn,
            opdscreen.hn,
            opdscreen.bw,
            CASE
                WHEN opdscreen.vn = MAX(opdscreen.vn)
                    OVER (PARTITION BY opdscreen.hn)
                THEN '1'
                ELSE '0'
            END AS flag_last_vn
        FROM opdscreen
        WHERE opdscreen.hn = $1
          AND opdscreen.bw IS NOT NULL
        ORDER BY vn
    ) AS pbw ON pbw.hn = ovst.hn AND pbw.flag_last_vn = '1'
    WHERE ovst.hn = $1
      AND ovst.vstdate = CURRENT_DATE
    ORDER BY ovst.vn DESC
    LIMIT 1;
  `;

  try {
    // HN เป็น parameter แยกจาก SQL
    const result = await pool.query(SIT_SQL, [body.hn]);
    const row = result.rows[0];

    // ไม่พบ visit วันนี้
    if (!row) {
      console.log("not foound visit: ", body.hn)
      return res.json(["888", "0", "0"]);
    }

    let pttype =
      row.pttype == null || String(row.pttype) === ""
        ? "888"
        : String(row.pttype);

    // ต่อท้าย d/n/o เฉพาะสิทธิ์ 25 และ 47
    if (CHECK_PTTYPE.has(pttype) && ["d", "n", "o"].includes(row.lst)) {
      pttype += row.lst;
    }
    
    // ส่งกลับ [pttype, bw, oqueue] เป็น string ทั้งสามค่า
    const responseData = [
      pttype,
      String(row.bw ?? 0),
      String(row.oqueue ?? 0),
    ];

    console.log("query_student result:", responseData);

    return res.json(responseData);
  } catch (err) {
    // ไม่บันทึก HN หรือข้อมูลผู้ป่วยลง log
    console.error("Student database query failed", {
      code: err.code || "UNKNOWN",
    });

    return res.status(500).json({
      error: "DATABASE_QUERY_FAILED",
    });
  }
});

// ==================== ELDER ======================
// เพิ่ม app.post("/query_elder", ...) ตรงนี้ภายหลัง

// ==================== ERROR HANDLER ==============
// วางหลัง endpoint ทั้งหมด
app.use((err, req, res, next) => {
  if (err.type === "entity.parse.failed") {
    return res.status(400).json({
      error: "INVALID_JSON",
    });
  }

  if (err.type === "entity.too.large") {
    return res.status(413).json({
      error: "BODY_TOO_LARGE",
    });
  }

  console.error("Request failed");

  return res.status(500).json({
    error: "INTERNAL_ERROR",
  });
});

// ==================== START SERVER ===============
const localhost_port = 3074;

app.listen(localhost_port, "127.0.0.1", () => {
  console.log(
    `API server running on http://127.0.0.1:${localhost_port}`
  );
});