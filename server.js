"use strict";

const express = require("express");
const mysql = require("mysql2/promise");
const multer = require("multer");
const XLSX = require("xlsx");
const helmet = require("helmet");
const path = require("path");
const os = require("os");
const fs = require("fs/promises");

const app = express();

app.disable("x-powered-by");
app.use(helmet({ contentSecurityPolicy: false }));
app.use(express.json({ limit: "1mb" }));

const PORT = Number(process.env.PORT) || 5000;

const DELETE_USER_ID = "MYMadmin";
const DELETE_PASSWORD = "MYMadmin2026";


const useSSL = process.env.DB_SSL === "true";



const pool = mysql.createPool({
  host: process.env.DB_HOST,
  port: Number(process.env.DB_PORT) || 4000,
  user: process.env.DB_USER,
  password: process.env.DB_PASSWORD,
  database: process.env.DB_NAME || "test",

  ssl: {
    minVersion: "TLSv1.2"
  },

  enableCleartextPlugin: true,

  waitForConnections: true,
  connectionLimit: 10,
  queueLimit: 0,
  charset: "utf8mb4",
  supportBigNumbers: true,
  bigNumberStrings: true
});




const upload = multer({
  dest: os.tmpdir(),
  limits: {
    fileSize: 50 * 1024 * 1024,
    files: 1
  }
});

const C = [
  "Sl#", "Mode", "Check", "Sub id", "Plan Name", "Client Name",
  "Tally ID", "Reporting Month", "Particulars", "Txn Date",
  "Value Date", "Reference_1", "Reference_2", "Reference_3",
  "Bank Dr", "Bank Cr", "Bank Bal", "Cash Dr", "Cash Cr",
  "Cash Bal", "Adj Dr", "Adj Cr", "Adj Bal", "Total Debit",
  "Total Credit", "Date_Val", "User"
];

const MONEY_COLUMNS = new Set([
  "Bank Dr", "Bank Cr", "Bank Bal",
  "Cash Dr", "Cash Cr", "Cash Bal",
  "Adj Dr", "Adj Cr", "Adj Bal",
  "Total Debit", "Total Credit"
]);

const DATE_COLUMNS = new Set([
  "Txn Date", "Value Date", "Date_Val"
]);

function quoteId(name) {
  if (!C.includes(name)) {
    throw new Error("Invalid column name");
  }

  return "`" + name.replace(/`/g, "") + "`";
}

function parseMoney(value, column, excelRowNumber) {
  if (value === null || value === undefined) return null;

  if (typeof value === "number") {
    if (!Number.isFinite(value)) {
      throw new Error(
        `Invalid number in "${column}", Excel row ${excelRowNumber}.`
      );
    }

    return value;
  }

  let text = String(value)
    .replace(/\u00A0/g, " ")
    .trim()
    .replace(/[\u2212\u2012\u2013\u2014]/g, "-");

  if (!text) return null;

  let negativeByParentheses = false;

  if (text.startsWith("(") && text.endsWith(")")) {
    negativeByParentheses = true;
    text = text.slice(1, -1).trim();
  }

  text = text
    .replace(/[₹$€£¥]/g, "")
    .replace(/\b(?:INR|USD|EUR|GBP|Rs\.?)\b/gi, "")
    .replace(/,/g, "")
    .replace(/\s/g, "")
    .trim();

  if (text === "-" || text === "") return null;

  const numericPattern =
    /^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?$/;

  if (!numericPattern.test(text)) {
    throw new Error(
      `Invalid number in "${column}", Excel row ${excelRowNumber}. ` +
      `Value: ${JSON.stringify(value)}`
    );
  }

  let number = Number(text);

  if (!Number.isFinite(number)) {
    throw new Error(
      `Invalid number in "${column}", Excel row ${excelRowNumber}.`
    );
  }

  if (negativeByParentheses) {
    if (number < 0) {
      throw new Error(
        `Invalid negative-number format in "${column}", ` +
        `Excel row ${excelRowNumber}.`
      );
    }

    number = -number;
  }

  return number;
}

function normalizeDate(value) {
  if (value instanceof Date) {
    if (Number.isNaN(value.getTime())) {
      throw new Error("Invalid Excel date.");
    }

    return value.toISOString().slice(0, 10);
  }

  return String(value).trim();
}

function formatMoneyValue(value) {
  if (value === null || value === undefined || value === "") {
    return null;
  }

  const number = Number(value);
  return Number.isFinite(number) ? Math.round(number) : value;
}

async function initDatabase() {
  const definitions = C.map(column => {
    let type = "VARCHAR(255)";

    if (MONEY_COLUMNS.has(column)) {
      type = "DECIMAL(20,4)";
    } else if (column === "Particulars") {
      type = "TEXT";
    } else if (
      [
        "Sl#", "Mode", "Check", "Sub id",
        "Tally ID", "Reporting Month"
      ].includes(column) ||
      DATE_COLUMNS.has(column)
    ) {
      type = "VARCHAR(100)";
    }

    return `${quoteId(column)} ${type} NULL`;
  });

  await pool.query(`
    CREATE TABLE IF NOT EXISTS transactions (
      id BIGINT UNSIGNED NOT NULL AUTO_INCREMENT PRIMARY KEY,
      ${definitions.join(",\n      ")},
      created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
      INDEX idx_client (\`Client Name\`),
      INDEX idx_plan (\`Plan Name\`),
      INDEX idx_month (\`Reporting Month\`),
      INDEX idx_tally (\`Tally ID\`),
      INDEX idx_txndate (\`Txn Date\`)
    ) ENGINE=InnoDB
      DEFAULT CHARACTER SET utf8mb4
      COLLATE utf8mb4_unicode_ci
  `);

  const [columns] = await pool.query(
    "SHOW COLUMNS FROM transactions LIKE 'User'"
  );

  if (columns.length === 0) {
    await pool.query(
      "ALTER TABLE transactions ADD COLUMN `User` VARCHAR(255) NULL"
    );
  }

  console.log("MySQL table is ready.");
}

function makeWhere(query) {
  const clauses = [];
  const params = [];
  const search = String(query.search || "").trim();

  if (search) {
    clauses.push(
      "(" + C.map(column => `${quoteId(column)} LIKE ?`).join(" OR ") + ")"
    );

    for (const column of C) {
      params.push(`%${search}%`);
    }
  }

  C.forEach((column, index) => {
    const value = String(query[`f${index}`] || "").trim();

    if (value) {
      clauses.push(`${quoteId(column)} LIKE ?`);
      params.push(`%${value}%`);
    }
  });

  return {
    sql: clauses.length ? ` WHERE ${clauses.join(" AND ")}` : "",
    params
  };
}

function getPage(value) {
  const number = Number(value);
  return Number.isSafeInteger(number) && number >= 1 ? number : 1;
}

function getLimit(value) {
  const number = Number(value);
  return [25, 50, 100, 200].includes(number) ? number : 50;
}

app.get("/", (req, res, next) => {
  res.sendFile(path.join(__dirname, "index.html"), error => {
    if (error) next(error);
  });
});

app.get("/api/columns", (req, res) => {
  res.json(C);
});

app.get("/api/data", async (req, res) => {
  try {
    const page = getPage(req.query.page);
    const limit = getLimit(req.query.limit);
    const offset = (page - 1) * limit;

    if (!Number.isSafeInteger(offset)) {
      return res.status(400).json({
        error: "Page number is too large."
      });
    }

    const { sql, params } = makeWhere(req.query);

    const [[count]] = await pool.execute(
      `SELECT COUNT(*) AS total FROM transactions${sql}`,
      params
    );

    const [[totals]] = await pool.execute(
      `SELECT
        COALESCE(SUM(\`Total Debit\`), 0) AS debit,
        COALESCE(SUM(\`Total Credit\`), 0) AS credit
       FROM transactions${sql}`,
      params
    );

    
const [rows] = await pool.query(
  `SELECT id, ${C.map(quoteId).join(", ")}
   FROM transactions${sql}
   ORDER BY id ASC
   LIMIT ${limit} OFFSET ${offset}`,
  params
);


    const formattedRows = rows.map(row => {
      for (const column of MONEY_COLUMNS) {
        row[column] = formatMoneyValue(row[column]);
      }

      return row;
    });

    const total = Number(count.total);

    res.json({
      rows: formattedRows,
      total,
      page,
      limit,
      pages: Math.max(1, Math.ceil(total / limit)),
      totals: {
        debit: formatMoneyValue(totals.debit),
        credit: formatMoneyValue(totals.credit)
      }
    });
  } catch (error) {
    console.error("Data API error:", error);
    res.status(500).json({ error: "Unable to load records." });
  }
});

// Deletes all rows in the transactions table after credential verification.
app.post("/api/delete-all", async (req, res) => {
  try {
    const userId =
      typeof req.body?.userId === "string"
        ? req.body.userId.trim()
        : "";

    const password =
      typeof req.body?.password === "string"
        ? req.body.password.trim()
        : "";

    if (!userId || !password) {
      return res.status(400).json({
        success: false,
        error: "User ID and password are required."
      });
    }

    if (userId !== DELETE_USER_ID || password !== DELETE_PASSWORD) {
      return res.status(401).json({
        success: false,
        error: "Invalid user ID or password."
      });
    }

    const [result] = await pool.query("DELETE FROM transactions");

    console.log(
      `Delete-all request succeeded. Rows deleted: ${result.affectedRows}`
    );

    return res.json({
      success: true,
      deletedRows: result.affectedRows,
      message:
        result.affectedRows === 0
          ? "No transaction records were found to delete."
          : `${result.affectedRows} transaction record(s) deleted.`
    });
  } catch (error) {
    console.error("Delete error:", error);

    return res.status(500).json({
      success: false,
      error: "Unable to delete records. Check the server log."
    });
  }
});

app.post("/api/import", upload.single("file"), async (req, res) => {
  let connection;

  try {
    if (!req.file) {
      return res.status(400).json({
        error: "Select an XLSX or CSV file."
      });
    }

    const extension = path.extname(req.file.originalname).toLowerCase();

    if (![".xlsx", ".csv"].includes(extension)) {
      return res.status(400).json({
        error: "Only .xlsx or .csv files are supported."
      });
    }

    const workbook = XLSX.readFile(req.file.path, {
      cellDates: true,
      raw: true
    });

    const firstSheetName = workbook.SheetNames[0];
    const sheet = firstSheetName
      ? workbook.Sheets[firstSheetName]
      : null;

    if (!sheet) {
      return res.status(400).json({
        error: "The uploaded workbook is empty."
      });
    }

    const data = XLSX.utils.sheet_to_json(sheet, {
      header: 1,
      defval: null,
      raw: true,
      blankrows: true
    });

    if (!data.length) {
      return res.status(400).json({
        error: "No rows found in the file."
      });
    }

    const headers = (data[0] || []).map(value =>
      String(value ?? "").trim()
    );

    const headersMatch =
      headers.length === C.length &&
      C.every((column, index) => column === headers[index]);

    if (!headersMatch) {
      return res.status(422).json({
        error:
          "Header mismatch. File must contain exactly these 27 " +
          "columns in the correct order: " + C.join(", ")
      });
    }

    const insertSql = `
      INSERT INTO transactions (${C.map(quoteId).join(", ")})
      VALUES (${C.map(() => "?").join(", ")})
    `;

    connection = await pool.getConnection();
    await connection.beginTransaction();

    let inserted = 0;

    for (let start = 1; start < data.length; start += 250) {
      const chunk = data.slice(start, start + 250);

      for (let rowIndex = 0; rowIndex < chunk.length; rowIndex++) {
        const row = chunk[rowIndex];

        if (
          !row ||
          !row.some(value =>
            value !== null &&
            value !== undefined &&
            value !== ""
          )
        ) {
          continue;
        }

        const excelRowNumber = start + rowIndex + 1;

        const values = C.map((column, index) => {
          const value = row[index];

          if (
            value === null ||
            value === undefined ||
            value === ""
          ) {
            return null;
          }

          if (MONEY_COLUMNS.has(column)) {
            return parseMoney(value, column, excelRowNumber);
          }

          if (DATE_COLUMNS.has(column)) {
            return normalizeDate(value);
          }

          if (value instanceof Date) {
            return value.toISOString().slice(0, 10);
          }

          return String(value).trim();
        });

        await connection.execute(insertSql, values);
        inserted++;
      }
    }

    await connection.commit();

    return res.json({
      success: true,
      inserted,
      message: `Import completed: ${inserted} records inserted.`
    });
  } catch (error) {
    if (connection) {
      try {
        await connection.rollback();
      } catch (rollbackError) {
        console.error("Rollback failed:", rollbackError);
      }
    }

    console.error("Import error:", error);

    return res.status(400).json({
      error: error.message || "Import failed. Check the file."
    });
  } finally {
    if (connection) connection.release();

    if (req.file) {
      await fs.unlink(req.file.path).catch(() => {});
    }
  }
});

app.get("/api/export", async (req, res) => {
  let connection;

  try {
    connection = await pool.getConnection();

    const { sql, params } = makeWhere(req.query);

    res.setHeader("Content-Type", "text/csv; charset=utf-8");
    res.setHeader(
      "Content-Disposition",
      'attachment; filename="transactions.csv"'
    );

    res.write("\uFEFF");

    function csv(value) {
      let text = value == null ? "" : String(value);

      if (/^[\s]*[=+\-@\t\r]/.test(text)) {
        text = "'" + text;
      }

      return `"${text.replace(/"/g, '""')}"`;
    }

    res.write(C.map(csv).join(",") + "\r\n");

    let offset = 0;
    const batchSize = 1000;

    while (true) {
      const [rows] = await connection.query(
  `SELECT ${C.map(quoteId).join(", ")}
   FROM transactions${sql}
   ORDER BY id ASC
   LIMIT ${batchSize} OFFSET ${offset}`,
  params
);

      if (rows.length === 0) break;

      for (const row of rows) {
        const line =
          C.map(column => csv(row[column])).join(",") + "\r\n";

        if (!res.write(line)) {
          await new Promise(resolve => res.once("drain", resolve));
        }
      }

      offset += rows.length;

      if (rows.length < batchSize) break;
    }

    res.end();
  } catch (error) {
    console.error("Export error:", error);

    if (!res.headersSent) {
      res.status(500).json({ error: "CSV export failed." });
    } else {
      res.destroy(error);
    }
  } finally {
    if (connection) connection.release();
  }
});

app.use((error, req, res, next) => {
  console.error("Server error:", error);

  if (error instanceof multer.MulterError) {
    return res.status(400).json({
      error:
        error.code === "LIMIT_FILE_SIZE"
          ? "File exceeds the 50 MB upload limit."
          : "File upload failed."
    });
  }

  return res.status(500).json({
    error: "Internal server error."
  });
});

async function startServer() {
  await initDatabase();
  await pool.query("SELECT 1");

  return app.listen(PORT, "0.0.0.0", () => {
    console.log(`Transaction Manager running at http://localhost:${PORT}`);
    console.log(
      `Connected database: ${process.env.DB_NAME || "transaction_manager"}`
    );
  });
}

let server;

startServer()
  .then(startedServer => {
    server = startedServer;
  })
  .catch(error => {
    console.error("Failed to start application:", error);
    process.exit(1);
  });

async function shutdown() {
  console.log("Shutting down server...");

  if (server) {
    await new Promise(resolve => server.close(resolve));
  }

  await pool.end();
  process.exit(0);
}

process.on("SIGINT", shutdown);
process.on("SIGTERM", shutdown);
