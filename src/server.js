require("dotenv").config();

const express = require("express");
const path = require("path");
const crypto = require("crypto");
const fs = require("fs");
const { ExportService } = require("./exportService");

const app = express();

app.use(express.json({ limit: "1mb" }));
app.use(express.static(path.join(__dirname, "../public")));
app.use("/media", express.static(path.join(__dirname, "../media")));

const config = {
  port: parseInt(process.env.PORT || "8080", 10),
  googleServiceAccountJson: process.env.GOOGLE_SERVICE_ACCOUNT_JSON || "",
  googleCalendarId: process.env.GOOGLE_CALENDAR_ID || "",
  googleSpreadsheetId: process.env.GOOGLE_SPREADSHEET_ID || "",
  googleTimeZone: process.env.GOOGLE_TIMEZONE || "Asia/Ho_Chi_Minh",
  sheetGid: process.env.SHEET_GID || "0",
  appToken: process.env.APP_TOKEN || "",
  sessionSecret: process.env.SESSION_SECRET || "",
  sessionTtlHours: parseInt(process.env.SESSION_TTL_HOURS || "12", 10),
  appUsersJson: process.env.APP_USERS_JSON || "",
  studentRateJson: process.env.STUDENT_RATE_JSON || "",
  defaultSessionRate: parseInt(process.env.DEFAULT_SESSION_RATE || "0", 10),
  databaseUrl: process.env.DATABASE_URL || "",
};

if (!config.googleServiceAccountJson || !config.googleCalendarId) {
  // eslint-disable-next-line no-console
  console.error("Missing required env vars. Check GOOGLE_SERVICE_ACCOUNT_JSON, GOOGLE_CALENDAR_ID.");
  process.exit(1);
}

const exportService = new ExportService(config);
let lastRun = null;

const SESSION_COOKIE_NAME = "taiga_session";
const SESSION_SECRET = config.sessionSecret || config.appToken || "please-change-session-secret";

function normalizeText(text) {
  return String(text || "")
    .trim()
    .replace(/\s+/g, " ");
}

function normalizeStudentKey(name) {
  return normalizeText(name)
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/đ/g, "d")
    .replace(/[^a-z0-9 ]/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function parseMonthInput(mmYYYY) {
  const m = normalizeText(mmYYYY).match(/^(0?[1-9]|1[0-2])\/(\d{4})$/);
  if (!m) {
    throw new Error("Invalid month format. Expected MM/YYYY.");
  }
  return {
    month: parseInt(m[1], 10) - 1,
    year: parseInt(m[2], 10),
  };
}

function getMondayOfWeek(date) {
  const d = new Date(date);
  d.setDate(d.getDate() - ((d.getDay() + 6) % 7));
  d.setHours(0, 0, 0, 0);
  return d;
}

function formatSessionDay(date, timeZone) {
  return new Intl.DateTimeFormat("vi-VN", {
    day: "2-digit",
    month: "2-digit",
    timeZone,
  }).format(date);
}

function hashPassword(rawPassword) {
  return crypto.createHash("sha256").update(String(rawPassword || "")).digest("hex");
}

function parseAppUsers(rawJson) {
  if (!rawJson) {
    return [
      {
        username: "teacher",
        passwordHash: hashPassword("teacher123"),
        role: "teacher",
        displayName: "Teacher",
      },
      {
        username: "student",
        passwordHash: hashPassword("student123"),
        role: "student",
        displayName: "Student",
        studentKey: "student",
      },
    ];
  }

  let parsed;
  try {
    parsed = JSON.parse(rawJson);
  } catch (_err) {
    throw new Error("APP_USERS_JSON must be valid JSON");
  }

  if (!Array.isArray(parsed) || parsed.length === 0) {
    throw new Error("APP_USERS_JSON must be a non-empty array");
  }

  return parsed.map((item) => {
    const username = normalizeText(item.username);
    const role = item.role === "teacher" ? "teacher" : "student";
    const displayName = normalizeText(item.displayName || username);

    if (!username) {
      throw new Error("Each account in APP_USERS_JSON must have username");
    }

    if (!item.password && !item.passwordHash) {
      throw new Error(`Account ${username} must have password or passwordHash`);
    }

    return {
      username,
      role,
      displayName,
      studentKey: normalizeStudentKey(item.studentKey || username),
      passwordHash: item.passwordHash || hashPassword(item.password),
    };
  });
}

function parseStudentRates(rawJson) {
  if (!rawJson) return {};
  let parsed;
  try {
    parsed = JSON.parse(rawJson);
  } catch (_err) {
    throw new Error("STUDENT_RATE_JSON must be valid JSON");
  }

  if (!parsed || typeof parsed !== "object" || Array.isArray(parsed)) {
    throw new Error("STUDENT_RATE_JSON must be a JSON object");
  }

  const rates = {};
  Object.entries(parsed).forEach(([studentName, rawRate]) => {
    const key = normalizeStudentKey(studentName);
    const rate = Number(rawRate);
    if (key && Number.isFinite(rate) && rate > 0) {
      rates[key] = Math.round(rate);
    }
  });
  return rates;
}

const users = parseAppUsers(config.appUsersJson);
const userMap = new Map(users.map((u) => [u.username, u]));
const studentRateMap = parseStudentRates(config.studentRateJson);
const paymentsStorePath = path.join(__dirname, "../data/payments.json");
const PAYMENTS_STORE_DB_KEY = "payments_store";
const STUDENT_RATES_KEY = "__studentRates";
const STUDENT_RATES_TABLE = "student_rates";
const useDatabaseStore = Boolean(config.databaseUrl);

let pgPool = null;
let dbInitPromise = null;

function getPgPool() {
  if (!useDatabaseStore) return null;
  if (!pgPool) {
    const isLocal =
      config.databaseUrl.includes("localhost") || config.databaseUrl.includes("127.0.0.1");
    const { Pool } = require("pg");
    pgPool = new Pool({
      connectionString: config.databaseUrl,
      ssl: isLocal ? false : { rejectUnauthorized: false },
      connectionTimeoutMillis: 7000,
      idleTimeoutMillis: 30000,
      max: 10,
    });

    pgPool.on("error", (err) => {
      // eslint-disable-next-line no-console
      console.error("[pg-pool] Unexpected error on idle client:", err.message);
    });
  }
  return pgPool;
}

function monthKeyFromYearMonth(year, monthIndex) {
  return `${year}-${String(monthIndex + 1).padStart(2, "0")}`;
}

function ensurePaymentsStoreFile() {
  const dirPath = path.dirname(paymentsStorePath);
  if (!fs.existsSync(dirPath)) {
    fs.mkdirSync(dirPath, { recursive: true });
  }

  if (!fs.existsSync(paymentsStorePath)) {
    fs.writeFileSync(paymentsStorePath, JSON.stringify({}, null, 2));
  }
}

function readPaymentsStoreFromFile() {
  ensurePaymentsStoreFile();
  try {
    const raw = fs.readFileSync(paymentsStorePath, "utf8");
    const parsed = JSON.parse(raw || "{}");
    if (!parsed || typeof parsed !== "object" || Array.isArray(parsed)) {
      return {};
    }
    return parsed;
  } catch (_err) {
    return {};
  }
}

function writePaymentsStoreToFile(store) {
  ensurePaymentsStoreFile();
  fs.writeFileSync(paymentsStorePath, JSON.stringify(store, null, 2));
}

async function ensureDatabaseInitialized() {
  if (!useDatabaseStore) return false;
  if (!dbInitPromise) {
    dbInitPromise = (async () => {
      const pool = getPgPool();
      await pool.query(`
        CREATE TABLE IF NOT EXISTS app_kv_store (
          store_key TEXT PRIMARY KEY,
          store_value JSONB NOT NULL,
          updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
        )
      `);
      await pool.query(`
        CREATE TABLE IF NOT EXISTS ${STUDENT_RATES_TABLE} (
          student_key TEXT PRIMARY KEY,
          rate INTEGER NOT NULL,
          updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
        )
      `);
      return true;
    })().catch((err) => {
      dbInitPromise = null;
      throw err;
    });
  }
  return dbInitPromise;
}

async function readStudentRatesFromDatabase() {
  if (!useDatabaseStore) return {};

  try {
    await ensureDatabaseInitialized();
    const pool = getPgPool();
    const res = await pool.query(`SELECT student_key, rate FROM ${STUDENT_RATES_TABLE}`);
    const map = {};
    res.rows.forEach((row) => {
      const key = normalizeStudentKey(row.student_key);
      const rate = Number(row.rate);
      if (key && Number.isFinite(rate) && rate > 0) {
        map[key] = Math.round(rate);
      }
    });
    return map;
  } catch (err) {
    // eslint-disable-next-line no-console
    console.error("[rates] Postgres read failed:", err.message);
    return {};
  }
}

async function upsertStudentRateToDatabase(studentKey, rate) {
  await ensureDatabaseInitialized();
  const pool = getPgPool();
  await pool.query(
    `
      INSERT INTO ${STUDENT_RATES_TABLE} (student_key, rate, updated_at)
      VALUES ($1, $2, NOW())
      ON CONFLICT (student_key)
      DO UPDATE SET rate = EXCLUDED.rate, updated_at = NOW()
    `,
    [studentKey, rate],
  );
}

async function readPaymentsStore() {
  const fileStore = readPaymentsStoreFromFile();
  if (!useDatabaseStore) {
    return fileStore;
  }

  try {
    await ensureDatabaseInitialized();
    const pool = getPgPool();
    const res = await pool.query(
      "SELECT store_value FROM app_kv_store WHERE store_key = $1",
      [PAYMENTS_STORE_DB_KEY],
    );

    if (res.rowCount > 0 && res.rows[0].store_value && typeof res.rows[0].store_value === "object") {
      return res.rows[0].store_value;
    }

    if (Object.keys(fileStore).length > 0) {
      await writePaymentsStore(fileStore);
    }
    return fileStore;
  } catch (err) {
    // eslint-disable-next-line no-console
    console.error("[payments] Postgres read failed, fallback to file store:", err.message);
    return fileStore;
  }
}

async function writePaymentsStore(store) {
  writePaymentsStoreToFile(store);

  if (!useDatabaseStore) {
    return;
  }

  try {
    await ensureDatabaseInitialized();
    const pool = getPgPool();
    await pool.query(
      `
        INSERT INTO app_kv_store (store_key, store_value, updated_at)
        VALUES ($1, $2::jsonb, NOW())
        ON CONFLICT (store_key)
        DO UPDATE SET store_value = EXCLUDED.store_value, updated_at = NOW()
      `,
      [PAYMENTS_STORE_DB_KEY, JSON.stringify(store)],
    );
  } catch (err) {
    // eslint-disable-next-line no-console
    console.error("[payments] Postgres write failed, saved to file store:", err.message);
  }
}

function sanitizePaidWeeks(rawPaidWeeks, maxWeekCount) {
  const valid = new Set();
  (rawPaidWeeks || []).forEach((w) => {
    const n = Number(w);
    if (Number.isInteger(n) && n >= 0 && n < maxWeekCount) {
      valid.add(n);
    }
  });
  return Array.from(valid).sort((a, b) => a - b);
}

function getStudentRatesRef(store) {
  if (!store[STUDENT_RATES_KEY] || typeof store[STUDENT_RATES_KEY] !== "object" || Array.isArray(store[STUDENT_RATES_KEY])) {
    store[STUDENT_RATES_KEY] = {};
  }
  return store[STUDENT_RATES_KEY];
}

function buildEffectiveRateMap(store) {
  const map = { ...studentRateMap };
  const runtimeRates = getStudentRatesRef(store);
  Object.entries(runtimeRates).forEach(([studentKey, rawRate]) => {
    const normalizedKey = normalizeStudentKey(studentKey);
    const rate = Number(rawRate);
    if (normalizedKey && Number.isFinite(rate) && rate > 0) {
      map[normalizedKey] = Math.round(rate);
    }
  });
  return map;
}

function getCurrentMonthParts() {
  const now = new Date();
  return {
    year: now.getFullYear(),
    month: now.getMonth(),
  };
}

function parseCookies(req) {
  const cookieHeader = req.headers.cookie || "";
  const cookies = {};
  cookieHeader.split(";").forEach((entry) => {
    const index = entry.indexOf("=");
    if (index <= 0) return;
    const key = entry.slice(0, index).trim();
    const value = entry.slice(index + 1).trim();
    cookies[key] = decodeURIComponent(value);
  });
  return cookies;
}

function signPayload(base64Payload) {
  return crypto
    .createHmac("sha256", SESSION_SECRET)
    .update(base64Payload)
    .digest("base64url");
}

function createSessionToken(user) {
  const exp = Date.now() + config.sessionTtlHours * 60 * 60 * 1000;
  const payload = {
    username: user.username,
    role: user.role,
    displayName: user.displayName,
    studentKey: user.studentKey || "",
    exp,
  };
  const base64Payload = Buffer.from(JSON.stringify(payload)).toString("base64url");
  const signature = signPayload(base64Payload);
  return `${base64Payload}.${signature}`;
}

function safeEqualString(a, b) {
  const aBuf = Buffer.from(String(a || ""));
  const bBuf = Buffer.from(String(b || ""));
  if (aBuf.length !== bBuf.length) return false;
  return crypto.timingSafeEqual(aBuf, bBuf);
}

function verifySessionToken(token) {
  if (!token || !token.includes(".")) return null;
  const [base64Payload, signature] = token.split(".");
  const expectedSignature = signPayload(base64Payload);
  if (!safeEqualString(signature, expectedSignature)) return null;

  let payload;
  try {
    payload = JSON.parse(Buffer.from(base64Payload, "base64url").toString("utf8"));
  } catch (_err) {
    return null;
  }

  if (!payload || payload.exp < Date.now()) return null;
  return payload;
}

function requireLogin(req, res, next) {
  const cookies = parseCookies(req);
  const token = cookies[SESSION_COOKIE_NAME];
  const session = verifySessionToken(token);

  if (!session) {
    return res.status(401).json({ ok: false, error: "Unauthorized" });
  }

  req.user = session;
  return next();
}

function requireRole(...roles) {
  return (req, res, next) => {
    if (!req.user || !roles.includes(req.user.role)) {
      return res.status(403).json({ ok: false, error: "Forbidden" });
    }
    return next();
  };
}

// Thay thế requireToken để hỗ trợ cả header và query string
function requireToken(req, res, next) {
  if (!config.appToken) return next();
  const authHeader = req.headers.authorization || "";
  let bearer = "";
  if (authHeader.startsWith("Bearer ")) {
    bearer = authHeader.slice(7);
  } else if (req.query && req.query.token) {
    bearer = req.query.token;
  }
  if (bearer !== config.appToken) {
    return res.status(401).json({ ok: false, error: "Unauthorized" });
  }
  return next();
}

function buildCalendarEmbedUrl() {
  const ctz = encodeURIComponent(config.googleTimeZone);
  return `https://calendar.google.com/calendar/embed?src=${encodeURIComponent(config.googleCalendarId)}&ctz=${ctz}&mode=WEEK&showTitle=0&showPrint=0&showTabs=0`;
}

function buildSheetEmbedUrl() {
  if (!config.googleSpreadsheetId) return "";
  return `https://docs.google.com/spreadsheets/d/${config.googleSpreadsheetId}/edit?gid=${encodeURIComponent(config.sheetGid)}&rm=minimal`;
}

function requireSpreadsheetConfig(req, res, next) {
  if (!config.googleSpreadsheetId) {
    return res.status(400).json({
      ok: false,
      error: "GOOGLE_SPREADSHEET_ID is required only for legacy export-to-sheet endpoints.",
    });
  }
  return next();
}

async function runAndTrack(runFn) {
  const startedAt = new Date().toISOString();
  try {
    const result = await runFn();
    lastRun = {
      ok: true,
      startedAt,
      finishedAt: new Date().toISOString(),
      result,
    };
    return result;
  } catch (err) {
    lastRun = {
      ok: false,
      startedAt,
      finishedAt: new Date().toISOString(),
      error: err.message,
    };
    throw err;
  }
}

app.get("/api/health", (_req, res) => {
  res.json({ ok: true, service: "calendar-sheet-manager", now: new Date().toISOString() });
});

app.post("/api/auth/login", (req, res) => {
  const username = normalizeText(req.body?.username);
  const password = String(req.body?.password || "");
  const user = userMap.get(username);

  if (!user || !safeEqualString(user.passwordHash, hashPassword(password))) {
    return res.status(401).json({ ok: false, error: "Sai tài khoản hoặc mật khẩu" });
  }

  const token = createSessionToken(user);
  const isProd = process.env.NODE_ENV === "production";
  res.cookie(SESSION_COOKIE_NAME, token, {
    httpOnly: true,
    sameSite: "lax",
    secure: isProd,
    maxAge: config.sessionTtlHours * 60 * 60 * 1000,
    path: "/",
  });

  return res.json({
    ok: true,
    user: {
      username: user.username,
      role: user.role,
      displayName: user.displayName,
    },
  });
});

app.post("/api/auth/logout", (_req, res) => {
  res.clearCookie(SESSION_COOKIE_NAME, { path: "/" });
  return res.json({ ok: true });
});

app.get("/api/auth/me", requireLogin, (req, res) => {
  res.json({
    ok: true,
    user: {
      username: req.user.username,
      role: req.user.role,
      displayName: req.user.displayName,
    },
  });
});

app.get("/api/dashboard/month", requireLogin, requireRole("teacher", "student"), async (req, res) => {
  try {
    const monthInput = req.query.month;
    const now = new Date();
    const { year, month } = monthInput
      ? parseMonthInput(monthInput)
      : { year: now.getFullYear(), month: now.getMonth() };

    const firstDay = new Date(year, month, 1);
    const lastDay = new Date(year, month + 1, 0, 23, 59, 59, 999);

    const firstMonday = getMondayOfWeek(firstDay);
    const lastSunday = new Date(lastDay);
    lastSunday.setDate(lastSunday.getDate() + ((7 - lastSunday.getDay()) % 7));
    lastSunday.setHours(23, 59, 59, 999);

    const weekStarts = [];
    const cursor = new Date(firstMonday);
    while (cursor <= lastSunday) {
      weekStarts.push(new Date(cursor));
      cursor.setDate(cursor.getDate() + 7);
    }

    const weekLabels = weekStarts.map((weekStart) => {
      const weekEnd = new Date(weekStart);
      weekEnd.setDate(weekEnd.getDate() + 6);

      const start = new Date(Math.max(weekStart.getTime(), firstDay.getTime()));
      const end = new Date(Math.min(weekEnd.getTime(), lastDay.getTime()));

      const startDay = String(start.getDate()).padStart(2, "0");
      const endDay = String(end.getDate()).padStart(2, "0");
      return `${startDay}-${endDay}`;
    });

    const events = await exportService.fetchEvents(firstDay, lastDay);
    const studentMap = new Map();

    events.forEach((ev) => {
      const rawName = normalizeText(ev.summary || "");
      const key = normalizeStudentKey(rawName);
      if (!key) return;

      const startISO = ev.start?.dateTime || ev.start?.date;
      if (!startISO) return;

      const eventDate = new Date(startISO);
      const weekIndex = Math.floor((getMondayOfWeek(eventDate) - firstMonday) / (7 * 24 * 60 * 60 * 1000));

      if (!studentMap.has(key)) {
        studentMap.set(key, {
          studentKey: key,
          studentName: rawName,
          sessions: 0,
          weekly: new Array(weekStarts.length).fill(0),
          weeklyDates: Array.from({ length: weekStarts.length }, () => []),
        });
      }

      const item = studentMap.get(key);
      item.sessions += 1;
      if (weekIndex >= 0 && weekIndex < item.weekly.length) {
        item.weekly[weekIndex] += 1;
        const dayLabel = formatSessionDay(eventDate, config.googleTimeZone);
        if (!item.weeklyDates[weekIndex].includes(dayLabel)) {
          item.weeklyDates[weekIndex].push(dayLabel);
        }
      }
    });

    const paymentsStore = await readPaymentsStore();
    const fileStore = readPaymentsStoreFromFile();
    const fallbackRates = {
      ...getStudentRatesRef(fileStore),
      ...getStudentRatesRef(paymentsStore),
    };

    const effectiveRateMap = buildEffectiveRateMap(paymentsStore);
    Object.entries(fallbackRates).forEach(([k, r]) => {
      const normKey = normalizeStudentKey(k);
      const numRate = Number(r);
      if (normKey && Number.isFinite(numRate) && numRate > 0) {
        effectiveRateMap[normKey] = Math.round(numRate);
      }
    });

    const dbRateMap = await readStudentRatesFromDatabase();
    Object.assign(effectiveRateMap, dbRateMap);

    // Auto-sync fallback rates to Postgres student_rates table if DB is up
    if (useDatabaseStore && Object.keys(dbRateMap).length > 0) {
      Object.entries(fallbackRates).forEach(([k, r]) => {
        if (!dbRateMap[k] && Number.isFinite(Number(r)) && Number(r) > 0) {
          upsertStudentRateToDatabase(k, Math.round(Number(r))).catch(() => {});
        }
      });
    }

    const students = Array.from(studentMap.values())
      .map((item) => {
        const rate = effectiveRateMap[item.studentKey] || 0;
        return {
          ...item,
          weeklyDetails: item.weeklyDates.map((days) => days.join(", ")),
          rate,
          tuition: rate * item.sessions,
        };
      })
      .sort((a, b) => a.studentName.localeCompare(b.studentName, "vi"));

    const visibleStudents = req.user.role === "teacher"
      ? students
      : students.filter((student) => student.studentKey === (req.user.studentKey || normalizeStudentKey(req.user.username)));

    const monthKey = monthKeyFromYearMonth(year, month);
    const monthPayments = paymentsStore[monthKey] || {};

    const visibleStudentsWithPayment = visibleStudents.map((student) => {
      const paymentState = monthPayments[student.studentKey] || {};
      const monthlyPaid = Boolean(paymentState.monthlyPaid);
      const paidWeeks = sanitizePaidWeeks(paymentState.paidWeeks, weekStarts.length);
      const manualPaidAmount = Math.max(0, Math.round(Number(paymentState.manualPaidAmount || 0)));

      const paidSessions = monthlyPaid
        ? student.sessions
        : paidWeeks.reduce((sum, weekIndex) => sum + (student.weekly[weekIndex] || 0), 0);

      const paidAmount = Math.min(student.tuition, paidSessions * student.rate + manualPaidAmount);
      const outstanding = Math.max(0, student.tuition - paidAmount);

      return {
        ...student,
        payment: {
          monthlyPaid,
          paidWeeks,
          manualPaidAmount,
          updatedAt: paymentState.updatedAt || null,
        },
        paidAmount,
        outstanding,
      };
    });

    const totalSessions = visibleStudentsWithPayment.reduce((sum, s) => sum + s.sessions, 0);
    const totalRevenue = visibleStudentsWithPayment.reduce((sum, s) => sum + s.tuition, 0);
    const totalPaid = visibleStudentsWithPayment.reduce((sum, s) => sum + s.paidAmount, 0);
    const totalOutstanding = visibleStudentsWithPayment.reduce((sum, s) => sum + s.outstanding, 0);

    return res.json({
      ok: true,
      month: `${String(month + 1).padStart(2, "0")}/${year}`,
      role: req.user.role,
      weeks: weekLabels,
      summary: {
        students: visibleStudentsWithPayment.length,
        totalSessions,
        totalRevenue,
        totalPaid,
        totalOutstanding,
      },
      students: visibleStudentsWithPayment,
    });
  } catch (err) {
    return res.status(400).json({ ok: false, error: err.message });
  }
});

app.post("/api/payments/monthly", requireLogin, requireRole("teacher"), async (req, res) => {
  try {
    const monthInput = req.body?.month;
    const studentKey = normalizeStudentKey(req.body?.studentKey);
    const paid = Boolean(req.body?.paid);

    if (!studentKey) {
      throw new Error("Missing studentKey");
    }

    const { year, month } = parseMonthInput(monthInput);
    const monthKey = monthKeyFromYearMonth(year, month);
    const store = await readPaymentsStore();
    store[monthKey] = store[monthKey] || {};

    const current = store[monthKey][studentKey] || { monthlyPaid: false, paidWeeks: [] };
    store[monthKey][studentKey] = {
      ...current,
      monthlyPaid: paid,
      updatedAt: new Date().toISOString(),
    };

    await writePaymentsStore(store);
    return res.json({ ok: true });
  } catch (err) {
    return res.status(400).json({ ok: false, error: err.message });
  }
});

app.post("/api/payments/weekly", requireLogin, requireRole("teacher"), async (req, res) => {
  try {
    const monthInput = req.body?.month;
    const studentKey = normalizeStudentKey(req.body?.studentKey);
    const weekIndex = Number(req.body?.weekIndex);
    const paid = Boolean(req.body?.paid);

    if (!studentKey) {
      throw new Error("Missing studentKey");
    }

    if (!Number.isInteger(weekIndex) || weekIndex < 0 || weekIndex > 5) {
      throw new Error("Invalid weekIndex");
    }

    const { year, month } = parseMonthInput(monthInput);
    const monthKey = monthKeyFromYearMonth(year, month);
    const store = await readPaymentsStore();
    store[monthKey] = store[monthKey] || {};

    const current = store[monthKey][studentKey] || { monthlyPaid: false, paidWeeks: [] };
    const paidWeeks = sanitizePaidWeeks(current.paidWeeks, 6);
    const set = new Set(paidWeeks);

    if (paid) {
      set.add(weekIndex);
    } else {
      set.delete(weekIndex);
    }

    store[monthKey][studentKey] = {
      ...current,
      paidWeeks: Array.from(set).sort((a, b) => a - b),
      updatedAt: new Date().toISOString(),
    };

    await writePaymentsStore(store);
    return res.json({ ok: true });
  } catch (err) {
    return res.status(400).json({ ok: false, error: err.message });
  }
});

app.post("/api/rates/update", requireLogin, requireRole("teacher"), async (req, res) => {
  try {
    const studentKey = normalizeStudentKey(req.body?.studentKey);
    const rate = Math.round(Number(req.body?.rate || 0));

    if (!studentKey) {
      throw new Error("Missing studentKey");
    }
    if (!Number.isFinite(rate) || rate <= 0) {
      throw new Error("Invalid rate");
    }

    let persistedTo = useDatabaseStore ? "database" : "file";
    let warning = null;

    if (useDatabaseStore) {
      try {
        await upsertStudentRateToDatabase(studentKey, rate);
      } catch (err) {
        persistedTo = "fallback";
        warning = `Database unavailable, saved to fallback store: ${err.message}`;
        // eslint-disable-next-line no-console
        console.error("[rates] Database upsert failed, fallback to local store:", err.message);
      }
    }

    // Always keep payments store (both local file and DB app_kv_store) updated
    const store = await readPaymentsStore();
    const rates = getStudentRatesRef(store);
    rates[studentKey] = rate;
    await writePaymentsStore(store);

    return res.json({ ok: true, studentKey, rate, persistedTo, warning });
  } catch (err) {
    return res.status(400).json({ ok: false, error: err.message });
  }
});

app.get("/api/config", requireToken, (_req, res) => {
  res.json({
    ok: true,
    calendarEmbedUrl: buildCalendarEmbedUrl(),
    sheetEmbedUrl: "",
    hasSheet: false,
    hasAuth: Boolean(config.appToken),
  });
});

app.get("/api/status", requireToken, (_req, res) => {
  res.json({ ok: true, lastRun });
});

app.post("/api/export/weekly-current", requireToken, (_req, res) => {
  res.status(410).json({
    ok: false,
    error: "This endpoint is disabled. Use /app dashboard flow instead of Google Sheet export.",
  });
});

app.post("/api/export/month-current", requireToken, (_req, res) => {
  res.status(410).json({
    ok: false,
    error: "This endpoint is disabled. Use /app dashboard flow instead of Google Sheet export.",
  });
});

app.post("/api/export/month-custom", requireToken, (_req, res) => {
  res.status(410).json({
    ok: false,
    error: "This endpoint is disabled. Use /app dashboard flow instead of Google Sheet export.",
  });
});

app.post("/api/events/create", requireToken, async (req, res) => {
  try {
    const payload = req.body;
    const result = await exportService.createEvent(payload);
    res.json({ ok: true, result });
  } catch (err) {
    res.status(500).json({ ok: false, error: err.message });
  }
});

app.get("/api/events", requireToken, async (req, res) => {
  try {
    // Lấy toàn bộ sự kiện trong 1 năm (hoặc giới hạn theo nhu cầu)
    const now = new Date();
    const start = new Date(now.getFullYear(), 0, 1);
    const end = new Date(now.getFullYear(), 11, 31, 23, 59, 59);
    
    // eslint-disable-next-line no-console
    console.log(`[/api/events] Fetching events from ${start.toISOString()} to ${end.toISOString()}`);
    
    const events = await exportService.fetchEvents(start, end);
    
    // eslint-disable-next-line no-console
    console.log(`[/api/events] Fetched ${events ? events.length : 'null'} events`);
    
    if (!events) {
      throw new Error("fetchEvents returned null");
    }
    
    if (!Array.isArray(events)) {
      throw new Error(`fetchEvents returned ${typeof events}, expected array`);
    }
    
    // Định dạng đúng cho FullCalendar: start, end
    const result = events.map(ev => ({
      id: ev.id,
      title: ev.summary,
      start: ev.start?.dateTime || ev.start?.date,
      end: ev.end?.dateTime || ev.end?.date,
      description: ev.description,
      meetLink: ev.conferenceData?.entryPoints?.find(e => e.entryPointType === "video")?.uri || "",
    }));
    res.json({ ok: true, events: result });
  } catch (err) {
    // eslint-disable-next-line no-console
    console.error(`[/api/events] Error: ${err.message}`, err.stack);
    res.status(500).json({ ok: false, error: err.message });
  }
});

app.put("/api/events/:eventId", requireToken, async (req, res) => {
  try {
    const { eventId } = req.params;
    const payload = req.body;
    const result = await exportService.updateEvent(eventId, payload);
    res.json({ ok: true, result });
  } catch (err) {
    res.status(500).json({ ok: false, error: err.message });
  }
});

app.delete("/api/events/:eventId", requireToken, async (req, res) => {
  try {
    const { eventId } = req.params;
    const result = await exportService.deleteEvent(eventId);
    res.json({ ok: true, result });
  } catch (err) {
    res.status(500).json({ ok: false, error: err.message });
  }
});

app.get("/home", (_req, res) => {
  res.sendFile(path.join(__dirname, "../public/home.html"));
});

app.get("/app", (_req, res) => {
  res.sendFile(path.join(__dirname, "../public/app.html"));
});

app.get("/learn", (_req, res) => {
  res.sendFile(path.join(__dirname, "../public/learn.html"));
});

app.get("/contact", (_req, res) => {
  res.sendFile(path.join(__dirname, "../public/contact.html"));
});

app.get("/articles", (_req, res) => {
  res.sendFile(path.join(__dirname, "../public/articles.html"));
});

app.get("*", (_req, res) => {
  res.sendFile(path.join(__dirname, "../public/home.html"));
});

app.listen(config.port, () => {
  // eslint-disable-next-line no-console
  console.log(`Server listening on port ${config.port}`);
});
