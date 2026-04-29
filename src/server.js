require("dotenv").config();

const express = require("express");
const path = require("path");
const { ExportService } = require("./exportService");

const app = express();

app.use(express.json({ limit: "1mb" }));
app.use(express.static(path.join(__dirname, "../public")));

const config = {
  port: parseInt(process.env.PORT || "8080", 10),
  googleServiceAccountJson: process.env.GOOGLE_SERVICE_ACCOUNT_JSON || "",
  googleCalendarId: process.env.GOOGLE_CALENDAR_ID || "",
  googleSpreadsheetId: process.env.GOOGLE_SPREADSHEET_ID || "",
  googleTimeZone: process.env.GOOGLE_TIMEZONE || "Asia/Ho_Chi_Minh",
  sheetGid: process.env.SHEET_GID || "0",
  appToken: process.env.APP_TOKEN || "",
};

if (!config.googleServiceAccountJson || !config.googleCalendarId || !config.googleSpreadsheetId) {
  // eslint-disable-next-line no-console
  console.error("Missing required env vars. Check GOOGLE_SERVICE_ACCOUNT_JSON, GOOGLE_CALENDAR_ID, GOOGLE_SPREADSHEET_ID.");
  process.exit(1);
}

const exportService = new ExportService(config);
let lastRun = null;

function requireToken(req, res, next) {
  if (!config.appToken) return next();
  const authHeader = req.headers.authorization || "";
  const bearer = authHeader.startsWith("Bearer ") ? authHeader.slice(7) : "";
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
  return `https://docs.google.com/spreadsheets/d/${config.googleSpreadsheetId}/edit?gid=${encodeURIComponent(config.sheetGid)}&rm=minimal`;
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

app.get("/api/config", requireToken, (_req, res) => {
  res.json({
    ok: true,
    calendarEmbedUrl: buildCalendarEmbedUrl(),
    sheetEmbedUrl: buildSheetEmbedUrl(),
    hasAuth: Boolean(config.appToken),
  });
});

app.get("/api/status", requireToken, (_req, res) => {
  res.json({ ok: true, lastRun });
});

app.post("/api/export/weekly-current", requireToken, async (_req, res) => {
  try {
    const result = await runAndTrack(() => exportService.exportWeeklyCurrentMonth());
    res.json({ ok: true, result });
  } catch (err) {
    res.status(500).json({ ok: false, error: err.message });
  }
});

app.post("/api/export/month-current", requireToken, async (_req, res) => {
  try {
    const result = await runAndTrack(() => exportService.exportFullCurrentMonth());
    res.json({ ok: true, result });
  } catch (err) {
    res.status(500).json({ ok: false, error: err.message });
  }
});

app.post("/api/export/month-custom", requireToken, async (req, res) => {
  try {
    const monthInput = req.body?.month;
    const result = await runAndTrack(() => exportService.exportFullByMonth(monthInput));
    res.json({ ok: true, result });
  } catch (err) {
    res.status(500).json({ ok: false, error: err.message });
  }
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
    const events = await exportService.fetchEvents(start, end);
    // Chuẩn hóa dữ liệu cho FullCalendar
    const result = events.map(ev => ({
      id: ev.id,
      title: ev.summary,
      startTime: ev.start?.dateTime || ev.start?.date,
      endTime: ev.end?.dateTime || ev.end?.date,
      description: ev.description,
      meetLink: ev.conferenceData?.entryPoints?.find(e => e.entryPointType === "video")?.uri || "",
    }));
    res.json({ ok: true, result });
  } catch (err) {
    res.status(500).json({ ok: false, error: err.message });
  }
});

app.get("*", (_req, res) => {
  res.sendFile(path.join(__dirname, "../public/index.html"));
});

app.listen(config.port, () => {
  // eslint-disable-next-line no-console
  console.log(`Server listening on port ${config.port}`);
});
