const { google } = require("googleapis");

function normalizeDisplayText(text) {
  if (!text) return "";
  return String(text)
    .replace(/[\u00A0\u2000-\u200B\u202F\u205F\u3000]/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function normalizeVietnameseNameKey(name) {
  let s = normalizeDisplayText(name).toLowerCase();
  if (!s) return "";
  s = s.normalize("NFD").replace(/[\u0300-\u036f]/g, "");
  s = s.replace(/đ/g, "d");
  s = s.replace(/[^a-z0-9 ]/g, " ").replace(/\s+/g, " ").trim();
  return s;
}

function getMonthOfWeek(startOfWeek) {
  const counts = {};
  for (let i = 0; i < 7; i += 1) {
    const d = new Date(startOfWeek);
    d.setDate(d.getDate() + i);
    const m = d.getMonth();
    counts[m] = (counts[m] || 0) + 1;
  }

  let maxCount = -1;
  let monthWithMaxDays = new Date(startOfWeek).getMonth();
  Object.keys(counts).forEach((key) => {
    if (counts[key] > maxCount) {
      maxCount = counts[key];
      monthWithMaxDays = parseInt(key, 10);
    }
  });

  return monthWithMaxDays;
}

function getMondayOfWeek(date) {
  const monday = new Date(date);
  monday.setDate(date.getDate() - ((date.getDay() + 6) % 7));
  monday.setHours(0, 0, 0, 0);
  return monday;
}

function formatEventDate(date, timeZone) {
  const weekday = new Intl.DateTimeFormat("en-US", {
    weekday: "short",
    timeZone,
  }).format(date);

  const dayMonth = new Intl.DateTimeFormat("en-GB", {
    day: "2-digit",
    month: "2-digit",
    timeZone,
  }).format(date);

  return `${weekday} (${dayMonth})`;
}

function parseMonthInput(mmYYYY) {
  const text = normalizeDisplayText(mmYYYY);
  const m = text.match(/^(0?[1-9]|1[0-2])\/(\d{4})$/);
  if (!m) {
    throw new Error("Invalid month format. Expected MM/YYYY.");
  }
  return {
    month: parseInt(m[1], 10) - 1,
    year: parseInt(m[2], 10),
  };
}

function coerceEventStart(event) {
  if (!event || !event.start) return null;
  if (event.start.dateTime) return new Date(event.start.dateTime);
  if (event.start.date) return new Date(`${event.start.date}T00:00:00`);
  return null;
}

function deepEqualArray(a, b) {
  if (!Array.isArray(a) || !Array.isArray(b) || a.length !== b.length) return false;
  for (let i = 0; i < a.length; i += 1) {
    if (a[i] !== b[i]) return false;
  }
  return true;
}

class ExportService {
  constructor(config) {
    this.config = config;
    this.auth = null;
    this.calendarApi = null;
    this.sheetsApi = null;
  }

  async init() {
    if (this.calendarApi && this.sheetsApi) return;

    let credentials;
    try {
      credentials = JSON.parse(this.config.googleServiceAccountJson);
    } catch (err) {
      throw new Error("GOOGLE_SERVICE_ACCOUNT_JSON is invalid JSON");
    }

    this.auth = new google.auth.GoogleAuth({
      credentials,
      scopes: [
        "https://www.googleapis.com/auth/calendar",
        "https://www.googleapis.com/auth/spreadsheets",
      ],
    });

    this.calendarApi = google.calendar({ version: "v3", auth: this.auth });
    this.sheetsApi = google.sheets({ version: "v4", auth: this.auth });
  }

  async exportWeeklyCurrentMonth() {
    const today = new Date();
    let year = today.getFullYear();

    const monday = getMondayOfWeek(today);
    const month = getMonthOfWeek(monday);

    if (month === 0 && today.getMonth() === 11) {
      year += 1;
    } else if (month === 11 && today.getMonth() === 0) {
      year -= 1;
    }

    return this.exportCalendarCore(year, month, {
      mode: "WEEKLY_INCREMENTAL",
      today,
    });
  }

  async exportFullCurrentMonth() {
    const today = new Date();
    return this.exportCalendarCore(today.getFullYear(), today.getMonth(), {
      mode: "FULL_MONTH",
      today,
    });
  }

  async exportFullByMonth(monthInput) {
    const { year, month } = parseMonthInput(monthInput);
    return this.exportCalendarCore(year, month, { mode: "FULL_MONTH" });
  }

  async exportCalendarCore(year, month, options = {}) {
    await this.init();

    const mode = options.mode || "WEEKLY_INCREMENTAL";
    const today = options.today || new Date();

    const firstDayOfMonth = new Date(year, month, 1);
    const lastDayOfMonth = new Date(year, month + 1, 0);

    const firstMondayOverallRange = getMondayOfWeek(firstDayOfMonth);
    const lastSundayOverallRange = new Date(lastDayOfMonth);
    lastSundayOverallRange.setDate(lastSundayOverallRange.getDate() + ((7 - lastSundayOverallRange.getDay()) % 7));
    lastSundayOverallRange.setHours(23, 59, 59, 999);

    const weekOverallIndexToMonthMap = {};
    const tmpDate = new Date(firstMondayOverallRange);
    let overallWeekCount = 0;
    while (tmpDate <= lastSundayOverallRange) {
      overallWeekCount += 1;
      weekOverallIndexToMonthMap[overallWeekCount] = getMonthOfWeek(tmpDate);
      tmpDate.setDate(tmpDate.getDate() + 7);
    }

    const getOverallWeekNumber = (date) => {
      const diffInDays = Math.floor((date - firstMondayOverallRange) / (1000 * 60 * 60 * 24));
      return Math.floor(diffInDays / 7) + 1;
    };

    let currentOverallWeek = getOverallWeekNumber(today);
    currentOverallWeek = Math.max(1, Math.min(currentOverallWeek, overallWeekCount));

    const sheetName = `TongKet_${year}_${month + 1}`;
    const summarySheetId = await this.ensureSummarySheet(sheetName);

    const currentMonthOverallWeeks = [];
    for (let w = 1; w <= overallWeekCount; w += 1) {
      if (weekOverallIndexToMonthMap[w] === month) currentMonthOverallWeeks.push(w);
    }

    const sheetWeekNumToOverallWeekNumMap = {};
    const overallWeekNumToSheetWeekMap = {};
    for (let i = 0; i < currentMonthOverallWeeks.length && i < 5; i += 1) {
      const sheetWeekNum = i + 1;
      const overallWeekNum = currentMonthOverallWeeks[i];
      sheetWeekNumToOverallWeekNumMap[sheetWeekNum] = overallWeekNum;
      overallWeekNumToSheetWeekMap[overallWeekNum] = sheetWeekNum;
    }

    const events = await this.fetchEvents(firstMondayOverallRange, lastSundayOverallRange);

    const classData = {};
    events.forEach((event) => {
      const startTime = coerceEventStart(event);
      if (!startTime) return;

      const overallWeekNum = getOverallWeekNumber(startTime);
      if (!(overallWeekNum in weekOverallIndexToMonthMap)) return;
      if (weekOverallIndexToMonthMap[overallWeekNum] !== month) return;
      if (!overallWeekNumToSheetWeekMap[overallWeekNum]) return;

      const rawClassName = normalizeDisplayText(event.summary || "");
      const classKey = normalizeVietnameseNameKey(rawClassName);
      if (!classKey) return;

      if (!classData[classKey]) {
        classData[classKey] = { displayName: rawClassName, weeks: {} };
      }

      if (!classData[classKey].weeks[overallWeekNum]) {
        classData[classKey].weeks[overallWeekNum] = [];
      }

      const eventDateStr = formatEventDate(startTime, this.config.googleTimeZone);
      if (!classData[classKey].weeks[overallWeekNum].includes(eventDateStr)) {
        classData[classKey].weeks[overallWeekNum].push(eventDateStr);
      }
    });

    let overallWeeksToUpdate = [];
    let weeksForTotal = [];

    if (mode === "FULL_MONTH") {
      overallWeeksToUpdate = currentMonthOverallWeeks.slice();
      weeksForTotal = currentMonthOverallWeeks.slice();
    } else {
      const isFirstTime = await this.isSheetNew(summarySheetId);
      if (isFirstTime) {
        currentMonthOverallWeeks.forEach((w) => {
          if (w <= currentOverallWeek) overallWeeksToUpdate.push(w);
        });
      } else {
        if (overallWeekNumToSheetWeekMap[currentOverallWeek]) {
          overallWeeksToUpdate.push(currentOverallWeek);
        }
        const idx = currentMonthOverallWeeks.indexOf(currentOverallWeek);
        if (idx > 0) overallWeeksToUpdate.push(currentMonthOverallWeeks[idx - 1]);
      }

      weeksForTotal = currentMonthOverallWeeks.filter((w) => w <= currentOverallWeek);
    }

    overallWeeksToUpdate = [...new Set(overallWeeksToUpdate)].sort((a, b) => a - b);
    const weeksToUpdateSet = new Set(overallWeeksToUpdate);

    const existingData = await this.readSheetRows(sheetName);
    const allRowsData = existingData.allRowsData;
    const studentRowMap = existingData.studentRowMap;

    let nextRow = existingData.nextRow;

    const rowUpdates = {};
    const formulaCells = [];

    Object.keys(classData).forEach((classKey) => {
      const classInfo = classData[classKey];
      const weeks = classInfo.weeks;
      let rowIndex = null;
      let finalDisplayName = classInfo.displayName || "";

      if (studentRowMap[classKey]) {
        rowIndex = studentRowMap[classKey].rowIndex;
        if (studentRowMap[classKey].displayName) {
          finalDisplayName = studentRowMap[classKey].displayName;
        }
      } else {
        rowIndex = nextRow;
        nextRow += 1;
        studentRowMap[classKey] = { rowIndex, displayName: finalDisplayName };
        allRowsData[rowIndex] = [finalDisplayName, "", "", "", "", "", "", ""];
        formulaCells.push(`I${rowIndex}`);
      }

      if (!rowUpdates[rowIndex]) {
        rowUpdates[rowIndex] = (allRowsData[rowIndex] || ["", "", "", "", "", "", "", ""]).slice();
      }

      const rowBuf = rowUpdates[rowIndex];
      rowBuf[0] = finalDisplayName;

      for (let sheetW = 1; sheetW <= 5; sheetW += 1) {
        const overallWNum = sheetWeekNumToOverallWeekNumMap[sheetW];
        if (weeksToUpdateSet.has(overallWNum)) {
          const colIdx = sheetW;
          rowBuf[colIdx] = weeks[overallWNum] ? weeks[overallWNum].join(", ") : "";
        }
      }

      let calculatedTotalInMonth = 0;
      weeksForTotal.forEach((w) => {
        if (weeks[w]) calculatedTotalInMonth += weeks[w].length;
      });
      rowBuf[6] = calculatedTotalInMonth;
    });

    const writeRanges = [];
    let updatedRows = 0;
    Object.keys(rowUpdates).forEach((rowKey) => {
      const rowIndex = parseInt(rowKey, 10);
      const newRow = rowUpdates[rowIndex].slice(0, 7);
      const oldRow = (allRowsData[rowIndex] || ["", "", "", "", "", "", ""]).slice(0, 7);

      if (!deepEqualArray(newRow, oldRow)) {
        writeRanges.push({
          range: `${sheetName}!A${rowIndex}:G${rowIndex}`,
          values: [newRow],
        });
        updatedRows += 1;
      }
    });

    if (writeRanges.length > 0) {
      await this.sheetsApi.spreadsheets.values.batchUpdate({
        spreadsheetId: this.config.googleSpreadsheetId,
        requestBody: {
          valueInputOption: "USER_ENTERED",
          data: writeRanges,
        },
      });
    }

    if (formulaCells.length > 0) {
      await this.sheetsApi.spreadsheets.batchUpdate({
        spreadsheetId: this.config.googleSpreadsheetId,
        requestBody: {
          requests: formulaCells.map((a1) => ({
            repeatCell: {
              range: a1NotationToGridRange(a1, summarySheetId),
              cell: {
                userEnteredValue: {
                  formulaValue: `=H${a1.slice(1)}*G${a1.slice(1)}`,
                },
              },
              fields: "userEnteredValue",
            },
          })),
        },
      });
    }

    return {
      ok: true,
      mode,
      sheetName,
      month: month + 1,
      year,
      weeksUpdated: overallWeeksToUpdate,
      classesSeen: Object.keys(classData).length,
      rowsUpdated: updatedRows,
      newRows: formulaCells.length,
      eventsFetched: events.length,
    };
  }

  async fetchEvents(startDate, endDate) {
    await this.init();
    const items = [];
    let pageToken;
    do {
      const res = await this.calendarApi.events.list({
        calendarId: this.config.googleCalendarId,
        timeMin: startDate.toISOString(),
        timeMax: endDate.toISOString(),
        singleEvents: true,
        orderBy: "startTime",
        maxResults: 2500,
        pageToken,
      });
      const pageItems = res.data.items || [];
      pageItems.forEach((ev) => {
        if (ev.status !== "cancelled") items.push(ev);
      });
      pageToken = res.data.nextPageToken;
    } while (pageToken);
    return items;
  }

  async ensureSummarySheet(sheetName) {
    const spreadsheet = await this.sheetsApi.spreadsheets.get({
      spreadsheetId: this.config.googleSpreadsheetId,
      includeGridData: false,
    });

    const found = (spreadsheet.data.sheets || []).find(
      (s) => s.properties && s.properties.title === sheetName,
    );

    if (found) return found.properties.sheetId;

    const addRes = await this.sheetsApi.spreadsheets.batchUpdate({
      spreadsheetId: this.config.googleSpreadsheetId,
      requestBody: {
        requests: [{ addSheet: { properties: { title: sheetName } } }],
      },
    });

    const sheetId = addRes.data.replies[0].addSheet.properties.sheetId;

    await this.sheetsApi.spreadsheets.values.update({
      spreadsheetId: this.config.googleSpreadsheetId,
      range: `${sheetName}!A1:I1`,
      valueInputOption: "USER_ENTERED",
      requestBody: {
        values: [[
          "Học sinh",
          "Tuần 1",
          "Tuần 2",
          "Tuần 3",
          "Tuần 4",
          "Tuần 5",
          "Tổng buổi",
          "Giá/buổi",
          "Tổng tiền",
        ]],
      },
    });

    return sheetId;
  }

  async isSheetNew(sheetId) {
    const res = await this.sheetsApi.spreadsheets.get({
      spreadsheetId: this.config.googleSpreadsheetId,
      ranges: [],
      includeGridData: false,
    });
    const target = (res.data.sheets || []).find((s) => s.properties.sheetId === sheetId);
    const meta = await this.sheetsApi.spreadsheets.values.get({
      spreadsheetId: this.config.googleSpreadsheetId,
      range: `${target.properties.title}!A2:A2`,
    });

    const firstDataCell = (meta.data.values || [])[0]?.[0];
    return !normalizeDisplayText(firstDataCell);
  }

  async readSheetRows(sheetName) {
    const res = await this.sheetsApi.spreadsheets.values.get({
      spreadsheetId: this.config.googleSpreadsheetId,
      range: `${sheetName}!A2:H`,
    });

    const values = res.data.values || [];
    const allRowsData = {};
    const studentRowMap = {};

    values.forEach((row, idx) => {
      const rowIndex = idx + 2;
      const normalizedRow = new Array(8).fill("");
      for (let i = 0; i < 8; i += 1) normalizedRow[i] = row[i] || "";
      allRowsData[rowIndex] = normalizedRow;

      const displayName = normalizeDisplayText(normalizedRow[0]);
      const studentKey = normalizeVietnameseNameKey(displayName);
      if (studentKey) {
        studentRowMap[studentKey] = { rowIndex, displayName };
      }
    });

    return {
      allRowsData,
      studentRowMap,
      nextRow: values.length + 2,
    };
  }

  async createEvent(payload) {
    await this.init();

    const { title, description, startTime, endTime } = payload;

    if (!title || !startTime || !endTime) {
      throw new Error("Missing required fields: title, startTime, endTime");
    }

    const event = {
      summary: title,
      description: description || "",
      start: { dateTime: new Date(startTime).toISOString() },
      end: { dateTime: new Date(endTime).toISOString() },
    };

    const res = await this.calendarApi.events.insert({
      calendarId: this.config.googleCalendarId,
      requestBody: event,
    });

    return {
      ok: true,
      eventId: res.data.id,
      eventTitle: res.data.summary,
      eventStart: res.data.start.dateTime,
    };
  }

  async updateEvent(eventId, payload) {
    await this.init();

    const { title, description, startTime, endTime } = payload;

    if (!eventId || !title || !startTime || !endTime) {
      throw new Error("Missing required fields: eventId, title, startTime, endTime");
    }

    const event = {
      summary: title,
      description: description || "",
      start: { dateTime: new Date(startTime).toISOString() },
      end: { dateTime: new Date(endTime).toISOString() },
    };

    const res = await this.calendarApi.events.update({
      calendarId: this.config.googleCalendarId,
      eventId: eventId,
      requestBody: event,
    });

    return {
      ok: true,
      eventId: res.data.id,
      eventTitle: res.data.summary,
      eventStart: res.data.start.dateTime,
    };
  }

  async deleteEvent(eventId) {
    await this.init();

    if (!eventId) {
      throw new Error("Missing required field: eventId");
    }

    await this.calendarApi.events.delete({
      calendarId: this.config.googleCalendarId,
      eventId: eventId,
    });

    return {
      ok: true,
      eventId: eventId,
      message: "Event deleted successfully",
    };
  }
}

function a1NotationToGridRange(a1, sheetId) {
  const match = /^([A-Z]+)(\d+)$/.exec(a1);
  if (!match) throw new Error(`Invalid A1 notation: ${a1}`);

  const col = match[1];
  const row = parseInt(match[2], 10);
  const colIndex = col.split("").reduce((acc, ch) => acc * 26 + (ch.charCodeAt(0) - 64), 0) - 1;

  return {
    sheetId,
    startRowIndex: row - 1,
    endRowIndex: row,
    startColumnIndex: colIndex,
    endColumnIndex: colIndex + 1,
  };
}

module.exports = {
  ExportService,
  parseMonthInput,
};
