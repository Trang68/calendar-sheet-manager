const fs = require("fs");
const path = require("path");

class DictionaryService {
  constructor(dataDir) {
    this.dataDir = dataDir || path.join(__dirname, "../data/dict");
    this.ruViList = null;
    this.ruViMap = null;
    this.viRuList = null;
    this.viRuMap = null;
    this.famData = null;
    this.relCache = new Map();
    this.formsCache = new Map();
    this.isLoaded = false;
    this.isInitializing = false;
  }

  fold(s) {
    return (s || "")
      .toLowerCase()
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .replace(/đ/g, "d")
      .trim();
  }

  stripStress(s) {
    return (s || "").replace(/\u0301/g, "").trim();
  }

  cleanLetter(word) {
    const w = this.stripStress(word).toLowerCase().replace(/ё/g, "е");
    return /^[а-я]/.test(w) ? w[0] : null;
  }

  async ensureFile(filename) {
    const filePath = path.join(this.dataDir, filename);
    if (fs.existsSync(filePath)) return true;
    try {
      if (!fs.existsSync(this.dataDir)) {
        fs.mkdirSync(this.dataDir, { recursive: true });
      }
      const url = `https://rusitrki.com/data/${encodeURIComponent(filename)}`;
      console.log(`[DictionaryService] Downloading ${filename}...`);
      const res = await fetch(url, { headers: { "User-Agent": "Mozilla/5.0" } });
      if (res.ok) {
        const buf = await res.arrayBuffer();
        fs.writeFileSync(filePath, Buffer.from(buf));
        console.log(`[DictionaryService] Successfully saved ${filename}`);
        return true;
      }
    } catch (e) {
      console.warn(`[DictionaryService] Auto-download failed for ${filename}:`, e.message);
    }
    return false;
  }

  async initAsync() {
    if (this.isLoaded || this.isInitializing) return;
    this.isInitializing = true;
    try {
      const ruViPath = path.join(this.dataDir, "dict-ru-vi.json");
      if (!fs.existsSync(ruViPath)) {
        await this.ensureFile("dict-ru-vi.json");
      }
      const viRuPath = path.join(this.dataDir, "dict-vi-ru.json");
      if (!fs.existsSync(viRuPath)) {
        await this.ensureFile("dict-vi-ru.json");
      }
      const famPath = path.join(this.dataDir, "dict-fam.json");
      if (!fs.existsSync(famPath)) {
        await this.ensureFile("dict-fam.json");
      }

      this.init();
    } catch (e) {
      console.error("[DictionaryService] initAsync error:", e);
    } finally {
      this.isInitializing = false;
    }
  }

  init() {
    if (this.isLoaded) return;
    try {
      const ruViPath = path.join(this.dataDir, "dict-ru-vi.json");
      if (fs.existsSync(ruViPath)) {
        const raw = fs.readFileSync(ruViPath, "utf8");
        const json = JSON.parse(raw);
        this.ruViMap = new Map();
        this.ruViList = (json.rows || []).map((r, i) => {
          const entry = {
            id: `ru-vi:${i}`,
            ru: r[0],
            pron: r[1] || "",
            pos: r[2] || "",
            cat: r[3] || "",
            vi: r[4] || "",
            ex: r[5] || "",
            gram: r[6] || "",
            dir: "ru-vi",
          };
          const cleanRu = this.stripStress(entry.ru).toLowerCase();
          entry._w = this.fold(entry.ru);
          entry._vi = this.fold(entry.vi);
          entry._hay = this.fold(`${entry.ru} ${entry.pron} ${entry.vi} ${entry.ex}`);
          if (!this.ruViMap.has(cleanRu)) {
            this.ruViMap.set(cleanRu, entry);
          }
          return entry;
        });
      }

      const viRuPath = path.join(this.dataDir, "dict-vi-ru.json");
      if (fs.existsSync(viRuPath)) {
        const raw = fs.readFileSync(viRuPath, "utf8");
        const json = JSON.parse(raw);
        this.viRuMap = new Map();
        this.viRuList = (json.rows || []).map((r, i) => {
          const entry = {
            id: `vi-ru:${i}`,
            vi: r[0],
            pron: r[1] || "",
            pos: r[2] || "",
            cat: r[3] || "",
            ru: r[4] || "",
            ex: r[5] || "",
            gram: r[6] || "",
            dir: "vi-ru",
          };
          const cleanVi = entry.vi.toLowerCase().trim();
          entry._w = this.fold(entry.vi);
          entry._ru = this.fold(entry.ru);
          entry._hay = this.fold(`${entry.vi} ${entry.ru} ${entry.ex}`);
          if (!this.viRuMap.has(cleanVi)) {
            this.viRuMap.set(cleanVi, entry);
          }
          return entry;
        });
      }

      const famPath = path.join(this.dataDir, "dict-fam.json");
      if (fs.existsSync(famPath)) {
        this.famData = JSON.parse(fs.readFileSync(famPath, "utf8"));
      }

      if (this.ruViList && this.ruViList.length > 0) {
        this.isLoaded = true;
        console.log(`[DictionaryService] Initialized: ${this.ruViList.length} Ru-Vi, ${this.viRuList?.length || 0} Vi-Ru words.`);
      }
    } catch (err) {
      console.error("[DictionaryService] Initialization failed:", err);
    }
  }

  isCyrillic(str) {
    return /[\u0400-\u04FF]/.test(str);
  }

  search(query, options = {}) {
    if (!this.isLoaded) this.init();
    const qRaw = String(query || "").trim();
    if (!qRaw) return { total: 0, results: [] };

    let dir = options.dir || "auto";
    if (dir === "auto") {
      dir = this.isCyrillic(qRaw) ? "ru-vi" : "vi-ru";
    }

    const limit = Math.min(Math.max(parseInt(options.limit, 10) || 20, 1), 100);
    const qFold = this.fold(qRaw);
    const qClean = this.stripStress(qRaw).toLowerCase();

    const list = dir === "ru-vi" ? this.ruViList : this.viRuList;
    if (!list || list.length === 0) {
      return { total: 0, results: [], dir };
    }

    const exactMatch = [];
    const prefixMatch = [];
    const wordBoundaryMatch = [];
    const containsMatch = [];
    const meaningMatch = [];

    const wbRegex = new RegExp(`(^|\\s)${qFold.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")}`, "i");

    for (let i = 0; i < list.length; i++) {
      const item = list[i];
      const wordClean = dir === "ru-vi" ? this.stripStress(item.ru).toLowerCase() : item.vi.toLowerCase();
      const wFold = item._w;

      if (wordClean === qClean || wFold === qFold) {
        exactMatch.push(item);
      } else if (wFold.startsWith(qFold)) {
        prefixMatch.push(item);
      } else if (wbRegex.test(wFold)) {
        wordBoundaryMatch.push(item);
      } else if (wFold.includes(qFold)) {
        containsMatch.push(item);
      } else if (item._hay.includes(qFold)) {
        meaningMatch.push(item);
      }

      if (exactMatch.length + prefixMatch.length >= limit * 2) {
        break;
      }
    }

    const combined = [...exactMatch, ...prefixMatch, ...wordBoundaryMatch, ...containsMatch, ...meaningMatch];
    const results = combined.slice(0, limit);

    return {
      total: combined.length,
      dir,
      query: qRaw,
      results,
    };
  }

  getRandomSuggestions(count = 6) {
    if (!this.isLoaded) this.init();
    if (!this.ruViList || this.ruViList.length === 0) return [];
    const pool = this.ruViList.filter((e) => e.ex && e.ex.length > 10 && e.ru.length <= 12 && !e.ru.includes("-"));
    const results = [];
    const used = new Set();
    const targetCount = Math.min(count, pool.length);
    while (results.length < targetCount) {
      const idx = Math.floor(Math.random() * pool.length);
      if (!used.has(idx)) {
        used.add(idx);
        results.push(pool[idx]);
      }
    }
    return results;
  }

  async getDetail(word, dir = "ru-vi") {
    if (!this.isLoaded) this.init();
    const clean = this.stripStress(word).toLowerCase().trim();
    let entry = null;

    if (dir === "ru-vi" && this.ruViMap) {
      entry = this.ruViMap.get(clean) || null;
      if (!entry) {
        const qFold = this.fold(clean);
        entry = this.ruViList ? this.ruViList.find((e) => e._w === qFold) || null : null;
      }
    } else if (this.viRuMap) {
      entry = this.viRuMap.get(clean) || null;
      if (!entry) {
        const qFold = this.fold(clean);
        entry = this.viRuList ? this.viRuList.find((e) => e._w === qFold) || null : null;
      }
    }

    let extra = null;
    const cyrillicWord = dir === "ru-vi" ? clean : entry ? this.stripStress(entry.ru).toLowerCase() : null;

    if (cyrillicWord && this.isCyrillic(cyrillicWord)) {
      const letter = this.cleanLetter(cyrillicWord);
      if (letter) {
        let relMap = this.relCache.get(letter);
        if (!relMap) {
          const relFilename = `dict-rel-${letter}.json`;
          const relPath = path.join(this.dataDir, relFilename);
          if (!fs.existsSync(relPath)) {
            await this.ensureFile(relFilename);
          }
          if (fs.existsSync(relPath)) {
            try {
              relMap = JSON.parse(fs.readFileSync(relPath, "utf8"));
              this.relCache.set(letter, relMap);
            } catch (e) {
              relMap = {};
            }
          } else {
            relMap = {};
          }
        }

        let formsMap = this.formsCache.get(letter);
        if (!formsMap) {
          const formsFilename = `dict-forms-${letter}.json`;
          const formsPath = path.join(this.dataDir, formsFilename);
          if (!fs.existsSync(formsPath)) {
            await this.ensureFile(formsFilename);
          }
          if (fs.existsSync(formsPath)) {
            try {
              formsMap = JSON.parse(fs.readFileSync(formsPath, "utf8"));
              this.formsCache.set(letter, formsMap);
            } catch (e) {
              formsMap = {};
            }
          } else {
            formsMap = {};
          }
        }

        const relItem = relMap ? relMap[cyrillicWord] : null;
        const formsItem = formsMap ? formsMap[cyrillicWord] : null;

        let famGroup = null;
        if (relItem && relItem.fg && this.famData && this.famData.g) {
          famGroup = this.famData.g[relItem.fg] || null;
        }

        extra = {
          forms: formsItem || null,
          syn: relItem?.s || null,
          ant: relItem?.a || null,
          der: relItem?.d || null,
          fam: famGroup || null,
          gov: relItem?.gv || null,
          audio: relItem?.au || null,
        };
      }
    }

    return {
      word,
      dir,
      entry,
      extra,
    };
  }

  async translateOnline(text, from = "auto", to = "vi") {
    const q = String(text || "").trim();
    if (!q) return null;
    try {
      const isRu = this.isCyrillic(q);
      const pair = isRu ? "ru|vi" : "vi|ru";
      const url = `https://api.mymemory.translated.net/get?q=${encodeURIComponent(q)}&langpair=${pair}`;
      const res = await fetch(url, { headers: { "User-Agent": "Mozilla/5.0" } });
      if (!res.ok) return null;
      const json = await res.json();
      if (json && json.responseData && json.responseData.translatedText) {
        return {
          translatedText: json.responseData.translatedText,
          matches: (json.matches || []).slice(0, 3).map((m) => ({
            segment: m.segment,
            translation: m.translation,
          })),
        };
      }
    } catch (e) {
      console.warn("[DictionaryService] Online translate failed:", e.message);
    }
    return null;
  }
}

module.exports = { DictionaryService };
