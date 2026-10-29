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

      if ((this.ruViList && this.ruViList.length > 0) || (this.viRuList && this.viRuList.length > 0)) {
        this.isLoaded = true;
        console.log(`[DictionaryService] Initialized: ${this.ruViList?.length || 0} Ru-Vi, ${this.viRuList?.length || 0} Vi-Ru words.`);
      }
    } catch (err) {
      console.error("[DictionaryService] Initialization failed:", err);
    }
  }

  isCyrillic(str) {
    return /[\u0400-\u04FF]/.test(str);
  }

  getRussianLemmaCandidates(word) {
    const w = this.stripStress(word).toLowerCase().trim().replace(/ё/g, "е");
    if (w.length < 3) return [];
    const candidates = new Set();

    // 1. Dạng chia thì của động từ (verb endings) -> -ать, -ять, -ить, -еть, -овать, -евать, -ти, -ся
    const verbSuffixes = [
      { s: "ает", r: ["ать"] },
      { s: "яет", r: ["ять"] },
      { s: "ует", r: ["овать", "евать"] },
      { s: "юет", r: ["евать"] },
      { s: "аешь", r: ["ать"] },
      { s: "яешь", r: ["ять"] },
      { s: "аем", r: ["ать"] },
      { s: "яем", r: ["ять"] },
      { s: "аете", r: ["ать"] },
      { s: "яете", r: ["ять"] },
      { s: "ают", r: ["ать"] },
      { s: "яют", r: ["ять"] },
      { s: "аю", r: ["ать"] },
      { s: "яю", r: ["ять"] },
      { s: "ит", r: ["ить", "еть"] },
      { s: "ишь", r: ["ить", "еть"] },
      { s: "им", r: ["ить", "еть"] },
      { s: "ите", r: ["ить", "еть"] },
      { s: "ят", r: ["ить", "еть"] },
      { s: "ат", r: ["ать", "ить"] },
      { s: "ал", r: ["ать"] },
      { s: "ала", r: ["ать"] },
      { s: "ало", r: ["ать"] },
      { s: "али", r: ["ать"] },
      { s: "ил", r: ["ить"] },
      { s: "ила", r: ["ить"] },
      { s: "ило", r: ["ить"] },
      { s: "или", r: ["ить"] },
      { s: "ел", r: ["еть"] },
      { s: "ела", r: ["еть"] },
      { s: "ели", r: ["еть"] },
      { s: "ут", r: ["ти", "ть"] },
      { s: "ет", r: ["еть", "ить", "ать"] },
      { s: "ем", r: ["еть", "ить", "ать"] },
      { s: "ешь", r: ["еть", "ить", "ать"] },
    ];

    let base = w;
    let isReflexive = false;
    if (base.endsWith("ся") || base.endsWith("сь")) {
      isReflexive = true;
      base = base.slice(0, -2);
    }

    for (const suf of verbSuffixes) {
      if (base.endsWith(suf.s)) {
        const stem = base.slice(0, -suf.s.length);
        for (const rep of suf.r) {
          candidates.add(stem + rep + (isReflexive ? "ся" : ""));
        }
      }
    }

    // 2. Biến cách danh từ và tính từ (noun & adjective declension endings)
    const nounAdjSuffixes = [
      { s: "ами", r: ["", "а", "я", "о", "е", "ь"] },
      { s: "ями", r: ["", "а", "я", "о", "е", "ь"] },
      { s: "ого", r: ["ый", "ий", "ой"] },
      { s: "его", r: ["ый", "ий", "ой"] },
      { s: "ому", r: ["ый", "ий", "ой"] },
      { s: "ему", r: ["ый", "ий", "ой"] },
      { s: "ыми", r: ["ый", "ий", "ой"] },
      { s: "ими", r: ["ый", "ий", "ой"] },
      { s: "ую", r: ["ая", "яя", "ый"] },
      { s: "юю", r: ["яя", "ий"] },
      { s: "ая", r: ["ый", "ий", "ой"] },
      { s: "яя", r: ["ый", "ий", "ой"] },
      { s: "ое", r: ["ый", "ий", "ой"] },
      { s: "ее", r: ["ый", "ий", "ой"] },
      { s: "ые", r: ["ый", "ий", "ой"] },
      { s: "ие", r: ["ый", "ий", "ой"] },
      { s: "ой", r: ["а", "я", "ь", "", "ый"] },
      { s: "ей", r: ["", "ь", "я", "е"] },
      { s: "ом", r: ["", "а", "о"] },
      { s: "ем", r: ["", "я", "е", "ь"] },
      { s: "ах", r: ["", "а", "о"] },
      { s: "ях", r: ["", "я", "е", "ь"] },
      { s: "ов", r: [""] },
      { s: "ев", r: ["", "ь"] },
      { s: "ам", r: ["", "а", "о"] },
      { s: "ям", r: ["", "я", "е", "ь"] },
      { s: "ы", r: ["", "а"] },
      { s: "и", r: ["", "ь", "я", "а", "е"] },
      { s: "а", r: ["", "о"] },
      { s: "я", r: ["", "е", "ь"] },
      { s: "у", r: ["", "а", "о"] },
      { s: "ю", r: ["", "я", "е", "ь"] },
      { s: "е", r: ["", "а", "о", "я"] },
    ];

    for (const suf of nounAdjSuffixes) {
      if (w.endsWith(suf.s)) {
        const stem = w.slice(0, -suf.s.length);
        for (const rep of suf.r) {
          if (stem.length + rep.length >= 2) {
            candidates.add(stem + rep);
          }
        }
      }
    }

    candidates.delete(w);
    return Array.from(candidates);
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
    const lemmaMatches = [];
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

    // Nếu là từ tiếng Nga và chưa có exact match, tìm dạng từ gốc (lemmatization)
    if (dir === "ru-vi" && exactMatch.length === 0 && this.ruViMap) {
      const candidates = this.getRussianLemmaCandidates(qClean);
      for (const cand of candidates) {
        if (this.ruViMap.has(cand)) {
          const item = this.ruViMap.get(cand);
          if (!lemmaMatches.some((m) => m.id === item.id)) {
            lemmaMatches.push({
              ...item,
              _isLemmaOf: qRaw,
            });
          }
        }
      }
    }

    const combined = [...exactMatch, ...lemmaMatches, ...prefixMatch, ...wordBoundaryMatch, ...containsMatch, ...meaningMatch];
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

  async translateOnline(text, from = "auto", to = null) {
    const q = String(text || "").trim().slice(0, 1000);
    if (!q) return null;
    try {
      const isRu = this.isCyrillic(q);
      const sl = from === "auto" ? (isRu ? "ru" : "vi") : from;
      const tl = to || (sl === "ru" ? "vi" : "ru");

      const ua = "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36";

      // 1. Động cơ 1: Google Chrome Dictionary Client API (kháng rate-limit tuyệt đối)
      try {
        const u1 = `https://clients5.google.com/translate_a/t?client=dict-chrome-ex&sl=${encodeURIComponent(sl)}&tl=${encodeURIComponent(tl)}&q=${encodeURIComponent(q)}`;
        const r1 = await fetch(u1, { headers: { "User-Agent": ua }, signal: AbortSignal.timeout(4000) });
        if (r1.ok) {
          const txt = await r1.text();
          const json = JSON.parse(txt);
          if (Array.isArray(json) && json[0]) {
            const trans = Array.isArray(json[0]) ? json[0][0] : json[0];
            if (trans && typeof trans === "string") {
              return {
                translatedText: trans,
                matches: [{ segment: q, translation: trans }],
                engine: "google",
              };
            }
          }
        }
      } catch (e) {
        // Tiếp tục thử động cơ 2
      }

      // 2. Động cơ 2: Google Translate GTX (dt=t)
      try {
        const gtxUrl = `https://translate.googleapis.com/translate_a/single?client=gtx&sl=${encodeURIComponent(sl)}&tl=${encodeURIComponent(tl)}&dt=t&q=${encodeURIComponent(q)}`;
        const gtxRes = await fetch(gtxUrl, {
          headers: { "User-Agent": ua },
          signal: AbortSignal.timeout(4000),
        });
        if (gtxRes.ok) {
          const txt = await gtxRes.text();
          const gtxJson = JSON.parse(txt);
          const translatedText = gtxJson[0] ? gtxJson[0].map((x) => x[0]).filter(Boolean).join("") : "";
          if (translatedText) {
            return {
              translatedText,
              matches: [{ segment: q, translation: translatedText }],
              engine: "google",
            };
          }
        }
      } catch (gtxErr) {
        // Tiếp tục thử động cơ 3
      }

      // 3. Động cơ 3: MyMemory API với quota 50k ký tự/ngày
      const pair = `${sl}|${tl}`;
      const url = `https://api.mymemory.translated.net/get?q=${encodeURIComponent(q)}&langpair=${encodeURIComponent(pair)}&de=contact@taigatalk.edu.vn`;
      const res = await fetch(url, {
        headers: { "User-Agent": ua },
        signal: AbortSignal.timeout(5000),
      });
      if (res.ok) {
        const json = await res.json();
        if (json && json.responseData && json.responseData.translatedText) {
          const trans = json.responseData.translatedText;
          if (!trans.includes("MYMEMORY WARNING")) {
            return {
              translatedText: trans,
              matches: (json.matches || []).slice(0, 3).map((m) => ({
                segment: m.segment,
                translation: m.translation,
              })),
              engine: "mymemory",
            };
          }
        }
      }
    } catch (e) {
      console.warn("[DictionaryService] Online translate failed completely:", e.message);
    }
    return null;
  }
}

module.exports = { DictionaryService };
