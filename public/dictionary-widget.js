/**
 * TAIGA CENTER - FLOATING CORNER DICTIONARY WIDGET (RUSSIAN - VIETNAMESE)
 * Integrated Russian-Vietnamese & Vietnamese-Russian Dictionary Tool
 */

(function () {
  "use strict";

  const STORAGE_BOOKMARKS = "taiga_dict_bookmarks";
  const STORAGE_HISTORY = "taiga_dict_history";

  const CASE_ROWS = [
    { key: "nomn", ru: "Именительный (Cách 1)", vn: "Danh cách — ai? cái gì?" },
    { key: "gent", ru: "Родительный (Cách 2)", vn: "Sinh cách — của ai? cái gì?" },
    { key: "datv", ru: "Дательный (Cách 3)", vn: "Dữ cách — cho ai? cái gì?" },
    { key: "accs", ru: "Винительный (Cách 4)", vn: "Đối cách — ai? cái gì? (tác động)" },
    { key: "ablt", ru: "Творительный (Cách 5)", vn: "Công cụ cách — bằng gì? với ai?" },
    { key: "loct", ru: "Предложный (Cách 6)", vn: "Giới cách — về ai? ở đâu?" },
  ];

  class ClientDictProvider {
    constructor() {
      this.cache = {};
      this.loading = {};
      this.formsCache = {};
      this.relCache = {};
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

    async load(dir) {
      if (this.cache[dir]) return this.cache[dir];
      if (this.loading[dir]) return this.loading[dir];

      this.loading[dir] = (async () => {
        let json = null;
        try {
          const r1 = await fetch(`/data/dict/dict-${dir}.json`);
          if (r1.ok) json = await r1.json();
        } catch (e) {}

        if (!json) {
          try {
            const r2 = await fetch(`https://rusitrki.com/data/dict-${dir}.json`);
            if (r2.ok) json = await r2.json();
          } catch (e) {}
        }

        if (!json || !json.rows) {
          throw new Error("Cannot load dictionary data");
        }

        const isRuVi = dir === "ru-vi";
        const map = new Map();
        const list = json.rows.map((r, i) => {
          const item = isRuVi
            ? {
                id: `ru-vi:${i}`,
                ru: r[0],
                pron: r[1] || "",
                pos: r[2] || "",
                cat: r[3] || "",
                vi: r[4] || "",
                ex: r[5] || "",
                gram: r[6] || "",
                dir: "ru-vi",
              }
            : {
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
          const wordKey = isRuVi ? this.stripStress(item.ru).toLowerCase() : item.vi.toLowerCase().trim();
          item._w = this.fold(isRuVi ? item.ru : item.vi);
          item._hay = this.fold(`${isRuVi ? item.ru : item.vi} ${item.pron} ${isRuVi ? item.vi : item.ru} ${item.ex}`);
          if (!map.has(wordKey)) map.set(wordKey, item);
          return item;
        });

        this.cache[dir] = { list, map };
        delete this.loading[dir];
        return this.cache[dir];
      })();

      return this.loading[dir];
    }

    async search(query, dir = "auto", limit = 25) {
      try {
        const isRu = /[\u0400-\u04FF]/.test(query);
        const targetDir = dir === "auto" ? (isRu ? "ru-vi" : "vi-ru") : dir;
        const data = await this.load(targetDir);
        const qFold = this.fold(query);
        const qClean = this.stripStress(query).toLowerCase();

        const exact = [];
        const prefix = [];
        const contains = [];
        const wbRegex = new RegExp(`(^|\\s)${qFold.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")}`, "i");

        for (const item of data.list) {
          const cleanW = targetDir === "ru-vi" ? this.stripStress(item.ru).toLowerCase() : item.vi.toLowerCase();
          if (cleanW === qClean || item._w === qFold) {
            exact.push(item);
          } else if (item._w.startsWith(qFold)) {
            prefix.push(item);
          } else if (wbRegex.test(item._w) || item._w.includes(qFold)) {
            contains.push(item);
          }
          if (exact.length + prefix.length >= limit * 2) break;
        }

        const combined = [...exact, ...prefix, ...contains].slice(0, limit);
        return { ok: true, dir: targetDir, results: combined };
      } catch (e) {
        return { ok: false, results: [] };
      }
    }

    async getDetail(word, dir = "ru-vi") {
      try {
        const isRu = /[\u0400-\u04FF]/.test(word);
        const targetDir = dir === "auto" ? (isRu ? "ru-vi" : "vi-ru") : dir;
        const data = await this.load(targetDir);
        const clean = this.stripStress(word).toLowerCase().trim();
        let entry = data.map.get(clean) || null;
        if (!entry) {
          const qFold = this.fold(clean);
          entry = data.list.find((e) => e._w === qFold) || null;
        }

        let extra = null;
        const cyrillic = targetDir === "ru-vi" ? clean : entry ? this.stripStress(entry.ru).toLowerCase() : null;

        if (cyrillic && /[\u0400-\u04FF]/.test(cyrillic)) {
          const letter = cyrillic.replace(/ё/g, "е")[0];
          if (/^[а-я]/.test(letter)) {
            let forms = this.formsCache[letter];
            let rel = this.relCache[letter];

            if (!forms) {
              try {
                const rfLocal = await fetch(`/data/dict/dict-forms-${encodeURIComponent(letter)}.json`);
                if (rfLocal.ok) forms = await rfLocal.json();
              } catch (e) {}
              if (!forms) {
                try {
                  const rf = await fetch(`https://rusitrki.com/data/dict-forms-${encodeURIComponent(letter)}.json`);
                  if (rf.ok) forms = await rf.json();
                } catch (e) {}
              }
              this.formsCache[letter] = forms || {};
            }

            if (!rel) {
              try {
                const rrLocal = await fetch(`/data/dict/dict-rel-${encodeURIComponent(letter)}.json`);
                if (rrLocal.ok) rel = await rrLocal.json();
              } catch (e) {}
              if (!rel) {
                try {
                  const rr = await fetch(`https://rusitrki.com/data/dict-rel-${encodeURIComponent(letter)}.json`);
                  if (rr.ok) rel = await rr.json();
                } catch (e) {}
              }
              this.relCache[letter] = rel || {};
            }

            const formsItem = (forms && forms[cyrillic]) || null;
            const relItem = (rel && rel[cyrillic]) || null;

            extra = {
              forms: formsItem,
              syn: relItem?.s || null,
              ant: relItem?.a || null,
              der: relItem?.d || null,
              audio: relItem?.au || null,
              gov: relItem?.gv || null,
            };
          }
        }

        return { ok: true, word, dir: targetDir, entry, extra };
      } catch (e) {
        return { ok: false, word, dir, entry: null, extra: null };
      }
    }
  }

  class DictionaryWidget {
    constructor() {
      this.isOpen = false;
      this.isExpanded = false;
      this.currentDir = "auto";
      this.activeTab = "search"; // 'search' | 'bookmarks' | 'history'
      this.searchDebounceTimer = null;
      this.currentDetail = null;
      this.audioObj = null;
      this.clientProvider = new ClientDictProvider();

      this.bookmarks = this.loadStorage(STORAGE_BOOKMARKS);
      this.history = this.loadStorage(STORAGE_HISTORY);

      this.initDOM();
      this.bindEvents();
      this.initTextSelectionLookup();
      this.updateBadgeCounts();
      this.loadInitialSuggestions();
    }

    loadStorage(key) {
      try {
        const val = localStorage.getItem(key);
        return val ? JSON.parse(val) : [];
      } catch (e) {
        return [];
      }
    }

    saveStorage(key, val) {
      try {
        localStorage.setItem(key, JSON.stringify(val));
      } catch (e) {}
    }

    escapeHtml(str) {
      return String(str || "")
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&#39;");
    }

    stripStress(str) {
      return (str || "").replace(/\u0301/g, "").trim();
    }

    isCyrillic(str) {
      return /[\u0400-\u04FF]/.test(str);
    }

    formatTimeAgo(timestamp) {
      if (!timestamp) return "";
      const diff = Math.floor((Date.now() - timestamp) / 1000);
      if (diff < 60) return "vừa xong";
      if (diff < 3600) return `${Math.floor(diff / 60)} phút trước`;
      if (diff < 86400) return `${Math.floor(diff / 3600)} giờ trước`;
      if (diff < 604800) return `${Math.floor(diff / 86400)} ngày trước`;
      return new Date(timestamp).toLocaleDateString("vi-VN");
    }

    updateBadgeCounts() {
      const bEl = this.panelEl ? this.panelEl.querySelector("#dictBadgeBookmarks") : null;
      const hEl = this.panelEl ? this.panelEl.querySelector("#dictBadgeHistory") : null;
      if (bEl) bEl.textContent = this.bookmarks.length;
      if (hEl) hEl.textContent = this.history.length;
    }

    initDOM() {
      // 1. Launcher button
      const launcher = document.createElement("button");
      launcher.className = "dict-widget-launcher";
      launcher.id = "dictWidgetLauncher";
      launcher.setAttribute("title", "Mở từ điển Nga - Việt (Alt+D)");
      launcher.innerHTML = `
        <span class="dict-launcher-icon">📖</span>
        <span class="dict-launcher-text">Từ điển Nga ⇄ Việt</span>
        <span class="dict-launcher-badge">81k+</span>
      `;
      document.body.appendChild(launcher);
      this.launcherEl = launcher;

      // 2. Main panel
      const panel = document.createElement("div");
      panel.className = "dict-widget-panel";
      panel.id = "dictWidgetPanel";
      panel.setAttribute("role", "dialog");
      panel.setAttribute("aria-label", "Từ điển Nga - Việt hai chiều");

      panel.innerHTML = `
        <!-- HEADER -->
        <div class="dict-widget-header">
          <div class="dict-header-title-wrap">
            <span class="dict-header-title">📖 Từ điển Nga ⇄ Việt</span>
            <span class="dict-header-count">81.000+ từ</span>
          </div>
          <div class="dict-header-actions">
            <button class="dict-tool-btn" id="dictExpandBtn" title="Phóng to / Thu nhỏ">🗖</button>
            <button class="dict-tool-btn" id="dictCloseBtn" title="Đóng từ điển (Esc)">✕</button>
          </div>
        </div>

        <!-- NAV BAR: AUTO DETECT BADGE + BOOKMARKS & HISTORY TABS -->
        <div class="dict-nav-bar">
          <div class="dict-auto-badge">
            <span class="dict-pulse-dot"></span>
            <span class="dict-auto-label">Tự động nhận diện Nga ⇄ Việt</span>
          </div>
          <div class="dict-tabs-nav">
            <button class="dict-nav-tab" data-tab="bookmarks" id="dictTabBtnBookmarks" title="Sổ tay từ đã lưu">
              ⭐ Đã lưu <span class="dict-tab-badge" id="dictBadgeBookmarks">0</span>
            </button>
            <button class="dict-nav-tab" data-tab="history" id="dictTabBtnHistory" title="Lịch sử tra cứu gần đây">
              🕒 Lịch sử <span class="dict-tab-badge" id="dictBadgeHistory">0</span>
            </button>
          </div>
        </div>

        <!-- SEARCH INPUT CONTAINER -->
        <div class="dict-search-container" id="dictSearchContainer">
          <div class="dict-search-box">
            <span class="dict-search-icon">🔍</span>
            <input 
              type="text" 
              class="dict-search-input" 
              id="dictSearchInput" 
              placeholder="Nhập từ hoặc câu tiếng Nga / tiếng Việt..."
              autocomplete="off"
              spellcheck="false"
            />
            <button class="dict-clear-btn" id="dictClearBtn" title="Xóa">✕</button>
          </div>
          <!-- Cyrillic Quick Keys -->
          <div class="dict-cyrillic-bar" id="dictCyrillicBar">
            <button class="dict-key-btn" data-char="ё">ё</button>
            <button class="dict-key-btn" data-char="й">й</button>
            <button class="dict-key-btn" data-char="ъ">ъ</button>
            <button class="dict-key-btn" data-char="ы">ы</button>
            <button class="dict-key-btn" data-char="э">э</button>
            <button class="dict-key-btn" data-char="щ">щ</button>
            <button class="dict-key-btn" data-char="ц">ц</button>
            <button class="dict-key-btn" data-char="ж">ж</button>
            <button class="dict-key-btn" data-char="ш">ш</button>
            <button class="dict-key-btn" data-char="ю">ю</button>
            <button class="dict-key-btn" data-char="я">я</button>
          </div>
        </div>

        <!-- BODY CONTENT -->
        <div class="dict-widget-body" id="dictWidgetBody">
          <div id="dictTabSearch" class="dict-tab-pane">
            <div id="dictSearchResults" class="dict-results-list"></div>
            <div id="dictEmptyState" class="dict-empty-state">
              <span class="dict-empty-icon">📚</span>
              <p class="dict-empty-title">Tra cứu nhanh từ vựng & chia động từ</p>
              <p class="dict-empty-sub">Gõ từ tiếng Nga/Việt cần tìm hoặc bôi đen văn bản trên bài viết để tra cứu ngay lập tức.</p>
              <div id="dictSuggestionsGrid" class="dict-suggestions-grid"></div>
            </div>
          </div>
          <div id="dictTabBookmarks" class="dict-tab-pane" style="display:none;"></div>
          <div id="dictTabHistory" class="dict-tab-pane" style="display:none;"></div>
        </div>
      `;

      document.body.appendChild(panel);
      this.panelEl = panel;
      this.searchInput = panel.querySelector("#dictSearchInput");
      this.clearBtn = panel.querySelector("#dictClearBtn");
      this.bodyEl = panel.querySelector("#dictWidgetBody");
      this.resultsEl = panel.querySelector("#dictSearchResults");
      this.emptyEl = panel.querySelector("#dictEmptyState");

      // 3. Selection quick tooltip popup
      const popup = document.createElement("div");
      popup.className = "dict-selection-popup";
      popup.id = "dictSelectionPopup";
      popup.innerHTML = `<span>📖 Tra từ:</span> <strong id="dictPopupWord"></strong>`;
      document.body.appendChild(popup);
      this.popupEl = popup;
    }

    bindEvents() {
      // Toggle widget
      this.launcherEl.addEventListener("click", () => this.toggle());
      this.panelEl.querySelector("#dictCloseBtn").addEventListener("click", () => this.close());
      this.panelEl.querySelector("#dictExpandBtn").addEventListener("click", () => this.toggleExpand());

      // Keyboard hotkeys
      document.addEventListener("keydown", (e) => {
        if (e.key === "Escape" && this.isOpen) {
          this.close();
        }
        if (e.altKey && (e.key === "d" || e.key === "D")) {
          e.preventDefault();
          this.toggle();
        }
      });

      // Tabs: Bookmarks & History
      this.panelEl.querySelectorAll(".dict-nav-tab").forEach((tab) => {
        tab.addEventListener("click", () => {
          const targetTab = tab.dataset.tab;
          if (this.activeTab === targetTab) {
            this.switchTab("search");
          } else {
            this.switchTab(targetTab);
          }
        });
      });

      // When clicking or focusing on search input while on Bookmarks/History, return to Search
      this.searchInput.addEventListener("focus", () => {
        if (this.activeTab !== "search") {
          this.switchTab("search");
        }
      });

      // Search input & debouncing
      this.searchInput.addEventListener("input", () => {
        if (this.activeTab !== "search") {
          this.switchTab("search");
        }
        const val = this.searchInput.value;
        this.clearBtn.style.display = val ? "grid" : "none";
        clearTimeout(this.searchDebounceTimer);
        this.searchDebounceTimer = setTimeout(() => {
          this.handleSearch(val.trim());
        }, 280);
      });

      this.searchInput.addEventListener("keydown", (e) => {
        if (e.key === "Enter") {
          e.preventDefault();
          clearTimeout(this.searchDebounceTimer);
          const val = this.searchInput.value.trim();
          if (val) {
            this.handleSearch(val, true);
          }
        }
      });

      this.clearBtn.addEventListener("click", () => {
        this.searchInput.value = "";
        this.clearBtn.style.display = "none";
        this.resultsEl.innerHTML = "";
        this.emptyEl.style.display = "flex";
        this.searchInput.focus();
      });

      // Cyrillic quick key buttons
      this.panelEl.querySelectorAll(".dict-key-btn").forEach((btn) => {
        btn.addEventListener("click", () => {
          if (this.activeTab !== "search") {
            this.switchTab("search");
          }
          const char = btn.dataset.char;
          const input = this.searchInput;
          const start = input.selectionStart || input.value.length;
          const end = input.selectionEnd || input.value.length;
          input.value = input.value.substring(0, start) + char + input.value.substring(end);
          input.selectionStart = input.selectionEnd = start + char.length;
          input.focus();
          input.dispatchEvent(new Event("input"));
        });
      });
    }

    initTextSelectionLookup() {
      let selectedText = "";

      document.addEventListener("mouseup", (e) => {
        if (this.panelEl.contains(e.target) || this.launcherEl.contains(e.target)) {
          return;
        }

        setTimeout(() => {
          const selection = window.getSelection();
          const text = selection ? selection.toString().trim() : "";

          if (text && text.length >= 2 && text.length <= 50 && !text.includes("\n")) {
            selectedText = text;
            const range = selection.getRangeAt(0);
            const rect = range.getBoundingClientRect();

            this.popupEl.querySelector("#dictPopupWord").textContent = text;
            this.popupEl.style.top = `${window.scrollY + rect.top - 6}px`;
            this.popupEl.style.left = `${window.scrollX + rect.left + rect.width / 2}px`;
            this.popupEl.style.display = "flex";
          } else {
            this.popupEl.style.display = "none";
          }
        }, 50);
      });

      document.addEventListener("mousedown", (e) => {
        if (!this.popupEl.contains(e.target)) {
          this.popupEl.style.display = "none";
        }
      });

      this.popupEl.addEventListener("click", () => {
        if (selectedText) {
          this.open();
          this.switchTab("search");
          this.searchInput.value = selectedText;
          this.clearBtn.style.display = "grid";
          this.handleSearch(selectedText);
          this.popupEl.style.display = "none";
        }
      });
    }

    open() {
      this.isOpen = true;
      this.panelEl.classList.add("open");
      this.launcherEl.style.transform = "scale(0.9)";
      this.launcherEl.style.opacity = "0.7";
      setTimeout(() => this.searchInput.focus(), 150);
    }

    close() {
      this.isOpen = false;
      this.panelEl.classList.remove("open");
      this.launcherEl.style.transform = "";
      this.launcherEl.style.opacity = "";
      if (this.audioObj) {
        this.audioObj.pause();
      }
      if (window.TaigaAudio) {
        window.TaigaAudio.stop();
      }
      if (window.speechSynthesis) {
        speechSynthesis.cancel();
      }
    }

    toggle() {
      if (this.isOpen) {
        this.close();
      } else {
        this.open();
      }
    }

    toggleExpand() {
      this.isExpanded = !this.isExpanded;
      this.panelEl.classList.toggle("expanded", this.isExpanded);
      this.panelEl.querySelector("#dictExpandBtn").textContent = this.isExpanded ? "🗗" : "🗖";
    }

    switchTab(tab) {
      this.activeTab = tab;
      this.panelEl.querySelectorAll(".dict-nav-tab").forEach((t) => {
        t.classList.toggle("active", t.dataset.tab === tab);
      });

      const paneSearch = this.panelEl.querySelector("#dictTabSearch");
      const paneBookmarks = this.panelEl.querySelector("#dictTabBookmarks");
      const paneHistory = this.panelEl.querySelector("#dictTabHistory");

      paneSearch.style.display = tab === "search" ? "block" : "none";
      paneBookmarks.style.display = tab === "bookmarks" ? "block" : "none";
      paneHistory.style.display = tab === "history" ? "block" : "none";

      if (tab === "bookmarks") {
        this.renderBookmarks();
      } else if (tab === "history") {
        this.renderHistory();
      }
    }

    async loadInitialSuggestions() {
      try {
        const res = await fetch("/api/dict/suggestions");
        const data = await res.json();
        if (data && data.ok && Array.isArray(data.suggestions)) {
          const grid = this.panelEl.querySelector("#dictSuggestionsGrid");
          grid.innerHTML = data.suggestions
            .map(
              (s) => `
            <div class="dict-sug-item" data-word="${this.escapeHtml(s.ru)}">
              <span class="dict-sug-word">${this.escapeHtml(s.ru)}</span>
              <span class="dict-sug-vi">${this.escapeHtml(s.vi.split("\n")[0])}</span>
            </div>
          `
            )
            .join("");

          grid.querySelectorAll(".dict-sug-item").forEach((item) => {
            item.addEventListener("click", () => {
              const word = item.dataset.word;
              this.searchInput.value = word;
              this.clearBtn.style.display = "grid";
              this.openWordDetail(word, "ru-vi");
            });
          });
        }
      } catch (e) {}
    }

    async handleSearch(query, directOpen = false) {
      if (!query) {
        this.resultsEl.innerHTML = "";
        this.emptyEl.style.display = "flex";
        return;
      }

      this.emptyEl.style.display = "none";
      this.resultsEl.innerHTML = `
        <div style="text-align: center; padding: 24px; color: #64748b;">
          <span>⏳ Đang tìm kiếm...</span>
        </div>
      `;

      let data = null;
      // 1. Thử gọi backend trước
      const direction = this.currentDir;
      try {
        const res = await fetch(`/api/dict/search?q=${encodeURIComponent(query)}&dir=${direction}&limit=25`);
        if (res.ok) {
          data = await res.json();
        }
      } catch (e) {}

      // Only load the browser-side dictionary if the backend did not respond.
      if (!data || !Array.isArray(data.results)) {
        try {
          data = await this.clientProvider.search(query, direction, 25);
        } catch (e) {}
      }

      if (data && data.results && data.results.length > 0) {
        const firstItem = data.results[0];
        const firstWord = data.dir === "ru-vi" ? firstItem.ru : firstItem.vi;
        const firstMeaning = (data.dir === "ru-vi" ? firstItem.vi : firstItem.ru) || "";

        const qClean = this.stripStress(query).toLowerCase().trim();
        const topClean = this.stripStress(firstWord).toLowerCase().trim();

        // Nếu ấn Enter hoặc từ tìm kiếm khớp chính xác 100% với từ đầu tiên, mở thẳng chi tiết (kèm chia động từ)
        if (directOpen || qClean === topClean) {
          this.openWordDetail(firstWord, data.dir);
        } else {
          this.addToHistory(query, firstMeaning, data.dir);
          this.renderSearchResults(data.results, query, data.dir);
        }
      } else {
        this.renderOnlineFallback(query, direction);
      }
    }

    renderSearchResults(results, query, dir) {
      const isRuVi = dir === "ru-vi";
      let html = "";

      results.forEach((item) => {
        const primaryWord = isRuVi ? item.ru : item.vi;
        const subWord = isRuVi ? item.vi.split("\n")[0] : item.ru;
        const pron = item.pron ? `<span class="dict-card-pron">${this.escapeHtml(item.pron)}</span>` : "";
        const pos = item.pos ? `<span class="dict-card-pos">${this.escapeHtml(item.pos)}</span>` : "";
        const exFirst = item.ex ? item.ex.split("\n")[0] : "";
        const lemmaBadge = item._isLemmaOf
          ? `<span class="dict-card-lemma-badge" style="background:#e0f2fe; color:#0369a1; font-size:0.75rem; padding:2px 7px; border-radius:4px; font-weight:600; margin-left:6px;" title="Dạng nguyên mẫu của từ '${this.escapeHtml(item._isLemmaOf)}'">Dạng gốc của "${this.escapeHtml(item._isLemmaOf)}"</span>`
          : "";

        html += `
          <div class="dict-result-card" data-word="${this.escapeHtml(item.ru || item.vi)}" data-dir="${item.dir || dir}">
            <div class="dict-card-head">
              <div style="display: flex; align-items: baseline; gap: 6px; flex-wrap: wrap;">
                <span class="dict-card-word">${this.escapeHtml(primaryWord)}</span>
                ${pron}
                ${lemmaBadge}
              </div>
              ${pos}
            </div>
            <div class="dict-card-meaning">${this.escapeHtml(subWord)}</div>
            ${exFirst ? `<div class="dict-card-ex-preview">${this.escapeHtml(exFirst)}</div>` : ""}
          </div>
        `;
      });

      html += `
        <div class="dict-online-shortcut" style="text-align: center; margin-top: 14px; padding: 10px; background: #f8fafc; border: 1px dashed #cbd5e1; border-radius: 8px;">
          <button id="dictBtnForceTranslate" style="background: none; border: none; color: #2563eb; font-weight: 600; font-size: 0.85rem; cursor: pointer; display: inline-flex; align-items: center; gap: 6px;">
            <span>🌐</span> Dịch trực tuyến nguyên văn "<strong>${this.escapeHtml(query)}</strong>" ➔
          </button>
        </div>
      `;

      this.resultsEl.innerHTML = html;

      const forceBtn = this.resultsEl.querySelector("#dictBtnForceTranslate");
      if (forceBtn) {
        forceBtn.addEventListener("click", () => {
          this.renderOnlineFallback(query, dir);
        });
      }

      this.resultsEl.querySelectorAll(".dict-result-card").forEach((card) => {
        card.addEventListener("click", () => {
          const w = card.dataset.word;
          const d = card.dataset.dir;
          this.openWordDetail(w, d);
        });
      });
    }

    async renderOnlineFallback(query, direction = this.currentDir) {
      this.resultsEl.innerHTML = `
        <div class="dict-fallback-card">
          <div class="dict-fallback-badge">🌐 Dịch tự động trực tuyến</div>
          <div class="dict-fallback-text">Đang dịch từ "<strong>${this.escapeHtml(query)}</strong>"...</div>
        </div>
      `;

      let trans = null;
      let engineLabel = "Google Translate";
      let matches = [];

      // 1. Thử gọi backend /api/dict/translate
      try {
        const sourceLanguage = direction === "ru-vi" ? "ru" : direction === "vi-ru" ? "vi" : "auto";
        const res = await fetch(`/api/dict/translate?text=${encodeURIComponent(query)}&from=${sourceLanguage}`, {
          signal: AbortSignal.timeout(3500),
        });
        if (res.ok) {
          const data = await res.json();
          if (data && data.ok && data.result && data.result.translatedText) {
            trans = data.result.translatedText;
            engineLabel = data.result.engine === "mymemory" ? "Song ngữ AI" : "Google Translate";
            matches = data.result.matches || [];
          }
        }
      } catch (e) {}

      if (trans) {
        const isRu = direction === "auto" ? this.isCyrillic(query) : direction === "ru-vi";
        const dir = isRu ? "ru-vi" : "vi-ru";
        this.addToHistory(query, trans, dir);

        const isSaved = this.bookmarks.some((b) => b.word.toLowerCase() === query.toLowerCase());

        this.resultsEl.innerHTML = `
          <div class="dict-fallback-card">
            <div class="dict-fallback-badge">🌐 Dịch trực tuyến (${engineLabel})</div>
            <h3 style="margin: 6px 0 4px 0; font-size: 1.25rem; color: #1e3a8a;">${this.escapeHtml(query)}</h3>
            <div style="font-size: 1.05rem; font-weight: 700; color: #047857; margin-top: 6px; line-height: 1.4;">➔ ${this.escapeHtml(trans)}</div>
            ${
              matches.length > 1
                ? `<div style="margin-top: 10px; font-size: 0.82rem; color: #64748b; background: #f8fafc; padding: 8px 10px; border-radius: 6px;">
                    <b>Các bản dịch khác:</b>
                    <ul style="margin: 4px 0 0 16px; padding: 0;">
                      ${matches
                        .slice(1)
                        .map((m) => `<li>${this.escapeHtml(m.translation)}</li>`)
                        .join("")}
                    </ul>
                  </div>`
                : ""
            }
            <div style="margin-top: 12px; display: flex; gap: 8px; flex-wrap: wrap;">
              <button class="dict-audio-btn" id="dictFallbackSpeakBtn" title="Nghe phát âm" style="display: inline-flex; align-items: center; gap: 4px; padding: 6px 12px; border-radius: 6px; font-size: 0.82rem; cursor: pointer;">🔊 Nghe</button>
              <button class="dict-action-btn" id="dictFallbackCopyBtn" title="Sao chép bản dịch" style="border: 1px solid #cbd5e1; background: #fff; padding: 6px 12px; border-radius: 6px; font-size: 0.82rem; cursor: pointer;">📋 Sao chép</button>
              <button class="dict-action-btn ${isSaved ? "active" : ""}" id="dictFallbackBookmarkBtn" title="Lưu vào sổ tay" style="border: 1px solid ${isSaved ? "#ef4444" : "#cbd5e1"}; background: ${isSaved ? "#fef2f2" : "#fff"}; color: ${isSaved ? "#ef4444" : "#475569"}; padding: 6px 12px; border-radius: 6px; font-size: 0.82rem; cursor: pointer; font-weight: 700;">
                ${isSaved ? "❤️ Đã lưu" : "🤍 Lưu vào sổ tay"}
              </button>
            </div>
          </div>
        `;

        const spkBtn = this.resultsEl.querySelector("#dictFallbackSpeakBtn");
        if (spkBtn) {
          spkBtn.addEventListener("click", () => this.speak(isRu ? query : trans));
        }

        const copyBtn = this.resultsEl.querySelector("#dictFallbackCopyBtn");
        if (copyBtn) {
          copyBtn.addEventListener("click", () => {
            if (navigator.clipboard) {
              navigator.clipboard.writeText(trans).then(() => {
                copyBtn.textContent = "✓ Đã chép";
                setTimeout(() => {
                  copyBtn.textContent = "📋 Sao chép";
                }, 2000);
              });
            }
          });
        }

        const bmBtn = this.resultsEl.querySelector("#dictFallbackBookmarkBtn");
        if (bmBtn) {
          bmBtn.addEventListener("click", () => {
            const saved = this.toggleBookmark({ word: query, meaning: trans, dir });
            bmBtn.innerHTML = saved ? "❤️ Đã lưu" : "🤍 Lưu vào sổ tay";
            bmBtn.style.borderColor = saved ? "#ef4444" : "#cbd5e1";
            bmBtn.style.background = saved ? "#fef2f2" : "#fff";
            bmBtn.style.color = saved ? "#ef4444" : "#475569";
          });
        }
      } else {
        this.resultsEl.innerHTML = `
          <div class="dict-empty-state">
            <span class="dict-empty-icon">🌐</span>
            <p class="dict-empty-title">Dịch trực tuyến chưa phản hồi</p>
            <p class="dict-empty-sub">Vui lòng thử lại sau.</p>
            <button class="dict-action-btn" id="dictBtnRetryTranslate" type="button">Thử dịch lại</button>
          </div>
        `;
        this.resultsEl.querySelector("#dictBtnRetryTranslate").addEventListener("click", () => {
          this.renderOnlineFallback(query, direction);
        });
      }
    }

    async openWordDetail(word, dir = "ru-vi") {
      this.resultsEl.innerHTML = `
        <div style="text-align: center; padding: 30px; color: #64748b;">
          <span>⏳ Đang tải thông tin chi tiết...</span>
        </div>
      `;
      this.emptyEl.style.display = "none";

      let data = null;
      try {
        const res = await fetch(`/api/dict/detail?word=${encodeURIComponent(word)}&dir=${dir}`);
        if (res.ok) {
          data = await res.json();
        }
      } catch (e) {}

      if (!data || !data.entry) {
        try {
          data = await this.clientProvider.getDetail(word, dir);
        } catch (e) {}
      }

      if (data && data.entry) {
        const firstM = data.entry ? data.entry.vi || data.entry.ru : "";
        this.addToHistory(word, firstM, data.dir || dir);
        this.currentDetail = data;
        this.renderDetailView(data);
      } else {
        this.handleSearch(word, false);
      }
    }

    // Check if word is a verb
    isVerb(mainWord, entry, extra) {
      const w = this.stripStress(mainWord).toLowerCase();
      if (extra && extra.forms && extra.forms.v) return true;
      if (entry) {
        if (entry.pos && entry.pos.toLowerCase().includes("động từ")) return true;
        if (entry.gram && (entry.gram.includes("гл.") || entry.gram.includes("несов.") || entry.gram.includes("сов."))) return true;
        if (entry.vi && (entry.vi.includes("động từ") || entry.vi.includes("тк. несов.") || entry.vi.includes("тк. сов."))) return true;
      }
      if (w.endsWith("ть") || w.endsWith("ти") || w.endsWith("чь") || w.endsWith("ться") || w.endsWith("тись")) return true;
      return false;
    }

    // Build comprehensive Verb Conjugation (thì, thời, thể)
    buildVerbConjugationSection(mainWord, formsV, entry, extra) {
      const inf = this.stripStress(mainWord).toLowerCase();
      const isReflexive = inf.endsWith("ся") || inf.endsWith("сь");
      const baseInf = isReflexive ? inf.slice(0, -2) : inf;

      // 1. Thể (Aspect)
      let isPerf = false;
      if (formsV && typeof formsV.perf === "number") {
        isPerf = formsV.perf === 1;
      } else if (entry && entry.gram) {
        if (entry.gram.includes("сов.") && !entry.gram.includes("несов.")) {
          isPerf = true;
        }
      }

      const aspectBadgeClass = isPerf ? "perf" : "impf";
      const aspectTitle = isPerf ? "★ Hoàn thành thể (СВ · Совершенный вид)" : "● Chưa hoàn thành thể (НСВ · Несовершенный вид)";
      const aspectDesc = isPerf
        ? "Diễn tả hành động đã hoàn tất hoặc sẽ hoàn tất trọn vẹn (không có thì hiện tại)."
        : "Diễn tả hành động diễn ra thường xuyên, đang diễn ra hoặc nhấn mạnh quá trình.";

      // 2. Tìm cặp thể tương ứng (Видовая пара)
      let pairWord = "";
      if (entry && entry.vi) {
        const m = entry.vi.match(/(?:см\.|соверш\.|несоверш\.|сов\.|несов\.)\s+([а-яё́]+)/i);
        if (m && m[1]) pairWord = this.stripStress(m[1]).toLowerCase();
      }
      if (!pairWord && extra && extra.der) {
        const commonPrefixes = ["по", "про", "с", "на", "за", "вы", "при", "от", "пере", "у"];
        for (const d of extra.der) {
          const cleanD = this.stripStress(d).toLowerCase();
          if (cleanD !== inf && (cleanD.endsWith("ть") || cleanD.endsWith("ться"))) {
            if (!isPerf && commonPrefixes.some((p) => cleanD === p + inf)) {
              pairWord = cleanD;
              break;
            } else if (isPerf && commonPrefixes.some((p) => inf === p + cleanD)) {
              pairWord = cleanD;
              break;
            }
          }
        }
      }

      // 3. Hiện tại / Tương lai đơn (6 ngôi)
      let pres = formsV && Array.isArray(formsV.pres) && formsV.pres.length >= 6 ? formsV.pres : null;
      if (!pres) {
        // Tự suy luận dạng chia tiêu chuẩn
        let stem = baseInf.endsWith("ать") || baseInf.endsWith("ять") || baseInf.endsWith("еть") ? baseInf.slice(0, -2) : baseInf.slice(0, -3);
        if (baseInf.endsWith("ить")) {
          stem = baseInf.slice(0, -3);
          pres = [stem + "ю", stem + "ишь", stem + "ит", stem + "им", stem + "ите", stem + "ят"].map((w) => w + (isReflexive ? "ся" : ""));
        } else {
          pres = [stem + "ю", stem + "ешь", stem + "ет", stem + "ем", stem + "ете", stem + "ют"].map((w) => w + (isReflexive ? "ся" : ""));
        }
      }

      // 4. Quá khứ (4 dạng)
      let past = formsV && formsV.past ? formsV.past : null;
      if (!past || !past.m) {
        let pStem = baseInf.endsWith("ть") ? baseInf.slice(0, -2) : baseInf;
        past = {
          m: pStem + "л" + (isReflexive ? "ся" : ""),
          f: pStem + "ла" + (isReflexive ? "сь" : ""),
          n: pStem + "ло" + (isReflexive ? "сь" : ""),
          p: pStem + "ли" + (isReflexive ? "сь" : ""),
        };
      }

      // 5. Mệnh lệnh (Повелительное наклонение)
      let imp = formsV && formsV.imp && (formsV.imp.s || formsV.imp.p) ? formsV.imp : null;
      if (!imp || (!imp.s && !imp.p)) {
        if (pres && pres[5]) {
          const pStem = pres[5].replace(/(ют|ут|ят|ат)(ся|сь)?$/, "");
          const endsVowel = /[аеёиоуыэюя]$/i.test(pStem);
          let s = endsVowel ? pStem + "й" : pStem + "и";
          let p = endsVowel ? pStem + "йте" : pStem + "ите";
          if (isReflexive) {
            s += endsVowel ? "ся" : "сь";
            p += "сь";
          }
          imp = { s, p };
        } else {
          imp = { s: inf + "й", p: inf + "йте" };
        }
      }

      // 6. Tương lai phức (Будущее сложное) cho НСВ
      const compoundFuture = [
        `буду ${mainWord}`,
        `будешь ${mainWord}`,
        `будет ${mainWord}`,
        `будем ${mainWord}`,
        `будете ${mainWord}`,
        `будут ${mainWord}`,
      ];

      const persons = ["я (tôi)", "ты (bạn)", "он / она / оно", "мы (chúng tôi)", "вы (các bạn)", "они (họ)"];

      return `
        <div class="dict-verb-container">
          <!-- 1. THỂ & CẶP THỂ ĐỐI ỨNG -->
          <div class="dict-verb-aspect-box">
            <div>
              <span class="dict-aspect-badge ${aspectBadgeClass}">${aspectTitle}</span>
              <p style="font-size:0.75rem; color:#475569; margin:4px 0 0 2px;">${aspectDesc}</p>
            </div>
            ${
              pairWord
                ? `<button class="dict-pair-link" data-pair="${this.escapeHtml(pairWord)}" title="Chuyển sang xem thể đối ứng">
                    🔁 Cặp thể đối ứng: <b>${this.escapeHtml(pairWord)}</b> (${isPerf ? "НСВ" : "СВ"})
                   </button>`
                : ""
            }
          </div>

          <!-- 2. THÌ HIỆN TẠI (cho НСВ) hoặc TƯƠNG LAI ĐƠN (cho СВ) -->
          <div class="dict-verb-block">
            <h4 class="dict-verb-block-title">
              <span>${isPerf ? "⚡ THÌ TƯƠNG LAI ĐƠN (Будущее простое время)" : "⏰ THÌ HIỆN TẠI (Настоящее время)"}</span>
              <span class="dict-verb-block-subtitle">${isPerf ? "Chia 6 ngôi (Động từ СВ)" : "Chia 6 ngôi (Động từ НСВ)"}</span>
            </h4>
            <div class="dict-verb-grid">
              ${persons
                .map(
                  (p, i) => `
                <div class="dict-verb-row">
                  <span class="dict-verb-person">${p}</span>
                  <span class="dict-verb-form">
                    ${this.escapeHtml(pres[i] || "—")}
                    ${pres[i] ? `<button class="dict-verb-audio-btn" data-vspeak="${this.escapeHtml(pres[i])}" title="Nghe phát âm">🔊</button>` : ""}
                  </span>
                </div>
              `
                )
                .join("")}
            </div>
          </div>

          <!-- 3. THÌ TƯƠNG LAI PHỨC (cho НСВ) -->
          ${
            !isPerf
              ? `
            <div class="dict-verb-block">
              <h4 class="dict-verb-block-title">
                <span>🚀 THÌ TƯƠNG LAI PHỨC (Будущее сложное время)</span>
                <span class="dict-verb-block-subtitle">буду + động từ nguyên thể</span>
              </h4>
              <div class="dict-verb-grid">
                ${persons
                  .map(
                    (p, i) => `
                  <div class="dict-verb-row">
                    <span class="dict-verb-person">${p}</span>
                    <span class="dict-verb-form">
                      ${this.escapeHtml(compoundFuture[i])}
                      <button class="dict-verb-audio-btn" data-vspeak="${this.escapeHtml(compoundFuture[i])}" title="Nghe phát âm">🔊</button>
                    </span>
                  </div>
                `
                  )
                  .join("")}
              </div>
            </div>
          `
              : ""
          }

          <!-- 4. THÌ QUÁ KHỨ (Прошедшее время) -->
          <div class="dict-verb-block">
            <h4 class="dict-verb-block-title">
              <span>🕰️ THÌ QUÁ KHỨ (Прошедшее время)</span>
              <span class="dict-verb-block-subtitle">Theo Giống (Đực, Cái, Trung) & Số nhiều</span>
            </h4>
            <div class="dict-verb-grid">
              <div class="dict-verb-row">
                <span class="dict-verb-person">он (giống đực)</span>
                <span class="dict-verb-form">
                  ${this.escapeHtml(past.m || "—")}
                  ${past.m ? `<button class="dict-verb-audio-btn" data-vspeak="${this.escapeHtml(past.m)}" title="Nghe phát âm">🔊</button>` : ""}
                </span>
              </div>
              <div class="dict-verb-row">
                <span class="dict-verb-person">она (giống cái)</span>
                <span class="dict-verb-form">
                  ${this.escapeHtml(past.f || "—")}
                  ${past.f ? `<button class="dict-verb-audio-btn" data-vspeak="${this.escapeHtml(past.f)}" title="Nghe phát âm">🔊</button>` : ""}
                </span>
              </div>
              <div class="dict-verb-row">
                <span class="dict-verb-person">оно (giống trung)</span>
                <span class="dict-verb-form">
                  ${this.escapeHtml(past.n || "—")}
                  ${past.n ? `<button class="dict-verb-audio-btn" data-vspeak="${this.escapeHtml(past.n)}" title="Nghe phát âm">🔊</button>` : ""}
                </span>
              </div>
              <div class="dict-verb-row">
                <span class="dict-verb-person">они / мы / вы (số nhiều)</span>
                <span class="dict-verb-form">
                  ${this.escapeHtml(past.p || "—")}
                  ${past.p ? `<button class="dict-verb-audio-btn" data-vspeak="${this.escapeHtml(past.p)}" title="Nghe phát âm">🔊</button>` : ""}
                </span>
              </div>
            </div>
          </div>

          <!-- 5. THỂ MỆNH LỆNH (Повелительное наклонение) -->
          <div class="dict-verb-block">
            <h4 class="dict-verb-block-title">
              <span>📣 THỂ MỆNH LỆNH (Повелительное наклонение)</span>
              <span class="dict-verb-block-subtitle">Yêu cầu, khuyên bảo hoặc ra lệnh</span>
            </h4>
            <div class="dict-verb-grid">
              <div class="dict-verb-row">
                <span class="dict-verb-person">ты (bạn - thân mật)</span>
                <span class="dict-verb-form">
                  ${this.escapeHtml(imp.s || "—")}
                  ${imp.s ? `<button class="dict-verb-audio-btn" data-vspeak="${this.escapeHtml(imp.s)}" title="Nghe phát âm">🔊</button>` : ""}
                </span>
              </div>
              <div class="dict-verb-row">
                <span class="dict-verb-person">вы (các bạn / lịch sự)</span>
                <span class="dict-verb-form">
                  ${this.escapeHtml(imp.p || "—")}
                  ${imp.p ? `<button class="dict-verb-audio-btn" data-vspeak="${this.escapeHtml(imp.p)}" title="Nghe phát âm">🔊</button>` : ""}
                </span>
              </div>
            </div>
          </div>

          <!-- 6. THỂ GIẢ ĐỊNH / ĐIỀU KIỆN (Сослагательное наклонение) -->
          <div class="dict-verb-block">
            <h4 class="dict-verb-block-title">
              <span>💡 THỂ GIẢ ĐỊNH / ĐIỀU KIỆN (Сослагательное наклонение)</span>
              <span class="dict-verb-block-subtitle">Dạng quá khứ + tiểu từ "бы" (giá như, nếu...)</span>
            </h4>
            <div class="dict-verb-grid">
              <div class="dict-verb-row">
                <span class="dict-verb-person">он (nam)</span>
                <span class="dict-verb-form">
                  ${this.escapeHtml(past.m)} бы
                  <button class="dict-verb-audio-btn" data-vspeak="${this.escapeHtml(past.m)} бы" title="Nghe phát âm">🔊</button>
                </span>
              </div>
              <div class="dict-verb-row">
                <span class="dict-verb-person">она (nữ)</span>
                <span class="dict-verb-form">
                  ${this.escapeHtml(past.f)} бы
                  <button class="dict-verb-audio-btn" data-vspeak="${this.escapeHtml(past.f)} бы" title="Nghe phát âm">🔊</button>
                </span>
              </div>
              <div class="dict-verb-row">
                <span class="dict-verb-person">они (số nhiều)</span>
                <span class="dict-verb-form">
                  ${this.escapeHtml(past.p)} бы
                  <button class="dict-verb-audio-btn" data-vspeak="${this.escapeHtml(past.p)} бы" title="Nghe phát âm">🔊</button>
                </span>
              </div>
            </div>
          </div>
        </div>
      `;
    }

    renderDetailView(data) {
      const entry = data.entry || { ru: data.word, vi: "" };
      const extra = data.extra || {};
      const isRuVi = data.dir === "ru-vi";
      const mainWord = isRuVi ? entry.ru : entry.vi;
      const isBookmarked = this.bookmarks.some((b) => b.word.toLowerCase() === mainWord.toLowerCase());

      const meaningLines = (entry.vi || "").split("\n").filter((l) => l.trim());
      const firstMeaning = meaningLines[0] || "";

      const sectionsHtml = [];

      // 1. Meanings list
      if (meaningLines.length > 0) {
        sectionsHtml.push(`
          <div class="dict-sec open" data-sec="meanings">
            <button class="dict-sec-head" type="button">
              <span class="dict-sec-head-left">
                <span class="dict-sec-icon">📖</span> Nghĩa tiếng Việt (${meaningLines.length})
              </span>
              <span class="dict-sec-arrow">▾</span>
            </button>
            <div class="dict-sec-body">
              ${meaningLines
                .map(
                  (m, idx) => `
                <div class="dict-meaning-item">
                  <span class="dict-meaning-bullet">${idx + 1}.</span> ${this.escapeHtml(m)}
                </div>
              `
                )
                .join("")}
            </div>
          </div>
        `);
      }

      // 2. Comprehensive Verb Conjugation (thì, thời, thể)
      const isWordVerb = isRuVi && this.isVerb(mainWord, entry, extra);
      if (isWordVerb) {
        const verbConjugationHtml = this.buildVerbConjugationSection(mainWord, extra.forms?.v, entry, extra);
        sectionsHtml.push(`
          <div class="dict-sec open" data-sec="verb-conjugation">
            <button class="dict-sec-head" type="button">
              <span class="dict-sec-head-left">
                <span class="dict-sec-icon">⏱️</span> Bảng chia động từ toàn diện (Thì, Thời, Thể)
              </span>
              <span class="dict-sec-arrow">▾</span>
            </button>
            <div class="dict-sec-body">
              ${verbConjugationHtml}
            </div>
          </div>
        `);
      }

      // 3. Examples
      if (entry.ex && entry.ex.trim()) {
        const exLines = entry.ex.split("\n").filter((l) => l.trim()).slice(0, 10);
        sectionsHtml.push(`
          <div class="dict-sec open" data-sec="examples">
            <button class="dict-sec-head" type="button">
              <span class="dict-sec-head-left">
                <span class="dict-sec-icon">💬</span> Câu ví dụ thực tế (${exLines.length})
              </span>
              <span class="dict-sec-arrow">▾</span>
            </button>
            <div class="dict-sec-body">
              ${exLines
                .map((line) => {
                  const arrowIdx = line.indexOf("→");
                  let ruPart = line;
                  let viPart = "";
                  if (arrowIdx !== -1) {
                    ruPart = line.substring(0, arrowIdx).trim();
                    viPart = line.substring(arrowIdx + 1).trim();
                  }
                  return `
                    <div class="dict-ex-item">
                      <div class="dict-ex-ru">
                        <span>${this.escapeHtml(ruPart)}</span>
                        <button class="dict-ex-speak-btn" data-speak="${this.escapeHtml(ruPart)}" title="Nghe câu này">🔊</button>
                      </div>
                      ${viPart ? `<div class="dict-ex-vi">→ ${this.escapeHtml(viPart)}</div>` : ""}
                    </div>
                  `;
                })
                .join("")}
            </div>
          </div>
        `);
      }

      // 4. Grammar: 6 Cases (Declension for Nouns & Adjectives)
      if (extra.forms && (extra.forms.n || extra.forms.a || extra.forms.pn)) {
        const tableHtml = this.buildDeclensionTable(extra.forms);
        if (tableHtml) {
          sectionsHtml.push(`
            <div class="dict-sec open" data-sec="declension">
              <button class="dict-sec-head" type="button">
                <span class="dict-sec-head-left">
                  <span class="dict-sec-icon">📊</span> Bảng biến cách 6 Cách (Падежи)
                </span>
                <span class="dict-sec-arrow">▾</span>
              </button>
              <div class="dict-sec-body">
                ${tableHtml}
              </div>
            </div>
          `);
        }
      }

      // 5. Word Family (Họ từ & Từ phái sinh)
      const derWords = extra.der || [];
      const famWords = extra.fam || [];
      const allRelatives = [...new Set([...derWords, ...famWords])].slice(0, 20);
      if (allRelatives.length > 0) {
        sectionsHtml.push(`
          <div class="dict-sec" data-sec="family">
            <button class="dict-sec-head" type="button">
              <span class="dict-sec-head-left">
                <span class="dict-sec-icon">🌳</span> Họ từ & Từ liên quan (${allRelatives.length})
              </span>
              <span class="dict-sec-arrow">▾</span>
            </button>
            <div class="dict-sec-body">
              <div class="dict-chips-wrap">
                ${allRelatives
                  .map(
                    (w) => `
                    <button class="dict-word-chip" data-lookup="${this.escapeHtml(this.stripStress(w))}">${this.escapeHtml(w)}</button>
                  `
                  )
                  .join("")}
              </div>
            </div>
          </div>
        `);
      }

      // 6. Synonyms & Antonyms
      if (extra.syn || extra.ant) {
        sectionsHtml.push(`
          <div class="dict-sec" data-sec="synant">
            <button class="dict-sec-head" type="button">
              <span class="dict-sec-head-left">
                <span class="dict-sec-icon">↔️</span> Đồng nghĩa · Trái nghĩa
              </span>
              <span class="dict-sec-arrow">▾</span>
            </button>
            <div class="dict-sec-body">
              ${
                extra.syn && extra.syn.length > 0
                  ? `<div style="margin-bottom:8px;">
                      <span style="font-size:0.75rem; font-weight:700; color:#166534;">🟢 Đồng nghĩa:</span>
                      <div class="dict-chips-wrap" style="margin-top:4px;">
                        ${extra.syn.map((s) => `<button class="dict-word-chip syn" data-lookup="${this.escapeHtml(s)}">${this.escapeHtml(s)}</button>`).join("")}
                      </div>
                    </div>`
                  : ""
              }
              ${
                extra.ant && extra.ant.length > 0
                  ? `<div>
                      <span style="font-size:0.75rem; font-weight:700; color:#991b1b;">🔴 Trái nghĩa:</span>
                      <div class="dict-chips-wrap" style="margin-top:4px;">
                        ${extra.ant.map((a) => `<button class="dict-word-chip ant" data-lookup="${this.escapeHtml(a)}">${this.escapeHtml(a)}</button>`).join("")}
                      </div>
                    </div>`
                  : ""
              }
            </div>
          </div>
        `);
      }

      const detailHtml = `
        <div class="dict-detail-view">
          <button class="dict-back-btn" id="dictBackBtn">← Quay lại danh sách kết quả</button>

          <!-- HERO -->
          <div class="dict-detail-hero">
            <div class="dict-hero-tags">
              ${entry.pos ? `<span class="dict-tag-pos">${this.escapeHtml(entry.pos)}</span>` : ""}
              ${entry.gram ? `<span class="dict-tag-gram">${this.escapeHtml(entry.gram)}</span>` : ""}
            </div>

            <div class="dict-hero-word-wrap">
              <h2 class="dict-hero-word">${this.escapeHtml(mainWord)}</h2>
              <div class="dict-hero-btn-group">
                <button class="dict-audio-btn" id="dictDetailSpeak" title="Nghe phát âm chuẩn">🔊</button>
                <button class="dict-audio-btn slow" id="dictDetailSlow" title="Nghe phát âm chậm">🐢</button>
                <button class="dict-bookmark-btn ${isBookmarked ? "active" : ""}" id="dictDetailBookmark" title="${isBookmarked ? "Bỏ lưu từ" : "Lưu vào sổ tay"}">
                  ${isBookmarked ? "❤️" : "🤍"}
                </button>
              </div>
            </div>

            ${entry.pron ? `<p class="dict-hero-pron">${this.escapeHtml(entry.pron)}</p>` : ""}
            ${firstMeaning ? `<p class="dict-hero-meaning-preview">${this.escapeHtml(firstMeaning)}</p>` : ""}
          </div>

          <!-- SECTIONS -->
          <div class="dict-sections-wrap">
            ${sectionsHtml.join("")}
          </div>
        </div>
      `;

      this.resultsEl.innerHTML = detailHtml;

      // Event listeners for detail view
      const backBtn = this.resultsEl.querySelector("#dictBackBtn");
      if (backBtn) {
        backBtn.addEventListener("click", () => {
          this.handleSearch(this.searchInput.value.trim());
        });
      }

      // Audio buttons
      const audioUrl = extra.audio ? extra.audio.m || extra.audio.o : null;
      const speakBtn = this.resultsEl.querySelector("#dictDetailSpeak");
      const slowBtn = this.resultsEl.querySelector("#dictDetailSlow");

      if (speakBtn) {
        speakBtn.addEventListener("click", () => this.playAudio(mainWord, audioUrl, 1));
      }
      if (slowBtn) {
        slowBtn.addEventListener("click", () => this.playAudio(mainWord, audioUrl, 0.6));
      }

      // Sentence speak buttons
      this.resultsEl.querySelectorAll(".dict-ex-speak-btn").forEach((btn) => {
        btn.addEventListener("click", (e) => {
          e.stopPropagation();
          this.speak(btn.dataset.speak);
        });
      });

      // Verb audio buttons
      this.resultsEl.querySelectorAll(".dict-verb-audio-btn").forEach((btn) => {
        btn.addEventListener("click", (e) => {
          e.stopPropagation();
          this.speak(btn.dataset.vspeak);
        });
      });

      // Verb aspect pair links
      this.resultsEl.querySelectorAll(".dict-pair-link").forEach((btn) => {
        btn.addEventListener("click", () => {
          const pair = btn.dataset.pair;
          this.searchInput.value = pair;
          this.clearBtn.style.display = "grid";
          this.openWordDetail(pair, "ru-vi");
        });
      });

      // Bookmark button
      const bookmarkBtn = this.resultsEl.querySelector("#dictDetailBookmark");
      if (bookmarkBtn) {
        bookmarkBtn.addEventListener("click", () => {
          const saved = this.toggleBookmark(entry, data.dir);
          bookmarkBtn.classList.toggle("active", saved);
          bookmarkBtn.innerHTML = saved ? "❤️" : "🤍";
        });
      }

      // Accordion toggles
      this.resultsEl.querySelectorAll(".dict-sec-head").forEach((head) => {
        head.addEventListener("click", () => {
          head.parentElement.classList.toggle("open");
        });
      });

      // Relative word chips
      this.resultsEl.querySelectorAll(".dict-word-chip").forEach((chip) => {
        chip.addEventListener("click", () => {
          const w = chip.dataset.lookup;
          this.searchInput.value = w;
          this.clearBtn.style.display = "grid";
          this.openWordDetail(w, "ru-vi");
        });
      });

      // Soft auto-play
      setTimeout(() => {
        this.playAudio(mainWord, audioUrl, 1, true);
      }, 200);
    }

    buildDeclensionTable(forms) {
      if (!forms) return "";
      let rows = [];
      let headers = [];

      if (forms.n) {
        headers = ["Cách", "Số ít (Ед.ч)", "Số nhiều (Мн.ч)"];
        rows = CASE_ROWS.map((c, i) => [
          `<span class="case-title">${c.ru}</span><span class="case-sub">${c.vn}</span>`,
          forms.n.s ? forms.n.s[i] || "—" : "—",
          forms.n.p ? forms.n.p[i] || "—" : "—",
        ]);
      } else if (forms.a) {
        headers = ["Cách", "Giống đực", "Giống cái", "Giống trung", "Số nhiều"];
        rows = CASE_ROWS.map((c, i) => [
          `<span class="case-title">${c.ru}</span><span class="case-sub">${c.vn}</span>`,
          forms.a.m ? forms.a.m[i] || "—" : "—",
          forms.a.f ? forms.a.f[i] || "—" : "—",
          forms.a.n ? forms.a.n[i] || "—" : "—",
          forms.a.p ? forms.a.p[i] || "—" : "—",
        ]);
      } else if (forms.pn) {
        headers = ["Cách", "Số ít", "Số nhiều"];
        rows = CASE_ROWS.map((c, i) => [
          `<span class="case-title">${c.ru}</span><span class="case-sub">${c.vn}</span>`,
          forms.pn.s ? forms.pn.s[i] || "—" : "—",
          forms.pn.p ? forms.pn.p[i] || "—" : "—",
        ]);
      }

      if (rows.length === 0) return "";

      return `
        <div class="dict-table-wrap">
          <table class="dict-table">
            <thead>
              <tr>${headers.map((h) => `<th>${this.escapeHtml(h)}</th>`).join("")}</tr>
            </thead>
            <tbody>
              ${rows
                .map(
                  (r) => `
                <tr>
                  <td>${r[0]}</td>
                  ${r
                    .slice(1)
                    .map((val) => `<td class="ru">${this.escapeHtml(val)}</td>`)
                    .join("")}
                </tr>
              `
                )
                .join("")}
            </tbody>
          </table>
        </div>
      `;
    }

    playAudio(text, audioUrl, speed = 1, silentFail = false) {
      if (audioUrl) {
        if (!this.audioObj) this.audioObj = new Audio();
        this.audioObj.src = audioUrl;
        this.audioObj.playbackRate = speed;
        this.audioObj.play().catch(() => {
          this.speak(text, speed);
        });
      } else {
        this.speak(text, speed);
      }
    }

    speak(text, rate = 0.9) {
      const clean = this.stripStress(text).trim();
      if (!clean) return;

      const isRu = this.isCyrillic(clean);
      if (isRu && window.TaigaAudio) {
        window.TaigaAudio.speak(clean, { rate: rate });
        return;
      }
      const langCode = isRu ? "ru-RU" : "vi-VN";
      const ttsLang = isRu ? "ru" : "vi";

      const playTTSFallback = () => {
        try {
          if (this.audioObj) {
            this.audioObj.pause();
          }
          const chunk = clean.slice(0, 200);
          const ttsUrl = `https://translate.google.com/translate_tts?ie=UTF-8&client=tw-ob&tl=${ttsLang}&q=${encodeURIComponent(chunk)}`;
          this.audioObj = new Audio(ttsUrl);
          this.audioObj.playbackRate = rate;
          this.audioObj.play().catch(() => {});
        } catch (e) {}
      };

      if (!("speechSynthesis" in window)) {
        playTTSFallback();
        return;
      }

      try {
        if (speechSynthesis.paused) speechSynthesis.resume();
        speechSynthesis.cancel();
      } catch (e) {}

      const voices = speechSynthesis.getVoices() || [];
      const matchedVoice = voices.find((v) => (isRu ? v.lang.startsWith("ru") : v.lang.startsWith("vi")));

      if (!matchedVoice) {
        playTTSFallback();
        return;
      }

      setTimeout(() => {
        try {
          if (speechSynthesis.paused) speechSynthesis.resume();
          const utter = new SpeechSynthesisUtterance(clean);
          utter.rate = rate;
          utter.lang = langCode;
          utter.voice = matchedVoice;

          let hasStarted = false;
          utter.onstart = () => {
            hasStarted = true;
          };
          utter.onerror = () => {
            playTTSFallback();
          };

          speechSynthesis.speak(utter);

          setTimeout(() => {
            if (!hasStarted) {
              try {
                speechSynthesis.cancel();
              } catch (e) {}
              playTTSFallback();
            }
          }, 700);
        } catch (e) {
          playTTSFallback();
        }
      }, 50);
    }

    toggleBookmark(item, dir = "auto") {
      const word = String(item.ru || item.vi || item.word || "").trim();
      if (!word) return false;
      const meaning = String(item.vi || item.ru || item.meaning || "").split("\n")[0].trim();

      const idx = this.bookmarks.findIndex((b) => b.word.toLowerCase() === word.toLowerCase());
      let isSaved = false;

      if (idx !== -1) {
        this.bookmarks.splice(idx, 1);
        isSaved = false;
      } else {
        this.bookmarks.unshift({
          id: "b_" + Date.now() + "_" + Math.random().toString(36).slice(2, 6),
          word,
          meaning,
          pron: item.pron || "",
          pos: item.pos || "",
          dir: dir || item.dir || "auto",
          date: Date.now(),
        });
        if (this.bookmarks.length > 150) this.bookmarks.pop();
        isSaved = true;
      }

      this.saveStorage(STORAGE_BOOKMARKS, this.bookmarks);
      this.updateBadgeCounts();
      return isSaved;
    }

    addToHistory(word, meaning = "", dir = "auto") {
      const clean = String(word || "").trim();
      if (!clean || clean.length < 2) return;
      this.history = this.history.filter((h) => h.word.toLowerCase() !== clean.toLowerCase());
      this.history.unshift({
        id: "h_" + Date.now() + "_" + Math.random().toString(36).slice(2, 6),
        word: clean,
        meaning: String(meaning || "").split("\n")[0].trim(),
        dir: dir || "auto",
        date: Date.now(),
      });
      if (this.history.length > 60) this.history.pop();
      this.saveStorage(STORAGE_HISTORY, this.history);
      this.updateBadgeCounts();
    }

    renderBookmarks() {
      const pane = this.panelEl.querySelector("#dictTabBookmarks");
      if (this.bookmarks.length === 0) {
        pane.innerHTML = `
          <div class="dict-pane-top-bar">
            <button class="dict-back-to-search-btn" id="dictBtnBackFromBookmarks">← Quay lại tra từ</button>
            <span style="font-weight:700; font-size:0.85rem; color:#64748b;">Sổ tay trống</span>
            <span></span>
          </div>
          <div class="dict-empty-state">
            <span class="dict-empty-icon">⭐</span>
            <p class="dict-empty-title">Chưa có từ nào được lưu</p>
            <p class="dict-empty-sub">Nhấn biểu tượng trái tim (❤️) khi tra từ để lưu vào sổ tay ôn tập.</p>
          </div>
        `;
        const backBtn = pane.querySelector("#dictBtnBackFromBookmarks");
        if (backBtn) backBtn.addEventListener("click", () => this.switchTab("search"));
        return;
      }

      pane.innerHTML = `
        <div class="dict-pane-top-bar">
          <button class="dict-back-to-search-btn" id="dictBtnBackFromBookmarks">← Quay lại tra từ</button>
          <span style="font-weight:700; font-size:0.88rem; color:#1e293b;">Từ đã lưu (${this.bookmarks.length})</span>
          <button class="dict-clear-all-btn" id="dictClearBookmarksBtn">Xóa tất cả</button>
        </div>
        <div class="dict-results-list">
          ${this.bookmarks
            .map(
              (b) => `
            <div class="dict-list-item-card" data-word="${this.escapeHtml(b.word)}" data-dir="${this.escapeHtml(b.dir)}">
              <div class="dict-list-item-left">
                <div class="dict-list-item-word-wrap">
                  <span class="dict-list-item-word">${this.escapeHtml(b.word)}</span>
                  ${b.pron ? `<span style="font-size:0.8rem; color:#64748b;">${this.escapeHtml(b.pron)}</span>` : ""}
                </div>
                ${b.meaning ? `<div class="dict-list-item-meaning">${this.escapeHtml(b.meaning)}</div>` : ""}
                <div class="dict-list-item-time">${this.formatTimeAgo(b.date)}</div>
              </div>
              <div class="dict-list-item-actions">
                <button class="dict-item-btn speak" data-speak="${this.escapeHtml(b.word)}" title="Nghe phát âm">🔊</button>
                <button class="dict-item-btn del" data-del="${this.escapeHtml(b.word)}" title="Bỏ lưu">🗑️</button>
              </div>
            </div>
          `
            )
            .join("")}
        </div>
      `;

      pane.querySelector("#dictBtnBackFromBookmarks")?.addEventListener("click", () => this.switchTab("search"));

      pane.querySelectorAll(".dict-list-item-card").forEach((card) => {
        card.addEventListener("click", (e) => {
          if (e.target.closest(".dict-item-btn")) return;
          const w = card.dataset.word;
          const d = card.dataset.dir || "auto";
          this.switchTab("search");
          this.searchInput.value = w;
          this.clearBtn.style.display = "grid";
          this.openWordDetail(w, d);
        });
      });

      pane.querySelectorAll(".dict-item-btn.speak").forEach((btn) => {
        btn.addEventListener("click", (e) => {
          e.stopPropagation();
          this.speak(btn.dataset.speak);
        });
      });

      pane.querySelectorAll(".dict-item-btn.del").forEach((btn) => {
        btn.addEventListener("click", (e) => {
          e.stopPropagation();
          const targetWord = btn.dataset.del;
          this.bookmarks = this.bookmarks.filter((b) => b.word !== targetWord);
          this.saveStorage(STORAGE_BOOKMARKS, this.bookmarks);
          this.updateBadgeCounts();
          this.renderBookmarks();
        });
      });

      const clearAll = pane.querySelector("#dictClearBookmarksBtn");
      if (clearAll) {
        clearAll.addEventListener("click", () => {
          if (confirm("Bạn có chắc muốn xóa tất cả từ đã lưu?")) {
            this.bookmarks = [];
            this.saveStorage(STORAGE_BOOKMARKS, this.bookmarks);
            this.updateBadgeCounts();
            this.renderBookmarks();
          }
        });
      }
    }

    renderHistory() {
      const pane = this.panelEl.querySelector("#dictTabHistory");
      if (this.history.length === 0) {
        pane.innerHTML = `
          <div class="dict-pane-top-bar">
            <button class="dict-back-to-search-btn" id="dictBtnBackFromHistory">← Quay lại tra từ</button>
            <span style="font-weight:700; font-size:0.85rem; color:#64748b;">Lịch sử trống</span>
            <span></span>
          </div>
          <div class="dict-empty-state">
            <span class="dict-empty-icon">🕒</span>
            <p class="dict-empty-title">Lịch sử tra cứu trống</p>
            <p class="dict-empty-sub">Các từ bạn tra cứu sẽ xuất hiện tại đây để tiện xem lại.</p>
          </div>
        `;
        const backBtn = pane.querySelector("#dictBtnBackFromHistory");
        if (backBtn) backBtn.addEventListener("click", () => this.switchTab("search"));
        return;
      }

      pane.innerHTML = `
        <div class="dict-pane-top-bar">
          <button class="dict-back-to-search-btn" id="dictBtnBackFromHistory">← Quay lại tra từ</button>
          <span style="font-weight:700; font-size:0.88rem; color:#1e293b;">Lịch sử tra cứu (${this.history.length})</span>
          <button class="dict-clear-all-btn" id="dictClearHistoryBtn">Xóa lịch sử</button>
        </div>
        <div class="dict-results-list">
          ${this.history
            .map(
              (h) => `
            <div class="dict-list-item-card" data-word="${this.escapeHtml(h.word)}" data-dir="${this.escapeHtml(h.dir)}">
              <div class="dict-list-item-left">
                <span class="dict-list-item-word">${this.escapeHtml(h.word)}</span>
                ${h.meaning ? `<div class="dict-list-item-meaning">${this.escapeHtml(h.meaning)}</div>` : ""}
                <div class="dict-list-item-time">${this.formatTimeAgo(h.date)}</div>
              </div>
              <div class="dict-list-item-actions">
                <button class="dict-item-btn speak" data-speak="${this.escapeHtml(h.word)}" title="Nghe phát âm">🔊</button>
                <button class="dict-item-btn del" data-del="${this.escapeHtml(h.word)}" title="Xóa khỏi lịch sử">🗑️</button>
              </div>
            </div>
          `
            )
            .join("")}
        </div>
      `;

      pane.querySelector("#dictBtnBackFromHistory")?.addEventListener("click", () => this.switchTab("search"));

      pane.querySelectorAll(".dict-list-item-card").forEach((card) => {
        card.addEventListener("click", (e) => {
          if (e.target.closest(".dict-item-btn")) return;
          const w = card.dataset.word;
          this.switchTab("search");
          this.searchInput.value = w;
          this.clearBtn.style.display = "grid";
          this.handleSearch(w);
        });
      });

      pane.querySelectorAll(".dict-item-btn.speak").forEach((btn) => {
        btn.addEventListener("click", (e) => {
          e.stopPropagation();
          this.speak(btn.dataset.speak);
        });
      });

      pane.querySelectorAll(".dict-item-btn.del").forEach((btn) => {
        btn.addEventListener("click", (e) => {
          e.stopPropagation();
          const targetWord = btn.dataset.del;
          this.history = this.history.filter((h) => h.word !== targetWord);
          this.saveStorage(STORAGE_HISTORY, this.history);
          this.updateBadgeCounts();
          this.renderHistory();
        });
      });

      const clearHist = pane.querySelector("#dictClearHistoryBtn");
      if (clearHist) {
        clearHist.addEventListener("click", () => {
          if (confirm("Bạn có chắc muốn xóa toàn bộ lịch sử tra cứu?")) {
            this.history = [];
            this.saveStorage(STORAGE_HISTORY, this.history);
            this.updateBadgeCounts();
            this.renderHistory();
          }
        });
      }
    }
  }

  // Auto initialize on DOM ready
  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", () => {
      window.TaigaDictWidget = new DictionaryWidget();
    });
  } else {
    window.TaigaDictWidget = new DictionaryWidget();
  }
})();
