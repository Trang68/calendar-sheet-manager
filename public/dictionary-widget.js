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

  class DictionaryWidget {
    constructor() {
      this.isOpen = false;
      this.isExpanded = false;
      this.currentDir = "auto";
      this.activeTab = "search"; // 'search' | 'bookmarks' | 'history'
      this.searchDebounceTimer = null;
      this.currentDetail = null;
      this.audioObj = null;

      this.bookmarks = this.loadStorage(STORAGE_BOOKMARKS);
      this.history = this.loadStorage(STORAGE_HISTORY);

      this.initDOM();
      this.bindEvents();
      this.initTextSelectionLookup();
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
      } catch (e) {
        // Ignore quota
      }
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

    initDOM() {
      // 1. Launcher button
      const launcher = document.createElement("button");
      launcher.className = "dict-widget-launcher";
      launcher.id = "dictWidgetLauncher";
      launcher.setAttribute("title", "Mở từ điển Nga - Việt (Alt+D)");
      launcher.innerHTML = `
        <span class="dict-launcher-icon">📖</span>
        <span class="dict-launcher-text">Từ điển Nga — Việt</span>
        <span class="dict-launcher-badge">81k+</span>
      `;
      document.body.appendChild(launcher);
      this.launcherEl = launcher;

      // 2. Main panel
      const panel = document.createElement("div");
      panel.className = "dict-widget-panel";
      panel.id = "dictWidgetPanel";
      panel.setAttribute("role", "dialog");
      panel.setAttribute("aria-label", "Từ điển Nga - Việt");

      panel.innerHTML = `
        <!-- HEADER -->
        <div class="dict-widget-header">
          <div class="dict-header-title-wrap">
            <span class="dict-header-title">📖 Từ điển Nga — Việt</span>
            <span class="dict-header-count">81.000+ từ</span>
          </div>
          <div class="dict-header-actions">
            <button class="dict-tool-btn" id="dictExpandBtn" title="Phóng to / Thu nhỏ">🗖</button>
            <button class="dict-tool-btn" id="dictCloseBtn" title="Đóng từ điển (Esc)">✕</button>
          </div>
        </div>

        <!-- NAV & DIRECTION BAR -->
        <div class="dict-nav-bar">
          <div class="dict-dir-switch" id="dictDirSwitch">
            <button class="dict-dir-btn active" data-dir="auto" title="Tự động nhận diện ngôn ngữ">Tự động</button>
            <button class="dict-dir-btn" data-dir="ru-vi" title="Nga sang Việt">🇷🇺 ➔ 🇻🇳</button>
            <button class="dict-dir-btn" data-dir="vi-ru" title="Việt sang Nga">🇻🇳 ➔ 🇷🇺</button>
          </div>
          <div class="dict-tabs-nav">
            <button class="dict-nav-tab active" data-tab="search">🔍 Tra</button>
            <button class="dict-nav-tab" data-tab="bookmarks">⭐ Lưu</button>
            <button class="dict-nav-tab" data-tab="history">🕒 Lịch sử</button>
          </div>
        </div>

        <!-- SEARCH INPUT -->
        <div class="dict-search-container" id="dictSearchContainer">
          <div class="dict-search-box">
            <span class="dict-search-icon">🔍</span>
            <input 
              type="text" 
              class="dict-search-input" 
              id="dictSearchInput" 
              placeholder="Nhập từ tiếng Nga hoặc tiếng Việt..."
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
              <p class="dict-empty-title">Tra cứu nhanh từ vựng Nga - Việt</p>
              <p class="dict-empty-sub">Gõ từ bạn cần tìm hoặc bôi đen từ trên bài viết để tra cứu ngay lập tức.</p>
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

      // Direction buttons
      this.panelEl.querySelectorAll(".dict-dir-btn").forEach((btn) => {
        btn.addEventListener("click", () => {
          this.panelEl.querySelectorAll(".dict-dir-btn").forEach((b) => b.classList.remove("active"));
          btn.classList.add("active");
          this.currentDir = btn.dataset.dir;
          if (this.searchInput.value.trim()) {
            this.handleSearch(this.searchInput.value.trim());
          }
        });
      });

      // Tabs (Tra cứu / Lưu / Lịch sử)
      this.panelEl.querySelectorAll(".dict-nav-tab").forEach((tab) => {
        tab.addEventListener("click", () => {
          this.switchTab(tab.dataset.tab);
        });
      });

      // Search input & debouncing
      this.searchInput.addEventListener("input", () => {
        const val = this.searchInput.value;
        this.clearBtn.style.display = val ? "grid" : "none";
        clearTimeout(this.searchDebounceTimer);
        this.searchDebounceTimer = setTimeout(() => {
          this.handleSearch(val.trim());
        }, 180);
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
        // Ignore clicks inside the dictionary widget itself
        if (this.panelEl.contains(e.target) || this.launcherEl.contains(e.target)) {
          return;
        }

        setTimeout(() => {
          const selection = window.getSelection();
          const text = selection ? selection.toString().trim() : "";

          // Check if selection is valid single word or short phrase (< 40 chars)
          if (text && text.length >= 2 && text.length <= 40 && !text.includes("\n")) {
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
      const searchContainer = this.panelEl.querySelector("#dictSearchContainer");

      paneSearch.style.display = tab === "search" ? "block" : "none";
      paneBookmarks.style.display = tab === "bookmarks" ? "block" : "none";
      paneHistory.style.display = tab === "history" ? "block" : "none";

      if (tab === "bookmarks") {
        searchContainer.style.display = "none";
        this.renderBookmarks();
      } else if (tab === "history") {
        searchContainer.style.display = "none";
        this.renderHistory();
      } else {
        searchContainer.style.display = "flex";
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
      } catch (e) {
        // Ignore fallback
      }
    }

    async handleSearch(query) {
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

      try {
        const dir = this.currentDir;
        const res = await fetch(`/api/dict/search?q=${encodeURIComponent(query)}&dir=${dir}&limit=25`);
        const data = await res.json();

        if (data && data.ok && data.results && data.results.length > 0) {
          this.renderSearchResults(data.results, query, data.dir);
        } else {
          // If no results in dataset, attempt online fallback translation
          this.renderOnlineFallback(query);
        }
      } catch (err) {
        this.resultsEl.innerHTML = `
          <div style="text-align: center; padding: 20px; color: #ef4444;">
            <p>Có lỗi khi tra từ. Vui lòng thử lại.</p>
          </div>
        `;
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

        html += `
          <div class="dict-result-card" data-word="${this.escapeHtml(item.ru || item.vi)}" data-dir="${item.dir || dir}">
            <div class="dict-card-head">
              <div style="display: flex; align-items: baseline; gap: 6px;">
                <span class="dict-card-word">${this.escapeHtml(primaryWord)}</span>
                ${pron}
              </div>
              ${pos}
            </div>
            <div class="dict-card-meaning">${this.escapeHtml(subWord)}</div>
            ${exFirst ? `<div class="dict-card-ex-preview">${this.escapeHtml(exFirst)}</div>` : ""}
          </div>
        `;
      });

      this.resultsEl.innerHTML = html;

      this.resultsEl.querySelectorAll(".dict-result-card").forEach((card) => {
        card.addEventListener("click", () => {
          const w = card.dataset.word;
          const d = card.dataset.dir;
          this.openWordDetail(w, d);
        });
      });
    }

    async renderOnlineFallback(query) {
      this.resultsEl.innerHTML = `
        <div class="dict-fallback-card">
          <div class="dict-fallback-badge">🌐 Tra cứu trực tuyến</div>
          <div class="dict-fallback-text">Đang dịch từ "<strong>${this.escapeHtml(query)}</strong>"...</div>
        </div>
      `;

      try {
        const res = await fetch(`/api/dict/translate?text=${encodeURIComponent(query)}`);
        const data = await res.json();
        if (data && data.ok && data.result) {
          const trans = data.result.translatedText || "Không tìm thấy kết quả.";
          this.resultsEl.innerHTML = `
            <div class="dict-fallback-card">
              <div class="dict-fallback-badge">🌐 Dịch tự động song ngữ</div>
              <h3 style="margin: 4px 0; font-size: 1.2rem; color: #1e3a8a;">${this.escapeHtml(query)}</h3>
              <div style="font-size: 1rem; font-weight: 700; color: #047857; margin-top: 4px;">➔ ${this.escapeHtml(trans)}</div>
              ${
                data.result.matches && data.result.matches.length > 1
                  ? `<div style="margin-top: 8px; font-size: 0.8rem; color: #64748b;">
                      <b>Các bản dịch khác:</b>
                      <ul style="margin: 4px 0 0 16px; padding: 0;">
                        ${data.result.matches
                          .slice(1)
                          .map((m) => `<li>${this.escapeHtml(m.translation)}</li>`)
                          .join("")}
                      </ul>
                    </div>`
                  : ""
              }
              <div style="margin-top: 10px; display: flex; gap: 8px;">
                <button class="dict-audio-btn" id="dictFallbackSpeakBtn" title="Phát âm">🔊</button>
              </div>
            </div>
          `;
          const spkBtn = this.resultsEl.querySelector("#dictFallbackSpeakBtn");
          if (spkBtn) {
            spkBtn.addEventListener("click", () => this.speak(query));
          }
        } else {
          this.resultsEl.innerHTML = `
            <div class="dict-empty-state">
              <span class="dict-empty-icon">🔍</span>
              <p class="dict-empty-title">Không tìm thấy từ "${this.escapeHtml(query)}"</p>
              <p class="dict-empty-sub">Hãy kiểm tra chính tả hoặc thử tìm một từ tương tự.</p>
            </div>
          `;
        }
      } catch (e) {
        this.resultsEl.innerHTML = `
          <div class="dict-empty-state">
            <span class="dict-empty-icon">🔍</span>
            <p class="dict-empty-title">Không tìm thấy từ "${this.escapeHtml(query)}"</p>
          </div>
        `;
      }
    }

    async openWordDetail(word, dir = "ru-vi") {
      this.resultsEl.innerHTML = `
        <div style="text-align: center; padding: 30px; color: #64748b;">
          <span>⏳ Đang tải thông tin chi tiết...</span>
        </div>
      `;
      this.emptyEl.style.display = "none";

      // Save to search history
      this.addToHistory(word, dir);

      try {
        const res = await fetch(`/api/dict/detail?word=${encodeURIComponent(word)}&dir=${dir}`);
        const data = await res.json();
        if (data && data.ok) {
          this.currentDetail = data;
          this.renderDetailView(data);
        } else {
          this.handleSearch(word);
        }
      } catch (e) {
        this.handleSearch(word);
      }
    }

    renderDetailView(data) {
      const entry = data.entry || { ru: data.word, vi: "" };
      const extra = data.extra || {};
      const isRuVi = data.dir === "ru-vi";
      const mainWord = isRuVi ? entry.ru : entry.vi;
      const subWord = isRuVi ? entry.vi : entry.ru;
      const isBookmarked = this.bookmarks.some((b) => b.word === mainWord);

      const meaningLines = (entry.vi || "").split("\n").filter((l) => l.trim());
      const firstMeaning = meaningLines[0] || "";

      // Accordion sections
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
                .map((m, idx) => `
                  <div class="dict-meaning-item">
                    <span class="dict-meaning-bullet">${idx + 1}.</span> ${this.escapeHtml(m)}
                  </div>
                `)
                .join("")}
            </div>
          </div>
        `);
      }

      // 2. Examples
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

      // 3. Grammar: 6 Cases (Declension)
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

      // 4. Grammar: Verb Conjugation
      if (extra.forms && extra.forms.v) {
        const verbHtml = this.buildVerbTable(extra.forms.v);
        if (verbHtml) {
          sectionsHtml.push(`
            <div class="dict-sec open" data-sec="verb">
              <button class="dict-sec-head" type="button">
                <span class="dict-sec-head-left">
                  <span class="dict-sec-icon">⏱️</span> Bảng chia động từ (Спряжение)
                </span>
                <span class="dict-sec-arrow">▾</span>
              </button>
              <div class="dict-sec-body">
                ${verbHtml}
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
          <button class="dict-back-btn" id="dictBackBtn">← Quay lại danh sách</button>

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

      // Bookmark button
      const bookmarkBtn = this.resultsEl.querySelector("#dictDetailBookmark");
      if (bookmarkBtn) {
        bookmarkBtn.addEventListener("click", () => {
          this.toggleBookmark(entry, data.dir);
          const hasBookmarked = this.bookmarks.some((b) => b.word === mainWord);
          bookmarkBtn.classList.toggle("active", hasBookmarked);
          bookmarkBtn.innerHTML = hasBookmarked ? "❤️" : "🤍";
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

      // Auto play audio softly once
      setTimeout(() => {
        this.playAudio(mainWord, audioUrl, 1, true);
      }, 200);
    }

    buildDeclensionTable(forms) {
      if (!forms) return "";
      let rows = [];
      let headers = [];

      if (forms.n) {
        // Noun: singular & plural
        headers = ["Cách", "Số ít (Ед.ч)", "Số nhiều (Мн.ч)"];
        rows = CASE_ROWS.map((c, i) => [
          `<span class="case-title">${c.ru}</span><span class="case-sub">${c.vn}</span>`,
          forms.n.s ? forms.n.s[i] || "—" : "—",
          forms.n.p ? forms.n.p[i] || "—" : "—",
        ]);
      } else if (forms.a) {
        // Adjective: masculine, feminine, neuter, plural
        headers = ["Cách", "Giống đực", "Giống cái", "Giống trung", "Số nhiều"];
        rows = CASE_ROWS.map((c, i) => [
          `<span class="case-title">${c.ru}</span><span class="case-sub">${c.vn}</span>`,
          forms.a.m ? forms.a.m[i] || "—" : "—",
          forms.a.f ? forms.a.f[i] || "—" : "—",
          forms.a.n ? forms.a.n[i] || "—" : "—",
          forms.a.p ? forms.a.p[i] || "—" : "—",
        ]);
      } else if (forms.pn) {
        // Pronoun
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

    buildVerbTable(v) {
      if (!v) return "";
      const persons = ["я (tôi)", "ты (bạn)", "он/она (anh ấy/cô ấy)", "мы (chúng tôi)", "вы (các bạn)", "они (họ)"];
      const isPerf = v.perf === 1;
      const tenseLabel = isPerf ? "Tương lai đơn (Будущее время)" : "Hiện tại (Настоящее время)";

      let html = `
        <div style="margin-bottom: 8px;">
          <span style="font-size: 0.76rem; font-weight: 700; background: ${isPerf ? "#fef3c7; color: #92400e;" : "#dbeafe; color: #1e40af;"} padding: 2px 8px; border-radius: 6px;">
            ${isPerf ? "★ Hoàn thành thể (СВ)" : "● Chưa hoàn thành thể (НСВ)"}
          </span>
        </div>
      `;

      if (v.pres && v.pres.length >= 6) {
        html += `
          <div class="dict-table-wrap" style="margin-bottom: 12px;">
            <table class="dict-table">
              <thead><tr><th>Ngôi</th><th>${tenseLabel}</th></tr></thead>
              <tbody>
                ${persons
                  .map(
                    (p, i) => `
                  <tr>
                    <td><b>${p}</b></td>
                    <td class="ru">${this.escapeHtml(v.pres[i] || "—")}</td>
                  </tr>
                `
                  )
                  .join("")}
              </tbody>
            </table>
          </div>
        `;
      }

      if (v.past) {
        html += `
          <div class="dict-table-wrap" style="margin-bottom: 12px;">
            <table class="dict-table">
              <thead><tr><th>Ngôi</th><th>Quá khứ (Прошедшее время)</th></tr></thead>
              <tbody>
                <tr><td><b>он (nam)</b></td><td class="ru">${this.escapeHtml(v.past.m || "—")}</td></tr>
                <tr><td><b>она (nữ)</b></td><td class="ru">${this.escapeHtml(v.past.f || "—")}</td></tr>
                <tr><td><b>оно (trung)</b></td><td class="ru">${this.escapeHtml(v.past.n || "—")}</td></tr>
                <tr><td><b>они (số nhiều)</b></td><td class="ru">${this.escapeHtml(v.past.p || "—")}</td></tr>
              </tbody>
            </table>
          </div>
        `;
      }

      if (v.imp && (v.imp.s || v.imp.p)) {
        html += `
          <div class="dict-table-wrap">
            <table class="dict-table">
              <thead><tr><th>Ngôi</th><th>Mệnh lệnh (Повелительное наклонение)</th></tr></thead>
              <tbody>
                <tr><td><b>ты (bạn)</b></td><td class="ru">${this.escapeHtml(v.imp.s || "—")}</td></tr>
                <tr><td><b>вы (các bạn/lịch sự)</b></td><td class="ru">${this.escapeHtml(v.imp.p || "—")}</td></tr>
              </tbody>
            </table>
          </div>
        `;
      }

      return html;
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
      if (!("speechSynthesis" in window)) return;
      speechSynthesis.cancel();
      const clean = this.stripStress(text);
      const utter = new SpeechSynthesisUtterance(clean);
      utter.rate = rate;
      utter.lang = "ru-RU";

      // Select natural Russian voice if available
      const voices = speechSynthesis.getVoices();
      const ruVoice = voices.find((v) => v.lang.startsWith("ru"));
      if (ruVoice) utter.voice = ruVoice;

      speechSynthesis.speak(utter);
    }

    toggleBookmark(entry, dir) {
      const isRuVi = dir === "ru-vi";
      const word = isRuVi ? entry.ru : entry.vi;
      const meaning = isRuVi ? entry.vi.split("\n")[0] : entry.ru;

      const idx = this.bookmarks.findIndex((b) => b.word === word);
      if (idx !== -1) {
        this.bookmarks.splice(idx, 1);
      } else {
        this.bookmarks.unshift({
          word,
          meaning,
          dir,
          date: Date.now(),
        });
        if (this.bookmarks.length > 100) this.bookmarks.pop();
      }
      this.saveStorage(STORAGE_BOOKMARKS, this.bookmarks);
    }

    addToHistory(word, dir) {
      const clean = word.trim();
      if (!clean) return;
      this.history = this.history.filter((h) => h.word.toLowerCase() !== clean.toLowerCase());
      this.history.unshift({ word: clean, dir, date: Date.now() });
      if (this.history.length > 50) this.history.pop();
      this.saveStorage(STORAGE_HISTORY, this.history);
    }

    renderBookmarks() {
      const pane = this.panelEl.querySelector("#dictTabBookmarks");
      if (this.bookmarks.length === 0) {
        pane.innerHTML = `
          <div class="dict-empty-state">
            <span class="dict-empty-icon">⭐</span>
            <p class="dict-empty-title">Chưa có từ nào được lưu</p>
            <p class="dict-empty-sub">Nhấn biểu tượng trái tim (❤️) khi xem chi tiết từ để lưu vào sổ tay ôn tập.</p>
          </div>
        `;
        return;
      }

      pane.innerHTML = `
        <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:10px;">
          <span style="font-weight:700; font-size:0.88rem; color:#1e293b;">Từ đã lưu (${this.bookmarks.length})</span>
          <button class="dict-nav-tab" id="dictClearBookmarksBtn" style="color:#ef4444; font-size:0.75rem;">Xóa tất cả</button>
        </div>
        <div class="dict-results-list">
          ${this.bookmarks
            .map(
              (b) => `
            <div class="dict-result-card" data-word="${this.escapeHtml(b.word)}" data-dir="${this.escapeHtml(b.dir)}">
              <div class="dict-card-head">
                <span class="dict-card-word">${this.escapeHtml(b.word)}</span>
              </div>
              <div class="dict-card-meaning">${this.escapeHtml(b.meaning)}</div>
            </div>
          `
            )
            .join("")}
        </div>
      `;

      pane.querySelectorAll(".dict-result-card").forEach((card) => {
        card.addEventListener("click", () => {
          this.switchTab("search");
          this.searchInput.value = card.dataset.word;
          this.clearBtn.style.display = "grid";
          this.openWordDetail(card.dataset.word, card.dataset.dir);
        });
      });

      const clearAll = pane.querySelector("#dictClearBookmarksBtn");
      if (clearAll) {
        clearAll.addEventListener("click", () => {
          if (confirm("Bạn có chắc muốn xóa tất cả từ đã lưu?")) {
            this.bookmarks = [];
            this.saveStorage(STORAGE_BOOKMARKS, this.bookmarks);
            this.renderBookmarks();
          }
        });
      }
    }

    renderHistory() {
      const pane = this.panelEl.querySelector("#dictTabHistory");
      if (this.history.length === 0) {
        pane.innerHTML = `
          <div class="dict-empty-state">
            <span class="dict-empty-icon">🕒</span>
            <p class="dict-empty-title">Lịch sử tra cứu trống</p>
            <p class="dict-empty-sub">Các từ bạn tra cứu sẽ xuất hiện tại đây để tiện xem lại.</p>
          </div>
        `;
        return;
      }

      pane.innerHTML = `
        <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:10px;">
          <span style="font-weight:700; font-size:0.88rem; color:#1e293b;">Tra cứu gần đây (${this.history.length})</span>
          <button class="dict-nav-tab" id="dictClearHistoryBtn" style="color:#ef4444; font-size:0.75rem;">Xóa lịch sử</button>
        </div>
        <div class="dict-results-list">
          ${this.history
            .map(
              (h) => `
            <div class="dict-result-card" data-word="${this.escapeHtml(h.word)}" data-dir="${this.escapeHtml(h.dir)}">
              <div class="dict-card-head">
                <span class="dict-card-word">${this.escapeHtml(h.word)}</span>
              </div>
            </div>
          `
            )
            .join("")}
        </div>
      `;

      pane.querySelectorAll(".dict-result-card").forEach((card) => {
        card.addEventListener("click", () => {
          this.switchTab("search");
          this.searchInput.value = card.dataset.word;
          this.clearBtn.style.display = "grid";
          this.openWordDetail(card.dataset.word, card.dataset.dir);
        });
      });

      const clearHist = pane.querySelector("#dictClearHistoryBtn");
      if (clearHist) {
        clearHist.addEventListener("click", () => {
          this.history = [];
          this.saveStorage(STORAGE_HISTORY, this.history);
          this.renderHistory();
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

