/**
 * TAIGA FLASH — FORUMFLASH ENGINE & GAMIFICATION CONTROLLER
 * Inspired by ForumFlash (https://www.forumflashstudy.com/)
 */

(function () {
  "use strict";

  const STORAGE_XP = "taiga_xp";
  const STORAGE_STREAK = "taiga_streak";
  const STORAGE_STAGE = "taiga_stage";
  const STORAGE_LAST_VISIT = "taiga_last_visit";

  const STAGE_THRESHOLDS = [0, 150, 400, 750, 1200, 1800]; // XP thresholds for stages 1 to 6
  const STAGE_NAMES = ["Hạt giống", "Nảy mầm", "Phát triển thân", "Phân cành", "Ra lá", "Kết trái"];
  const STAGE_NAMES_RU = ["Семечко", "Росток", "Стебель", "Ветви", "Листья", "Плоды"];

  class ForumFlashEngine {
    constructor() {
      this.xp = this.loadNumber(STORAGE_XP, 120);
      this.streak = this.loadStreakNumber(STORAGE_STREAK, 99);
      this.currentStage = this.calculateStage(this.xp);
      this.activeFlowStep = "flashcard"; // 'flashcard' | 'quiz'

      // Flashcard state
      this.flashcards = [];
      this.cardIndex = 0;
      this.cardFlipped = false;
      this.masteredWords = new Set();

      this.checkDailyStreak();
    }

    loadNumber(key, fallback) {
      try {
        const val = localStorage.getItem(key);
        return val !== null ? parseInt(val, 10) : fallback;
      } catch (e) {
        return fallback;
      }
    }

    saveNumber(key, val) {
      try {
        localStorage.setItem(key, String(val));
      } catch (e) {}
    }

    checkDailyStreak() {
      try {
        const today = new Date().toDateString();
        const lastVisit = localStorage.getItem(STORAGE_LAST_VISIT);
        if (!lastVisit) {
          localStorage.setItem(STORAGE_LAST_VISIT, today);
        } else if (lastVisit !== today) {
          const yesterday = new Date(Date.now() - 86400000).toDateString();
          if (lastVisit === yesterday) {
            this.streak += 1;
          } else {
            this.streak = 1;
          }
          this.saveNumber(STORAGE_STREAK, this.streak);
          localStorage.setItem(STORAGE_LAST_VISIT, today);
        }
      } catch (e) {}
    }

    calculateStage(xp) {
      for (let i = STAGE_THRESHOLDS.length - 1; i >= 0; i--) {
        if (xp >= STAGE_THRESHOLDS[i]) return i + 1;
      }
      return 1;
    }

    addXP(amount, triggerModal = false) {
      this.xp += amount;
      this.saveNumber(STORAGE_XP, this.xp);

      const oldStage = this.currentStage;
      this.currentStage = this.calculateStage(this.xp);
      this.saveNumber(STORAGE_STAGE, this.currentStage);

      this.updateHeaderStats();
      this.renderGrowthSection();

      if (triggerModal) {
        this.showCelebrationModal(amount, this.currentStage > oldStage);
      }
    }

    updateHeaderStats() {
      const xpEl = document.getElementById("ffHeaderXp");
      const streakEl = document.getElementById("ffHeaderStreak");
      if (xpEl) xpEl.textContent = `${this.xp} XP`;
      if (streakEl) streakEl.textContent = `${this.streak}`;

      // Update Hero Game Card XP
      const cardProgress = document.getElementById("heroCardProgressText");
      if (cardProgress) {
        cardProgress.textContent = `${this.xp} XP • Cấp độ ${this.currentStage}`;
      }
    }

    renderGrowthSection() {
      const stageIdx = Math.min(this.currentStage - 1, 5);
      const stageName = STAGE_NAMES[stageIdx];
      const stageNameRu = STAGE_NAMES_RU[stageIdx];

      const titleEl = document.getElementById("learningTreeTitle");
      if (titleEl) {
        titleEl.textContent = `${stageName} (${stageNameRu})`;
      }

      // Next stage progress
      const nextXP = STAGE_THRESHOLDS[Math.min(stageIdx + 1, 5)] || STAGE_THRESHOLDS[5];
      const prevXP = STAGE_THRESHOLDS[stageIdx];
      const range = Math.max(1, nextXP - prevXP);
      const currProgress = Math.min(100, Math.max(0, Math.round(((this.xp - prevXP) / range) * 100)));

      const goalTextEl = document.getElementById("treeGoalText");
      if (goalTextEl) {
        if (stageIdx < 5) {
          const remaining = Math.max(0, nextXP - this.xp);
          goalTextEl.innerHTML = `Còn <strong>${remaining} XP</strong> để đạt ${STAGE_NAMES[stageIdx + 1]}`;
        } else {
          goalTextEl.innerHTML = `Đã đạt <strong>giai đoạn tối đa</strong> (${STAGE_NAMES[5]})!`;
        }
      }

      const percentEl = document.getElementById("treePercentText");
      if (percentEl) percentEl.textContent = `${currProgress}%`;

      const fillEl = document.getElementById("treeProgressFill");
      if (fillEl) fillEl.style.width = `${currProgress}%`;

      // Update 6-stage track items
      document.querySelectorAll(".tree-stage-track li").forEach((li, idx) => {
        li.classList.toggle("active", idx === stageIdx);
        li.classList.toggle("done", idx < stageIdx);
      });

      // Render Tree SVG according to stage
      const svgContainer = document.getElementById("learningTreeContainer");
      if (svgContainer) {
        svgContainer.innerHTML = this.getTreeSVG(this.currentStage);
      }
    }

    getTreeSVG(stage) {
      let stageArt = "";
      if (stage === 1) {
        // Hạt giống (Seed)
        stageArt = `
          <g class="tree-seed" filter="url(#tree-shadow)">
            <path d="M168 304c23-13 60-14 84 1-23 12-60 13-84-1Z" fill="#d5ae83" opacity=".48"></path>
            <path d="M190 294c5-19 19-30 38-29-4 18-17 30-38 29Z" fill="url(#tree-seed-shell)"></path>
            <path d="M194 290c8-9 17-16 29-20" fill="none" stroke="#d9b184" stroke-width="3" stroke-linecap="round" opacity=".75"></path>
            <path d="M215 267c0 6 4 11 10 14" fill="none" stroke="#613d29" stroke-width="2.5" stroke-linecap="round"></path>
          </g>
        `;
      } else if (stage === 2) {
        // Nảy mầm (Sprout)
        stageArt = `
          <path d="M208 308 Q208 270 210 248" fill="none" stroke="url(#tree-leaves)" stroke-width="6" stroke-linecap="round"/>
          <path d="M210 248 Q185 240 180 220 Q200 222 210 246" fill="url(#tree-leaves)" filter="url(#tree-canopy-shadow)"/>
          <path d="M210 248 Q235 238 245 218 Q235 230 210 246" fill="url(#tree-leaves-light)" filter="url(#tree-canopy-shadow)"/>
        `;
      } else if (stage === 3) {
        // Phát triển thân (Stem)
        stageArt = `
          <path d="M200 310 Q206 240 208 190 Q212 240 220 310 Z" fill="url(#tree-trunk)"/>
          <path d="M208 210 Q170 195 160 170 Q185 178 208 206" fill="url(#tree-leaves)" filter="url(#tree-canopy-shadow)"/>
          <path d="M208 190 Q240 175 255 150 Q235 165 208 188" fill="url(#tree-leaves-light)" filter="url(#tree-canopy-shadow)"/>
        `;
      } else if (stage === 4) {
        // Phân cành (Branches)
        stageArt = `
          <path d="M195 312 Q205 240 208 170 Q215 240 225 312 Z" fill="url(#tree-trunk)"/>
          <path d="M208 210 Q160 180 135 150 Q160 160 206 200" fill="url(#tree-branch)" stroke-width="4"/>
          <path d="M208 190 Q255 165 285 135 Q255 150 208 182" fill="url(#tree-branch)" stroke-width="4"/>
          <circle cx="135" cy="150" r="28" fill="url(#tree-leaves)" filter="url(#tree-canopy-shadow)"/>
          <circle cx="285" cy="135" r="32" fill="url(#tree-leaves-light)" filter="url(#tree-canopy-shadow)"/>
          <circle cx="210" cy="140" r="36" fill="url(#tree-leaves)" filter="url(#tree-canopy-shadow)"/>
        `;
      } else if (stage === 5) {
        // Ra lá (Lush Leaves)
        stageArt = `
          <path d="M192 312 Q206 230 208 160 Q214 230 228 312 Z" fill="url(#tree-trunk)"/>
          <circle cx="210" cy="130" r="72" fill="url(#tree-leaves)" filter="url(#tree-canopy-shadow)"/>
          <circle cx="155" cy="145" r="54" fill="url(#tree-leaves-light)" filter="url(#tree-canopy-shadow)"/>
          <circle cx="265" cy="140" r="58" fill="url(#tree-leaves)" filter="url(#tree-canopy-shadow)"/>
          <circle cx="210" cy="90" r="48" fill="url(#tree-leaves-light)"/>
        `;
      } else {
        // Kết trái (Fruiting)
        stageArt = `
          <path d="M192 312 Q206 230 208 160 Q214 230 228 312 Z" fill="url(#tree-trunk)"/>
          <circle cx="210" cy="130" r="76" fill="url(#tree-leaves)" filter="url(#tree-canopy-shadow)"/>
          <circle cx="150" cy="140" r="58" fill="url(#tree-leaves-light)" filter="url(#tree-canopy-shadow)"/>
          <circle cx="270" cy="135" r="62" fill="url(#tree-leaves)" filter="url(#tree-canopy-shadow)"/>
          <circle cx="210" cy="85" r="52" fill="url(#tree-leaves-light)"/>
          <!-- Fruits -->
          <circle cx="170" cy="120" r="9" fill="#ffb703" stroke="#fff" stroke-width="1.5"/>
          <circle cx="240" cy="110" r="9" fill="#fb8500" stroke="#fff" stroke-width="1.5"/>
          <circle cx="205" cy="155" r="9" fill="#ffb703" stroke="#fff" stroke-width="1.5"/>
          <circle cx="280" cy="145" r="8" fill="#fb8500" stroke="#fff" stroke-width="1.5"/>
          <circle cx="140" cy="155" r="8" fill="#ffb703" stroke="#fff" stroke-width="1.5"/>
        `;
      }

      return `
        <svg class="learning-tree stage-${stage}" viewBox="0 0 420 350" role="img" aria-label="Cây tiến độ đang ở giai đoạn ${STAGE_NAMES[stage - 1]}">
          <defs>
            <linearGradient id="tree-ground" x1="0" x2="1">
              <stop offset="0" stop-color="#dff4f1"></stop>
              <stop offset=".55" stop-color="#edf7f4"></stop>
              <stop offset="1" stop-color="#f8f3e8"></stop>
            </linearGradient>
            <linearGradient id="tree-trunk" x1="0" y1=".5" x2="1" y2=".5">
              <stop offset="0" stop-color="#70472f"></stop>
              <stop offset=".48" stop-color="#a8734c"></stop>
              <stop offset=".75" stop-color="#ba865d"></stop>
              <stop offset="1" stop-color="#7f5338"></stop>
            </linearGradient>
            <linearGradient id="tree-branch" x1="0" y1="1" x2="1" y2="0">
              <stop offset="0" stop-color="#795036"></stop>
              <stop offset="1" stop-color="#ae7952"></stop>
            </linearGradient>
            <linearGradient id="tree-leaves" x1="0" y1="1" x2="1" y2="0">
              <stop offset="0" stop-color="#004646"></stop>
              <stop offset=".58" stop-color="#08736d"></stop>
              <stop offset="1" stop-color="#27c4b5"></stop>
            </linearGradient>
            <linearGradient id="tree-leaves-light" x1="0" y1="1" x2="1" y2="0">
              <stop offset="0" stop-color="#15978d"></stop>
              <stop offset="1" stop-color="#62d9cc"></stop>
            </linearGradient>
            <linearGradient id="tree-seed-shell" x1="0" y1="1" x2="1" y2="0">
              <stop offset="0" stop-color="#71472d"></stop>
              <stop offset="1" stop-color="#bf8b5d"></stop>
            </linearGradient>
            <radialGradient id="tree-sun">
              <stop offset="0" stop-color="#ffcf55" stop-opacity=".35"></stop>
              <stop offset="1" stop-color="#ffcf55" stop-opacity="0"></stop>
            </radialGradient>
            <filter id="tree-shadow" x="-30%" y="-30%" width="160%" height="170%">
              <feDropShadow dx="0" dy="8" stdDeviation="7" flood-color="#004646" flood-opacity=".14"></feDropShadow>
            </filter>
            <filter id="tree-canopy-shadow" x="-25%" y="-25%" width="150%" height="160%">
              <feDropShadow dx="0" dy="7" stdDeviation="6" flood-color="#004646" flood-opacity=".16"></feDropShadow>
            </filter>
          </defs>

          <!-- Sun & Sky -->
          <circle cx="326" cy="68" r="60" fill="url(#tree-sun)"></circle>
          <path d="M47 236c14-13 30-13 44 0M329 230c14-12 29-12 43 0" fill="none" stroke="#27c4b5" stroke-width="3" stroke-linecap="round" opacity=".28"></path>

          <!-- Ground -->
          <ellipse cx="210" cy="311" rx="160" ry="25" fill="url(#tree-ground)"></ellipse>
          <path d="M65 312c42-12 90-17 145-17 58 0 108 7 148 20" fill="none" stroke="#a9d9d3" stroke-width="3" stroke-linecap="round"></path>
          <path d="M97 321c17-4 35-4 53-1M280 321c14-2 28-1 41 3" fill="none" stroke="#66c8bd" stroke-width="2" stroke-linecap="round" opacity=".55"></path>

          <!-- Stage-specific Art -->
          ${stageArt}

          <!-- Badge Seal -->
          <g class="tree-stage-seal">
            <rect x="305" y="276" width="94" height="34" rx="17" fill="#fff" stroke="#d9e3e4"></rect>
            <circle cx="323" cy="293" r="10" fill="#e5f7f4"></circle>
            <text x="323" y="297" text-anchor="middle" fill="#004646" font-size="11" font-weight="800">${stage}</text>
            <text x="362" y="297" text-anchor="middle" fill="#004646" font-size="11" font-weight="800">GIAI ĐOẠN</text>
          </g>
        </svg>
      `;
    }

    // ======================================================================
    // FLASHCARD COMPONENT (STAGE 1 OF LESSON)
    // ======================================================================
    initFlashcardsFromLesson(lessonData) {
      this.flashcards = [];
      this.cardIndex = 0;
      this.cardFlipped = false;

      if (!lessonData) return;

      if (Array.isArray(lessonData.vocabCategories)) {
        lessonData.vocabCategories.forEach((cat) => {
          if (Array.isArray(cat.rows)) {
            cat.rows.forEach((r) => {
              this.flashcards.push({
                word: r.word,
                gram: r.gram || cat.title || "Từ vựng",
                meaning: r.meaning || "",
                example: r.example || "",
              });
            });
          }
        });
      }

      if (this.flashcards.length === 0 && Array.isArray(lessonData.vocabulary)) {
        lessonData.vocabulary.forEach((v) => {
          this.flashcards.push({
            word: v.ru || v.word,
            gram: v.pos || "Từ vựng",
            meaning: v.vi || v.meaning || "",
            example: v.example || "",
          });
        });
      }

      // Fallback if no vocab exists in lesson
      if (this.flashcards.length === 0) {
        this.flashcards.push({
          word: "Здравствуйте",
          gram: "lời chào lịch sự",
          meaning: "Xin chào (trang trọng)",
          example: "Здравствуйте, меня зовут Антон.",
        });
        this.flashcards.push({
          word: "Спасибо",
          gram: "từ cảm ơn",
          meaning: "Cảm ơn bạn",
          example: "Большое спасибо за урок!",
        });
      }

      this.renderFlashcard();
    }

    renderFlashcard() {
      if (this.flashcards.length === 0) return;
      const card = this.flashcards[this.cardIndex];
      const scene = document.getElementById("ffFlashcardScene");
      if (scene) scene.classList.remove("flipped");
      this.cardFlipped = false;

      // Meta count
      const meta = document.getElementById("ffCardMeta");
      if (meta) meta.textContent = `Thẻ ${this.cardIndex + 1} / ${this.flashcards.length}`;

      const fill = document.getElementById("ffCardMeterFill");
      if (fill) fill.style.width = `${Math.round(((this.cardIndex + 1) / this.flashcards.length) * 100)}%`;

      // Front
      const frontWord = document.getElementById("ffFrontWord");
      if (frontWord) frontWord.textContent = card.word;
      const frontGram = document.getElementById("ffFrontGram");
      if (frontGram) frontGram.textContent = card.gram;

      const frontPhonetic = document.getElementById("ffFrontPhonetic");
      if (frontPhonetic) {
        if (card.ipa || card.phonetic) {
          frontPhonetic.style.display = "block";
          frontPhonetic.textContent = card.ipa ? card.ipa : `[${card.phonetic}]`;
        } else {
          frontPhonetic.style.display = "none";
          frontPhonetic.textContent = "";
        }
      }

      // Back
      const backMeaning = document.getElementById("ffBackMeaning");
      if (backMeaning) backMeaning.textContent = card.meaning;
      const backGram = document.getElementById("ffBackGram");
      if (backGram) backGram.textContent = card.gram;

      const backEx = document.getElementById("ffBackExample");
      if (backEx) {
        if (card.exampleRu) {
          backEx.style.display = "block";
          backEx.innerHTML = `<strong>Ví dụ:</strong> ${card.exampleRu}${card.exampleVi ? `<div style="font-size:0.85rem; opacity:0.88; margin-top:4px; font-weight:400; color:#e0f2fe;">↳ ${card.exampleVi}</div>` : ''}`;
        } else if (card.example) {
          backEx.style.display = "block";
          backEx.innerHTML = `<strong>Ví dụ thực tế:</strong> ${card.example}`;
        } else {
          backEx.style.display = "none";
        }
      }
    }

    flipFlashcard() {
      const scene = document.getElementById("ffFlashcardScene");
      if (!scene) return;
      this.cardFlipped = !this.cardFlipped;
      scene.classList.toggle("flipped", this.cardFlipped);
    }

    nextFlashcard(isMastered = false) {
      if (isMastered) {
        const cur = this.flashcards[this.cardIndex];
        if (cur) this.masteredWords.add(cur.word);
        this.addXP(5); // +5 XP per mastered word
      }

      if (this.cardIndex < this.flashcards.length - 1) {
        this.cardIndex++;
        this.renderFlashcard();
      } else {
        // Reached end of flashcards!
        this.switchFlowStep("quiz");
      }
    }

    prevFlashcard() {
      if (this.cardIndex > 0) {
        this.cardIndex--;
        this.renderFlashcard();
      }
    }

    speakCurrentCard() {
      const card = this.flashcards[this.cardIndex];
      if (!card) return;
      if (window.TaigaDictWidget && typeof window.TaigaDictWidget.speak === "function") {
        window.TaigaDictWidget.speak(card.word);
      } else if (window.speechSynthesis) {
        const u = new SpeechSynthesisUtterance(card.word.replace(/\u0301/g, ""));
        u.lang = "ru-RU";
        window.speechSynthesis.speak(u);
      }
    }

    // ======================================================================
    // FLOW STEP SWITCHER (CHẶNG 1: FLASHCARD ⇄ CHẶNG 2: QUIZ TEST)
    // ======================================================================
    switchFlowStep(step) {
      this.activeFlowStep = step;

      const btnStage1 = document.getElementById("btnStageFlashcard");
      const btnStage2 = document.getElementById("btnStageQuiz");
      if (btnStage1) btnStage1.classList.toggle("active", step === "flashcard");
      if (btnStage2) btnStage2.classList.toggle("active", step === "quiz");

      // Update Hero Game Card Route state
      const heroStep1 = document.getElementById("heroRouteStep1");
      const heroStep2 = document.getElementById("heroRouteStep2");
      if (heroStep1) heroStep1.classList.toggle("active", step === "flashcard");
      if (heroStep2) heroStep2.classList.toggle("active", step === "quiz");

      const flashArea = document.getElementById("flashcardArena");
      const quizArea = document.getElementById("activeLessonContainer");

      if (step === "flashcard") {
        if (flashArea) flashArea.style.display = "flex";
        if (quizArea) quizArea.style.display = "none";
      } else {
        if (flashArea) flashArea.style.display = "none";
        if (quizArea) quizArea.style.display = "grid";
      }
    }

    // ======================================================================
    // CELEBRATION MODAL
    // ======================================================================
    showCelebrationModal(xpAmount, stageLeveledUp = false) {
      const existing = document.getElementById("ffCelebrationModal");
      if (existing) existing.remove();

      const modal = document.createElement("div");
      modal.className = "ff-modal-overlay";
      modal.id = "ffCelebrationModal";

      const stageNote = stageLeveledUp
        ? `<div style="background:#e0f2fe; color:#0369a1; padding:8px 16px; border-radius:12px; font-weight:800; font-size:0.9rem;">
             🎉 Chúc mừng! Cây tri thức của bạn đã phát triển lên <strong>Cấp độ ${this.currentStage}: ${STAGE_NAMES[this.currentStage - 1]}</strong>!
           </div>`
        : "";

      modal.innerHTML = `
        <div class="ff-celebration-card">
          <div class="ff-cel-icon">🏆</div>
          <h3>Hoàn thành xuất sắc!</h3>
          <p>Bạn đã hoàn thành bài học và củng cố vững chắc kiến thức tiếng Nga.</p>
          <div class="ff-xp-reward-badge">
            <span>⚡ +${xpAmount} XP ĐÃ CỘNG VÀO CÂY TRI THỨC</span>
          </div>
          ${stageNote}
          <button class="btn-ff-primary" id="ffBtnCloseCelebration" style="width:100%; justify-content:center; margin-top:8px;">
            Tiếp tục lộ trình học →
          </button>
        </div>
      `;

      document.body.appendChild(modal);

      modal.querySelector("#ffBtnCloseCelebration")?.addEventListener("click", () => {
        modal.remove();
      });

      modal.addEventListener("click", (e) => {
        if (e.target === modal) modal.remove();
      });
    }

    // ======================================================================
    // 25-TOPIC VOCABULARY & FLASHCARDS SYSTEM
    // ======================================================================
    initVocabTopicsSystem() {
      // 1. Hook trigger card
      const campCard = document.getElementById("campCardVocab");
      if (campCard) {
        campCard.addEventListener("click", () => this.openVocabTopicsModal());
        campCard.addEventListener("keydown", (e) => {
          if (e.key === "Enter" || e.key === " ") {
            e.preventDefault();
            this.openVocabTopicsModal();
          }
        });
      }

      // 2. Hook modal close button & backdrop
      const modal = document.getElementById("vocabTopicsModal");
      const btnClose = document.getElementById("btnCloseVocabTopics");
      if (btnClose) {
        btnClose.addEventListener("click", () => this.closeVocabTopicsModal());
      }
      if (modal) {
        modal.addEventListener("click", (e) => {
          if (e.target === modal) this.closeVocabTopicsModal();
        });
      }

      // 3. Search input filtering
      const searchInput = document.getElementById("topicSearchInput");
      if (searchInput) {
        searchInput.addEventListener("input", (e) => {
          this.renderVocabTopicsList(e.target.value.trim());
        });
      }

      // 4. Keyboard ESC to close
      document.addEventListener("keydown", (e) => {
        if (e.key === "Escape") {
          const m = document.getElementById("vocabTopicsModal");
          if (m && m.classList.contains("active")) {
            this.closeVocabTopicsModal();
          }
        }
      });

      // 5. Initial render of topics list
      this.renderVocabTopicsList();

      // 6. Check URL parameters (?openVocab=true or ?topic=topic-1)
      try {
        const params = new URLSearchParams(window.location.search);
        if (params.get("openVocab") === "true") {
          setTimeout(() => this.openVocabTopicsModal(), 300);
        }
        const topicParam = params.get("topic");
        if (topicParam) {
          setTimeout(() => this.selectVocabTopic(topicParam), 400);
        }
      } catch (e) {}
    }

    openVocabTopicsModal() {
      const modal = document.getElementById("vocabTopicsModal");
      if (!modal) return;
      modal.classList.add("active", "open");
      modal.setAttribute("aria-hidden", "false");
      const searchInput = document.getElementById("topicSearchInput");
      if (searchInput) {
        searchInput.value = "";
        setTimeout(() => searchInput.focus(), 150);
      }
      this.renderVocabTopicsList();
    }

    closeVocabTopicsModal() {
      const modal = document.getElementById("vocabTopicsModal");
      if (!modal) return;
      modal.classList.remove("active", "open");
      modal.setAttribute("aria-hidden", "true");
    }

    renderVocabTopicsList(filterText = "") {
      const container = document.getElementById("topicsGridContainer");
      if (!container) return;
      const topics = window.VOCAB_25_TOPICS;
      if (!Array.isArray(topics) || topics.length === 0) {
        container.innerHTML = `<div style="text-align:center; padding:30px; color:#64748b;">Đang tải danh sách 25 chủ đề từ vựng...</div>`;
        return;
      }

      const q = filterText.toLowerCase();
      const filtered = topics.filter((t) => {
        if (!q) return true;
        return (
          t.titleVi.toLowerCase().includes(q) ||
          t.titleRu.toLowerCase().includes(q) ||
          t.desc.toLowerCase().includes(q) ||
          t.num.includes(q)
        );
      });

      if (filtered.length === 0) {
        container.innerHTML = `
          <div style="grid-column: 1 / -1; text-align:center; padding:40px 20px; color:#64748b;">
            <div style="font-size:2rem; margin-bottom:8px;">🔍</div>
            <p style="margin:0; font-weight:700;">Không tìm thấy chủ đề phù hợp với "${filterText}"</p>
            <p style="margin:4px 0 0 0; font-size:0.85rem;">Hãy thử từ khóa khác như "chào hỏi", "gia đình", "ẩm thực", "du học"...</p>
          </div>
        `;
        return;
      }

      container.innerHTML = filtered
        .map(
          (t) => `
        <div class="topic-card-item" data-topic-id="${t.id}" role="button" tabindex="0">
          <div class="topic-card-top">
            <div class="topic-card-icon-num">
              <span class="topic-card-icon">${t.icon}</span>
              <span class="topic-card-num">CHỦ ĐỀ ${t.num}</span>
            </div>
            <span class="topic-card-count">${t.count} thẻ</span>
          </div>
          <div>
            <h4 class="topic-card-title-vi">${t.titleVi}</h4>
            <div class="topic-card-title-ru">${t.titleRu}</div>
            <p class="topic-card-desc">${t.desc}</p>
          </div>
          <div class="topic-card-action">
            <span>Luyện ${t.count} flashcards</span>
            <span class="arrow">➔</span>
          </div>
        </div>
      `
        )
        .join("");

      // Bind click handlers to cards
      container.querySelectorAll(".topic-card-item").forEach((cardEl) => {
        const topicId = cardEl.getAttribute("data-topic-id");
        cardEl.addEventListener("click", () => this.selectVocabTopic(topicId));
        cardEl.addEventListener("keydown", (e) => {
          if (e.key === "Enter" || e.key === " ") {
            e.preventDefault();
            this.selectVocabTopic(topicId);
          }
        });
      });
    }

    selectVocabTopic(topicId) {
      if (!window.VOCAB_25_TOPICS) return;
      const topic = window.VOCAB_25_TOPICS.find((t) => t.id === topicId);
      if (!topic || !Array.isArray(topic.cards) || topic.cards.length === 0) return;

      this.closeVocabTopicsModal();

      // Load into flashcards engine
      this.flashcards = topic.cards.map((c) => ({
        word: c.word,
        gram: c.gram,
        ipa: c.ipa || "",
        phonetic: c.phonetic || "",
        meaning: c.meaning,
        exampleRu: c.exampleRu || "",
        exampleVi: c.exampleVi || "",
        example: c.exampleRu ? `${c.exampleRu} — ${c.exampleVi || ""}` : (c.meaning || "")
      }));
      this.cardIndex = 0;
      this.masteredWords.clear();

      // Update Dual-Stage Tab Label
      const routeTab = document.getElementById("btnStageFlashcard");
      if (routeTab) {
        routeTab.innerHTML = `
          <span class="stage-num-badge">01</span>
          <span>Chặng 1: ${topic.icon} ${topic.titleVi} (${this.flashcards.length} thẻ)</span>
        `;
      }

      // Switch to flashcard step and render
      this.switchFlowStep("flashcard");
      this.renderFlashcard();

      // Scroll to arena
      const arena = document.getElementById("arenaAnchor");
      if (arena) {
        arena.scrollIntoView({ behavior: "smooth" });
      }

      // Show toast
      this.showToast(`✨ Đã nạp ${this.flashcards.length} thẻ từ vựng: ${topic.titleVi}`);
    }

    showToast(message) {
      const existing = document.getElementById("ffGlobalToast");
      if (existing) existing.remove();
      const toast = document.createElement("div");
      toast.id = "ffGlobalToast";
      toast.style.cssText = `
        position: fixed;
        bottom: 24px;
        left: 50%;
        transform: translateX(-50%);
        background: #003838;
        color: #ffffff;
        padding: 12px 24px;
        border-radius: 9999px;
        font-weight: 700;
        font-size: 0.9rem;
        box-shadow: 0 10px 30px rgba(0, 56, 56, 0.35);
        z-index: 99999;
        display: flex;
        align-items: center;
        gap: 8px;
        border: 1px solid rgba(255, 255, 255, 0.2);
      `;
      toast.textContent = message;
      document.body.appendChild(toast);
      setTimeout(() => {
        toast.style.transition = "opacity 0.4s ease, transform 0.4s ease";
        toast.style.opacity = "0";
        toast.style.transform = "translateX(-50%) translateY(10px)";
        setTimeout(() => toast.remove(), 400);
      }, 3000);
    }
  }

  // Auto initialize on DOM ready
  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", () => {
      window.TaigaFlash = new ForumFlashEngine();
      window.TaigaFlash.updateHeaderStats();
      window.TaigaFlash.renderGrowthSection();
      window.TaigaFlash.initVocabTopicsSystem();
    });
  } else {
    window.TaigaFlash = new ForumFlashEngine();
    window.TaigaFlash.updateHeaderStats();
    window.TaigaFlash.renderGrowthSection();
    window.TaigaFlash.initVocabTopicsSystem();
  }
})();

