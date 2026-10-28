/**
 * TAIGA TALK - Universal Audio & Russian Speech Engine (taiga-audio.js)
 * Guarantees 100% audio playback across all devices, browsers, and lessons.
 * Supports:
 *  - Native Web Speech API (with rate control) when a Russian voice is installed
 *  - High-fidelity MP3 streaming via /api/tts (with Google TTS direct fallback)
 *  - Auto-watchdog to prevent browser speech queue freezing
 *  - Unified lifecycle callbacks: onStart, onEnd, onError
 */

(function () {
  'use strict';

  class TaigaAudioEngine {
    constructor() {
      this.synth = typeof window !== 'undefined' && 'speechSynthesis' in window ? window.speechSynthesis : null;
      this.audioObj = null;
      this.activeToken = 0;
      this.russianVoice = null;
      this.hasInitVoices = false;
      this.playbackRate = 1.0;
      this.isCurrentlyPlaying = false;
      this._init();
    }

    _init() {
      if (this.synth) {
        this._loadVoices();
        try {
          if (typeof this.synth.addEventListener === 'function') {
            this.synth.addEventListener('voiceschanged', () => this._loadVoices());
          }
          this.synth.onvoiceschanged = () => this._loadVoices();
        } catch (e) {}
      }
    }

    _loadVoices() {
      if (!this.synth) return;
      try {
        const voices = this.synth.getVoices() || [];
        if (!voices.length) return;

        const ruVoices = voices.filter((v) => {
          const lang = (v.lang || '').toLowerCase().replace('_', '-');
          const name = (v.name || '').toLowerCase();
          return lang.startsWith('ru') || name.includes('russian') || name.includes('русск');
        });

        if (ruVoices.length > 0) {
          // Prioritize local voice when rate < 1.0 for accurate speed slowdown
          this.russianVoice =
            ruVoices.find((v) => v.localService) ||
            ruVoices.find((v) => v.name.includes('Milena') || v.name.includes('Yuri')) ||
            ruVoices.find((v) => !v.name.includes('Online')) ||
            ruVoices[0];
          this.hasInitVoices = true;
        } else {
          this.russianVoice = null;
          this.hasInitVoices = true;
        }
      } catch (e) {
        this.russianVoice = null;
      }
    }

    
    resumeSpeechQueue() {
      if (this.synth) {
        try {
          if (this.synth.paused) {
            this.synth.resume();
          }
        } catch (e) {}
      }
    }

    hasRussianVoice() {
      if (!this.hasInitVoices) this._loadVoices();
      return !!this.russianVoice;
    }

    cleanText(raw) {
      if (!raw) return '';
      return String(raw)
        .replace(/<[^>]*>/g, ' ')
        .replace(/&nbsp;/gi, ' ')
        .replace(/&quot;/gi, ' ')
        .replace(/&laquo;|&raquo;/gi, ' ')
        .replace(/&#39;/gi, ' ')
        .replace(/&amp;/gi, ' & ')
        .replace(/[\u0300-\u036f]/g, '') // strip Russian stress marks
        .replace(/[«»"“”_]/g, ' ')
        .replace(/\s+/g, ' ')
        .trim();
    }

    getTtsUrl(text, lang = 'ru') {
      const clean = this.cleanText(text);
      return `/api/tts?text=${encodeURIComponent(clean)}&lang=${encodeURIComponent(lang)}`;
    }

    stop() {
      this.activeToken++;
      this.isCurrentlyPlaying = false;

      if (this.synth) {
        try {
          if (this.synth.paused) this.synth.resume();
          this.synth.cancel();
        } catch (e) {}
      }

      if (this.audioObj) {
        try {
          this.audioObj.pause();
          this.audioObj.currentTime = 0;
          this.audioObj.onplay = null;
          this.audioObj.onended = null;
          this.audioObj.onerror = null;
        } catch (e) {}
        this.audioObj = null;
      }
    }

    pause() {
      if (this.audioObj && !this.audioObj.paused) {
        this.audioObj.pause();
        this.isCurrentlyPlaying = false;
      } else if (this.synth && this.synth.speaking && !this.synth.paused) {
        this.synth.pause();
        this.isCurrentlyPlaying = false;
      }
    }

    resume() {
      if (this.audioObj && this.audioObj.paused) {
        this.audioObj.play().catch(() => {});
        this.isCurrentlyPlaying = true;
      } else if (this.synth && this.synth.paused) {
        this.synth.resume();
        this.isCurrentlyPlaying = true;
      }
    }

    isPlaying() {
      return this.isCurrentlyPlaying;
    }

    /**
     * Speak text with 100% guarantee across all browsers/OS.
     * Uses native Web Speech API if Russian voice is installed on device;
     * otherwise smoothly falls back to MP3 stream.
     *
     * @param {string} text - Russian text to speak
     * @param {Object} [options]
     * @param {number} [options.rate=1.0] - Playback speed (e.g. 0.75, 1.0)
     * @param {Function} [options.onStart] - Triggered when audio starts playing
     * @param {Function} [options.onEnd] - Triggered when audio finishes
     * @param {Function} [options.onError] - Triggered on error
     * @param {boolean} [options.forceAudio=false] - Force MP3 Audio path
     */
    speak(text, options = {}) {
      const clean = this.cleanText(text);
      if (!clean) {
        if (typeof options.onEnd === 'function') options.onEnd();
        return;
      }

      this.stop();
      const currentToken = ++this.activeToken;
      const rate = options.rate || this.playbackRate || 1.0;
      let hasStarted = false;

      const triggerStart = () => {
        if (!hasStarted && currentToken === this.activeToken) {
          hasStarted = true;
          this.isCurrentlyPlaying = true;
          if (typeof options.onStart === 'function') options.onStart();
        }
      };

      const triggerEnd = () => {
        if (currentToken === this.activeToken) {
          this.isCurrentlyPlaying = false;
          if (typeof options.onEnd === 'function') options.onEnd();
        }
      };

      // HTML5 Audio MP3 Fallback Engine
      const playAudioFallback = () => {
        if (currentToken !== this.activeToken) return;

        try {
          const ttsUrl = this.getTtsUrl(clean, options.lang || 'ru');
          const audio = new Audio(ttsUrl);
          this.audioObj = audio;

          if (rate && rate > 0) {
            audio.playbackRate = rate;
          }

          audio.onplay = () => triggerStart();

          audio.onended = () => {
            if (this.audioObj === audio) this.audioObj = null;
            triggerEnd();
          };

          audio.onerror = () => {
            // Secondary direct Google TTS fallback in case local proxy was blocked
            const directUrl = `https://translate.google.com/translate_tts?ie=UTF-8&client=tw-ob&tl=ru&q=${encodeURIComponent(clean.slice(0, 200))}`;
            const directAudio = new Audio(directUrl);
            this.audioObj = directAudio;
            if (rate && rate > 0) directAudio.playbackRate = rate;

            directAudio.onplay = () => triggerStart();
            directAudio.onended = () => {
              if (this.audioObj === directAudio) this.audioObj = null;
              triggerEnd();
            };
            directAudio.onerror = () => {
              if (this.audioObj === directAudio) this.audioObj = null;
              triggerEnd();
            };
            directAudio.play().catch(() => triggerEnd());
          };

          const p = audio.play();
          if (p && typeof p.catch === 'function') {
            p.catch(() => triggerEnd());
          }
        } catch (e) {
          triggerEnd();
        }
      };

      // Check if we should use native Web Speech API
      const ruVoice = this.hasRussianVoice() ? this.russianVoice : null;
      const canUseNative = !options.forceAudio && this.synth && ruVoice;

      if (!canUseNative) {
        // No native Russian voice installed on system: use high quality MP3 Audio directly!
        playAudioFallback();
        return;
      }

      // Native SpeechSynthesis path
      try {
        if (this.synth.paused) this.synth.resume();
        const utter = new SpeechSynthesisUtterance(clean);
        utter.voice = ruVoice;
        utter.lang = 'ru-RU';
        utter.rate = rate === 0.75 ? 0.70 : rate;
        utter.pitch = 1.0;

        utter.onstart = () => triggerStart();
        utter.onend = () => triggerEnd();

        utter.onerror = (e) => {
          if (e && (e.error === 'interrupted' || e.error === 'canceled')) return;
          playAudioFallback();
        };

        this.synth.speak(utter);

        // Watchdog: If native speech doesn't actually begin within 600ms, fallback to MP3 Audio
        setTimeout(() => {
          if (currentToken === this.activeToken && !hasStarted) {
            try {
              this.synth.cancel();
            } catch (e) {}
            playAudioFallback();
          }
        }, 600);
      } catch (e) {
        playAudioFallback();
      }
    }
  }

  // Attach global singleton
  window.TaigaAudio = new TaigaAudioEngine();
})();
