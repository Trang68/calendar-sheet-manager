/**
 * TAIGA TALK - Progressive Web App (PWA) & Instant QR Code Installer
 * Features:
 *  - Service Worker registration & lifecycle management
 *  - Dynamic QR Code generation for seamless 1-scan mobile installation
 *  - Corner Floating Widget for laptop/desktop users
 *  - Step-by-step interactive install modal (iPhone iOS / Android / Desktop)
 *  - Native 'beforeinstallprompt' 1-click install trigger for Chrome/Edge
 */

(function () {
  'use strict';

  // 1. Register Service Worker & Handle Version Upgrades
  let swRegistration = null;
  if ('serviceWorker' in navigator) {
    window.addEventListener('load', () => {
      navigator.serviceWorker
        .register('/sw.js')
        .then((reg) => {
          swRegistration = reg;
          // Check for worker updates
          reg.addEventListener('updatefound', () => {
            const newWorker = reg.installing;
            if (newWorker) {
              newWorker.addEventListener('statechange', () => {
                if (newWorker.state === 'installed' && navigator.serviceWorker.controller) {
                  console.log('[PWA] Phiên bản mới đã cài đặt -> Yêu cầu kích hoạt ngay.');
                  newWorker.postMessage({ type: 'SKIP_WAITING' });
                }
              });
            }
          });
        })
        .catch((err) => {
          console.warn('[PWA] Service worker registration error:', err);
        });
    });

    // When the new worker takes control, reload to ensure user is on latest version
    let isRefreshing = false;
    navigator.serviceWorker.addEventListener('controllerchange', () => {
      if (isRefreshing) return;
      isRefreshing = true;
      console.log('[PWA] Service Worker controller changed -> Tự động nạp bản cập nhật mới.');
      window.location.reload();
    });
  }

  // -----------------------------------------------------------------
  // AUTO-RELOAD ON RESUME FROM BACKGROUND (PWA / Mobile / Standalone)
  // Khi người dùng mở lại app từ nền điện thoại (app switcher / icon),
  // tự động reload lại trang để nạp dữ liệu và giao diện mới nhất.
  // -----------------------------------------------------------------
  let lastHiddenTimestamp = 0;
  const MIN_BACKGROUND_MS = 2500;   // Đi vào nền >= 2.5s thì khi mở lại sẽ reload
  const RELOAD_COOLDOWN_MS = 5000;  // Chống reload lặp lại trong vòng 5s

  function hasUnsavedDraft() {
    try {
      const activeEl = document.activeElement;
      if (activeEl && (activeEl.tagName === 'INPUT' || activeEl.tagName === 'TEXTAREA')) {
        if (activeEl.value && activeEl.value.trim().length > 0) return true;
      }
      const essay = document.getElementById('essayTextarea');
      if (essay && essay.value && essay.value.trim().length > 10) return true;
    } catch (e) {}
    return false;
  }

  function triggerResumeReload(reason) {
    const now = Date.now();
    const lastReload = parseInt(sessionStorage.getItem('taiga_last_resume_reload') || '0', 10);
    if (now - lastReload < RELOAD_COOLDOWN_MS) {
      return;
    }

    if (hasUnsavedDraft()) {
      console.log('[PWA] Bỏ qua reload (' + reason + ') vì đang có nội dung nhập dở dang.');
      return;
    }

    sessionStorage.setItem('taiga_last_resume_reload', String(now));
    console.log('[PWA] Đã phát hiện mở lại app từ nền (' + reason + ') -> Đang tải lại trang...');

    // Kích hoạt kiểm tra Service Worker cập nhật ngay lập tức
    if (swRegistration && typeof swRegistration.update === 'function') {
      swRegistration.update().catch(() => {});
    } else if ('serviceWorker' in navigator) {
      navigator.serviceWorker.getRegistration().then((reg) => {
        if (reg) reg.update().catch(() => {});
      }).catch(() => {});
    }

    // Tự động tải lại trang với cache revalidation
    window.location.reload();
  }

  // A. Sự kiện visibilitychange (Chuẩn phát hiện app ẩn / hiện lại trên iOS & Android)
  document.addEventListener('visibilitychange', () => {
    if (document.visibilityState === 'hidden') {
      lastHiddenTimestamp = Date.now();
      sessionStorage.setItem('taiga_last_hidden_time', String(lastHiddenTimestamp));
    } else if (document.visibilityState === 'visible') {
      const savedHidden = parseInt(sessionStorage.getItem('taiga_last_hidden_time') || '0', 10);
      const hiddenSince = lastHiddenTimestamp || savedHidden;
      if (hiddenSince > 0 && (Date.now() - hiddenSince) >= MIN_BACKGROUND_MS) {
        triggerResumeReload('visibilitychange');
      }
      lastHiddenTimestamp = 0;
      sessionStorage.removeItem('taiga_last_hidden_time');
    }
  });

  // B. Sự kiện pageshow (Đặc biệt quan trọng cho iOS Safari / WebKit khi mở lại từ bfcache)
  window.addEventListener('pageshow', (event) => {
    if (event.persisted) {
      triggerResumeReload('pageshow-persisted');
    }
  });

  // C. Sự kiện window focus & blur (Bổ trợ cho các trình duyệt Android WebAPK)
  window.addEventListener('blur', () => {
    lastHiddenTimestamp = Date.now();
    sessionStorage.setItem('taiga_last_hidden_time', String(lastHiddenTimestamp));
  });

  window.addEventListener('focus', () => {
    const savedHidden = parseInt(sessionStorage.getItem('taiga_last_hidden_time') || '0', 10);
    const hiddenSince = lastHiddenTimestamp || savedHidden;
    if (hiddenSince > 0 && (Date.now() - hiddenSince) >= MIN_BACKGROUND_MS) {
      triggerResumeReload('window-focus');
    }
    lastHiddenTimestamp = 0;
    sessionStorage.removeItem('taiga_last_hidden_time');
  });

  // 2. Track Native Install Prompt (Desktop Chrome / Edge / Android)
  let deferredPrompt = null;
  window.addEventListener('beforeinstallprompt', (e) => {
    e.preventDefault();
    deferredPrompt = e;
    // Show any desktop direct-install triggers
    const desktopBtns = document.querySelectorAll('.pwa-desktop-native-btn');
    desktopBtns.forEach((btn) => {
      btn.style.display = 'inline-flex';
    });
  });

  window.addEventListener('appinstalled', () => {
    deferredPrompt = null;
    showToast('🎉 Đã cài đặt TAIGA TALK thành công!');
    const modal = document.getElementById('pwaModalOverlay');
    if (modal) modal.classList.remove('open');
  });

  // Check if currently running as standalone PWA
  const isStandalone =
    window.matchMedia('(display-mode: standalone)').matches ||
    window.navigator.standalone === true;

  // Detect Mobile OS
  function detectOS() {
    const ua = navigator.userAgent || navigator.vendor || window.opera || '';
    if (/iPad|iPhone|iPod/.test(ua) && !window.MSStream) return 'ios';
    if (/android/i.test(ua)) return 'android';
    return 'desktop';
  }

  // Toast Notification helper
  function showToast(msg) {
    let toast = document.getElementById('pwaToast');
    if (!toast) {
      toast = document.createElement('div');
      toast.id = 'pwaToast';
      toast.className = 'pwa-toast';
      document.body.appendChild(toast);
    }
    toast.textContent = msg;
    toast.classList.add('show');
    setTimeout(() => {
      toast.classList.remove('show');
    }, 3000);
  }

  // 3. Build & Inject DOM Elements
  function initPwaUI() {
    // Inject Modal HTML
    if (!document.getElementById('pwaModalOverlay')) {
      const modal = document.createElement('div');
      modal.id = 'pwaModalOverlay';
      modal.className = 'pwa-modal-overlay';
      modal.setAttribute('role', 'dialog');
      modal.setAttribute('aria-modal', 'true');
      modal.setAttribute('aria-labelledby', 'pwaModalTitle');

      const targetUrl = window.location.origin + '/home';
      const qrEndpoint = `/api/pwa/qr?url=${encodeURIComponent(targetUrl)}`;
      const userOS = detectOS();

      modal.innerHTML = `
        <div class="pwa-modal-card">
          <button class="pwa-modal-close" id="pwaModalClose" aria-label="Đóng">&times;</button>
          
          <div class="pwa-modal-header">
            <div class="pwa-modal-brand">
              <span class="pwa-modal-flag">🇷🇺</span>
              <span class="pwa-modal-logo">TAIGA TALK</span>
              <span class="pwa-modal-badge">PWA App</span>
            </div>
            <h2 id="pwaModalTitle" class="pwa-modal-title">Cài App Về Điện Thoại & Máy Tính</h2>
            <p class="pwa-modal-desc">Trải nghiệm toàn màn hình như ứng dụng tải từ App Store, mượt mà và không che giao diện.</p>
          </div>

          <div class="pwa-modal-body">
            <!-- QR Section (Center for Desktop, Scan on Mobile) -->
            <div class="pwa-qr-section">
              <div class="pwa-qr-box" id="pwaQrBox">
                <img src="${qrEndpoint}" alt="Mã QR cài đặt TAIGA TALK" class="pwa-qr-img" id="pwaQrImg" loading="lazy" />
                <div class="pwa-qr-center-icon">🇷🇺</div>
              </div>
              <div class="pwa-qr-caption">
                <span class="pwa-pulse-dot"></span>
                <span>Quét bằng <b>Camera điện thoại</b> để mở app ngay</span>
              </div>
              
              <div class="pwa-link-copy-row">
                <input type="text" class="pwa-link-input" id="pwaLinkInput" value="${targetUrl}" readonly />
                <button type="button" class="pwa-btn-copy" id="pwaBtnCopy" title="Sao chép link">
                  <span>📋 Sao chép</span>
                </button>
              </div>
            </div>

            <!-- OS Instructions Tabs -->
            <div class="pwa-instructions-section">
              <div class="pwa-os-tabs">
                <button type="button" class="pwa-os-tab ${userOS === 'ios' ? 'active' : ''}" data-target="pwaTabIos">
                  🍏 iPhone / iPad
                </button>
                <button type="button" class="pwa-os-tab ${userOS === 'android' ? 'active' : ''}" data-target="pwaTabAndroid">
                  🤖 Android
                </button>
                <button type="button" class="pwa-os-tab ${userOS === 'desktop' ? 'active' : ''}" data-target="pwaTabDesktop">
                  💻 Máy tính (PC/Laptop)
                </button>
              </div>

              <!-- iOS Guide -->
              <div class="pwa-tab-pane ${userOS === 'ios' ? 'active' : ''}" id="pwaTabIos">
                <ol class="pwa-step-list">
                  <li>
                    <span class="pwa-step-num">1</span>
                    <div class="pwa-step-content">
                      <b>Quét mã QR</b> bằng Camera iPhone để mở trang trên <b>Safari</b>.
                    </div>
                  </li>
                  <li>
                    <span class="pwa-step-num">2</span>
                    <div class="pwa-step-content">
                      Nhấn biểu tượng <b>Chia sẻ</b> (<svg class="pwa-inline-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M4 12v8a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2v-8"></path><polyline points="16 6 12 2 8 6"></polyline><line x1="12" y1="2" x2="12" y2="15"></line></svg> nút vuông có mũi tên lên) ở thanh dưới cùng Safari.
                    </div>
                  </li>
                  <li>
                    <span class="pwa-step-num">3</span>
                    <div class="pwa-step-content">
                      Cuộn xuống và chọn <b>"Thêm vào Màn hình chính"</b> (<i>Add to Home Screen</i>).
                    </div>
                  </li>
                  <li>
                    <span class="pwa-step-num">4</span>
                    <div class="pwa-step-content">
                      Nhấn <b>"Thêm" (Add)</b> góc trên phải. Biểu tượng Taiga Talk sẽ xuất hiện ngoài màn hình chính như app thật!
                    </div>
                  </li>
                </ol>
              </div>

              <!-- Android Guide -->
              <div class="pwa-tab-pane ${userOS === 'android' ? 'active' : ''}" id="pwaTabAndroid">
                <ol class="pwa-step-list">
                  <li>
                    <span class="pwa-step-num">1</span>
                    <div class="pwa-step-content">
                      <b>Quét mã QR</b> bằng Camera hoặc Google Lens để mở trên <b>Chrome</b>.
                    </div>
                  </li>
                  <li>
                    <span class="pwa-step-num">2</span>
                    <div class="pwa-step-content">
                      Chrome sẽ hiện thông báo chân trang <b>"Thêm TAIGA TALK vào màn hình chính"</b>.
                    </div>
                  </li>
                  <li>
                    <span class="pwa-step-num">3</span>
                    <div class="pwa-step-content">
                      Nếu không thấy, nhấn biểu tượng <b>3 chấm (⋮)</b> góc trên phải ➔ Chọn <b>"Cài đặt ứng dụng"</b>.
                    </div>
                  </li>
                  <li>
                    <span class="pwa-step-num">4</span>
                    <div class="pwa-step-content">
                      Nhấn <b>Cài đặt</b>. Ứng dụng sẽ được tích hợp vào danh sách App trên máy bạn.
                    </div>
                  </li>
                </ol>
              </div>

              <!-- Desktop Guide -->
              <div class="pwa-tab-pane ${userOS === 'desktop' ? 'active' : ''}" id="pwaTabDesktop">
                <div class="pwa-desktop-install-box">
                  <p>Bạn có thể cài đặt TAIGA TALK thành ứng dụng máy tính độc lập trên Chrome hoặc Microsoft Edge:</p>
                  <button type="button" class="pwa-desktop-native-btn" id="btnPwaNativeInstall">
                    💻 Cài Đặt Ứng Dụng Ngay Trên Máy Tính
                  </button>
                  <div class="pwa-desktop-hint">
                    Hoặc bấm biểu tượng <b>Cài đặt (Install / ⊕)</b> nằm ở góc phải thanh địa chỉ trình duyệt máy tính của bạn.
                  </div>
                </div>
              </div>
            </div>
          </div>

          <div class="pwa-modal-footer">
            <span class="pwa-footer-badge">⚡ Mở tức thì • Không cần tải qua App Store / CH Play • Không tốn dung lượng</span>
          </div>
        </div>
      `;

      document.body.appendChild(modal);

      // Bind Modal events
      const closeBtn = document.getElementById('pwaModalClose');
      closeBtn.addEventListener('click', () => modal.classList.remove('open'));
      modal.addEventListener('click', (e) => {
        if (e.target === modal) modal.classList.remove('open');
      });

      // Copy Link button
      const copyBtn = document.getElementById('pwaBtnCopy');
      const linkInput = document.getElementById('pwaLinkInput');
      copyBtn.addEventListener('click', () => {
        linkInput.select();
        navigator.clipboard.writeText(linkInput.value).then(() => {
          showToast('📋 Đã sao chép liên kết vào bộ nhớ tạm!');
        }).catch(() => {
          showToast('Liên kết: ' + linkInput.value);
        });
      });

      // Tab switcher
      const tabBtns = modal.querySelectorAll('.pwa-os-tab');
      const tabPanes = modal.querySelectorAll('.pwa-tab-pane');
      tabBtns.forEach((btn) => {
        btn.addEventListener('click', () => {
          tabBtns.forEach((b) => b.classList.remove('active'));
          tabPanes.forEach((p) => p.classList.remove('active'));
          btn.classList.add('active');
          const target = btn.getAttribute('data-target');
          const pane = document.getElementById(target);
          if (pane) pane.classList.add('active');
        });
      });

      // Native desktop install button
      const nativeBtn = document.getElementById('btnPwaNativeInstall');
      if (nativeBtn) {
        nativeBtn.addEventListener('click', async () => {
          if (deferredPrompt) {
            deferredPrompt.prompt();
            const { outcome } = await deferredPrompt.userChoice;
            if (outcome === 'accepted') {
              showToast('🎉 Đang tiến hành cài đặt!');
            }
            deferredPrompt = null;
          } else {
            showToast('💡 Bấm biểu tượng Cài đặt (⊕) trên thanh địa chỉ URL máy tính.');
          }
        });
      }
    }

    // 4. Inject Floating Corner Widget (Only for Desktop / Laptop screens)
    if (!isStandalone && !document.getElementById('taigaPwaCornerWidget')) {
      const isMinimized = localStorage.getItem('taiga_pwa_minimized') === 'true';

      const cornerWidget = document.createElement('aside');
      cornerWidget.id = 'taigaPwaCornerWidget';
      cornerWidget.className = `pwa-corner-widget ${isMinimized ? 'minimized' : ''}`;
      cornerWidget.setAttribute('aria-label', 'Cài App TAIGA TALK');

      cornerWidget.innerHTML = `
        <div class="pwa-widget-badge" id="pwaWidgetTrigger" title="Nhấn để mở mã QR cài App">
          <div class="pwa-widget-qr-icon">
            <svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
              <rect x="3" y="3" width="7" height="7" rx="1.5"></rect>
              <rect x="14" y="3" width="7" height="7" rx="1.5"></rect>
              <rect x="14" y="14" width="7" height="7" rx="1.5"></rect>
              <rect x="3" y="14" width="7" height="7" rx="1.5"></rect>
              <circle cx="6.5" cy="6.5" r="1" fill="currentColor"></circle>
              <circle cx="17.5" cy="6.5" r="1" fill="currentColor"></circle>
              <circle cx="6.5" cy="17.5" r="1" fill="currentColor"></circle>
              <circle cx="17.5" cy="17.5" r="1" fill="currentColor"></circle>
            </svg>
          </div>
          <div class="pwa-widget-text">
            <span class="pwa-widget-title">📲 Cài App về ĐT</span>
            <span class="pwa-widget-sub">Quét mã QR góc màn hình</span>
          </div>
        </div>
        <button type="button" class="pwa-widget-toggle" id="pwaWidgetToggle" title="${isMinimized ? 'Mở rộng' : 'Thu nhỏ'}" aria-label="Thu nhỏ">
          ${isMinimized ? '📲' : '&times;'}
        </button>
      `;

      document.body.appendChild(cornerWidget);

      const trigger = document.getElementById('pwaWidgetTrigger');
      const toggle = document.getElementById('pwaWidgetToggle');

      trigger.addEventListener('click', () => {
        openPwaModal();
      });

      toggle.addEventListener('click', (e) => {
        e.stopPropagation();
        cornerWidget.classList.toggle('minimized');
        const minNow = cornerWidget.classList.contains('minimized');
        localStorage.setItem('taiga_pwa_minimized', minNow ? 'true' : 'false');
        toggle.innerHTML = minNow ? '📲' : '&times;';
        toggle.title = minNow ? 'Mở rộng' : 'Thu nhỏ';
      });
    }

    // 5. Inject Topbar Button into .ff-topbar-right if not present
    const topbarRight = document.querySelector('.ff-topbar-right');
    if (topbarRight && !document.getElementById('btnHeaderInstallApp')) {
      const topBtn = document.createElement('button');
      topBtn.type = 'button';
      topBtn.id = 'btnHeaderInstallApp';
      topBtn.className = 'ff-btn-install-app';
      topBtn.title = 'Cài app về điện thoại hoặc máy tính';
      topBtn.setAttribute('aria-label', 'Cài đặt App');
      topBtn.innerHTML = `<span>📲 Cài App</span>`;
      topBtn.addEventListener('click', () => openPwaModal());

      // Insert before the last element (e.g. before "Lịch học" or at the start of stats)
      topbarRight.insertBefore(topBtn, topbarRight.firstChild);
    }

    // Attach click to any existing button with data-pwa-install or .btn-pwa-install
    document.querySelectorAll('[data-pwa-install], .btn-pwa-install').forEach((btn) => {
      btn.addEventListener('click', (e) => {
        e.preventDefault();
        openPwaModal();
      });
    });
  }

  // Open the Modal
  function openPwaModal() {
    const modal = document.getElementById('pwaModalOverlay');
    if (modal) {
      modal.classList.add('open');
    }
  }

  // Expose global helper
  window.TaigaPWA = {
    openModal: openPwaModal,
    isStandalone: isStandalone,
    detectOS: detectOS,
  };

  // Run on DOM Ready
  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', initPwaUI);
  } else {
    initPwaUI();
  }
})();
