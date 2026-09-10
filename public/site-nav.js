// ==========================================================================
// TAIGA Center - Unified Navigation & Routing Script (site-nav.js)
// ==========================================================================

(function () {
  function initNav() {
    const currentPath = window.location.pathname.replace(/\/$/, '') || '/home';
    const isHomePage = currentPath === '' || currentPath === '/home' || currentPath === '/index.html';
    
    // 1. Highlight Active Item
    const navLinks = document.querySelectorAll('.header .nav a');
    navLinks.forEach((link) => {
      link.classList.remove('active');
      const navTarget = link.getAttribute('data-nav');
      
      if (isHomePage && navTarget === 'home') {
        link.classList.add('active');
      } else if (currentPath.includes('roadmap') && navTarget === 'roadmap') {
        link.classList.add('active');
      } else if (currentPath.includes('learn') && navTarget === 'roadmap') {
        link.classList.add('active');
      } else if (currentPath.includes('articles') && navTarget === 'articles') {
        link.classList.add('active');
      } else if (currentPath.includes('contact') && navTarget === 'contact') {
        link.classList.add('active');
      } else if (currentPath.includes('app') && navTarget === 'app') {
        link.classList.add('active');
      }
    });

    // 2. Click Handling for Seamless Transitions
    navLinks.forEach((link) => {
      link.addEventListener('click', function (e) {
        const href = this.getAttribute('href');
        const navTarget = this.getAttribute('data-nav');

        // Case A: Clicking "Trang chu" while already on /home
        if (isHomePage && (navTarget === 'home' || href === '/home' || href === '#home')) {
          e.preventDefault();
          window.scrollTo({ top: 0, behavior: 'smooth' });
          history.pushState(null, '', '/home');
          setActive(this);
          return;
        }

        // Case B: Clicking an in-page section link (Courses, About, etc.) while on /home
        if (isHomePage && href.includes('#')) {
          const hashIndex = href.indexOf('#');
          const targetId = href.substring(hashIndex);
          const targetElem = document.querySelector(targetId);
          if (targetElem) {
            e.preventDefault();
            targetElem.scrollIntoView({ behavior: 'smooth' });
            history.pushState(null, '', targetId);
            setActive(this);
            return;
          }
        }

        // Case C: Clicking the current subpage again -> scroll to top
        if (
          (currentPath.includes('roadmap') && navTarget === 'roadmap') ||
          (currentPath.includes('articles') && navTarget === 'articles') ||
          (currentPath.includes('contact') && navTarget === 'contact')
        ) {
          e.preventDefault();
          window.scrollTo({ top: 0, behavior: 'smooth' });
          return;
        }

        // Other cases: normal navigation to other pages (e.g. /home#courses from /roadmap)
      });
    });

    function setActive(activeLink) {
      navLinks.forEach((l) => l.classList.remove('active'));
      activeLink.classList.add('active');
    }

    // 3. Scroll spy on /home to update active state for #courses and #about
    if (isHomePage) {
      const observerSections = [
        { id: 'home', nav: 'home' },
        { id: 'about', nav: 'about' },
        { id: 'courses', nav: 'courses' }
      ];

      window.addEventListener('scroll', () => {
        const scrollPosition = window.scrollY + 120;
        let currentSectionNav = 'home';

        observerSections.forEach((sec) => {
          const el = document.getElementById(sec.id);
          if (el && el.offsetTop <= scrollPosition) {
            currentSectionNav = sec.nav;
          }
        });

        navLinks.forEach((link) => {
          if (link.getAttribute('data-nav') === currentSectionNav) {
            link.classList.add('active');
          } else {
            link.classList.remove('active');
          }
        });
      }, { passive: true });
    }

    // 4. Smooth scroll on landing with hash (e.g. redirected from /roadmap to /home#courses)
    if (window.location.hash) {
      setTimeout(() => {
        const target = document.querySelector(window.location.hash);
        if (target) {
          target.scrollIntoView({ behavior: 'smooth' });
        }
      }, 150);
    }

    // 5. Language Switcher (Ru-Vi) Button
    const langBtn = document.querySelector('.header .language-btn');
    if (langBtn) {
      langBtn.addEventListener('click', () => {
        showToast('🇷🇺 Tiếng Nga & Tiếng Việt: Nội dung song ngữ đang được tối ưu hóa!');
      });
    }
  }

  // Toast Helper
  function showToast(message) {
    let toast = document.querySelector('.site-toast');
    if (!toast) {
      toast = document.createElement('div');
      toast.className = 'site-toast';
      document.body.appendChild(toast);
    }
    toast.textContent = message;
    toast.classList.add('show');
    clearTimeout(window._toastTimer);
    window._toastTimer = setTimeout(() => {
      toast.classList.remove('show');
    }, 2800);
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', initNav);
  } else {
    initNav();
  }
})();
