// ==========================================================================
// TAIGA Center - Modern Tree Navigation Script (site-nav.js)
// ==========================================================================

(function () {
  function initNav() {
    const currentPath = window.location.pathname.replace(/\/$/, '') || '/home';
    const isHomePage = currentPath === '' || currentPath === '/home' || currentPath === '/index.html';
    
    const header = document.querySelector('.header');
    if (!header) return;

    const nav = header.querySelector('.nav');
    const mobileToggle = header.querySelector('.mobile-toggle');
    const dropdowns = header.querySelectorAll('.nav-dropdown');
    const navLinks = header.querySelectorAll('.nav a, .dropdown-link');

    // 1. Mobile Hamburger Toggle
    if (mobileToggle && nav) {
      mobileToggle.addEventListener('click', (e) => {
        e.stopPropagation();
        nav.classList.toggle('nav-open');
        const isOpen = nav.classList.contains('nav-open');
        mobileToggle.setAttribute('aria-expanded', isOpen ? 'true' : 'false');
      });
    }

    // 2. Dropdown Tree Toggle (Click & Touch support)
    dropdowns.forEach((dropdown) => {
      const toggleBtn = dropdown.querySelector('.dropdown-toggle');
      if (!toggleBtn) return;

      toggleBtn.addEventListener('click', (e) => {
        e.preventDefault();
        e.stopPropagation();
        
        // Close other open dropdowns
        dropdowns.forEach((other) => {
          if (other !== dropdown) other.classList.remove('open');
        });

        dropdown.classList.toggle('open');
        const isOpen = dropdown.classList.contains('open');
        toggleBtn.setAttribute('aria-expanded', isOpen ? 'true' : 'false');
      });
    });

    // Close dropdowns and mobile nav on click outside
    document.addEventListener('click', (e) => {
      if (!header.contains(e.target)) {
        dropdowns.forEach((d) => d.classList.remove('open'));
        if (nav) nav.classList.remove('nav-open');
      }
    });

    // 3. Highlight Active Item & Dropdown Parent
    function updateActive() {
      // Clear current actives
      header.querySelectorAll('.nav-item, .dropdown-link, .nav-dropdown').forEach((el) => {
        el.classList.remove('active');
      });

      if (isHomePage) {
        const homeLink = header.querySelector('[data-nav="home"]');
        if (homeLink) homeLink.classList.add('active');
      } else if (currentPath.includes('roadmap') || currentPath.includes('learn')) {
        highlightTree('roadmap');
      } else if (currentPath.includes('articles')) {
        highlightTree('articles');
      } else if (currentPath.includes('contact')) {
        highlightTree('contact');
      } else if (currentPath.includes('app')) {
        const appLink = header.querySelector('[data-nav="app"]');
        if (appLink) appLink.classList.add('active');
      }
    }

    function highlightTree(navKey) {
      const targetLink = header.querySelector(`[data-nav="${navKey}"]`);
      if (targetLink) {
        targetLink.classList.add('active');
        const parentDropdown = targetLink.closest('.nav-dropdown');
        if (parentDropdown) {
          parentDropdown.classList.add('active');
        }
      }
    }

    updateActive();

    // 4. Smooth Anchor Scrolling & Cross-page Navigation
    navLinks.forEach((link) => {
      link.addEventListener('click', function (e) {
        const href = this.getAttribute('href');
        const navTarget = this.getAttribute('data-nav');

        // Close mobile nav when link clicked
        if (nav) nav.classList.remove('nav-open');
        dropdowns.forEach((d) => d.classList.remove('open'));

        // Case A: Clicking "Trang chu" while on /home
        if (isHomePage && (navTarget === 'home' || href === '/home' || href === '#home')) {
          e.preventDefault();
          window.scrollTo({ top: 0, behavior: 'smooth' });
          history.pushState(null, '', '/home');
          updateActive();
          return;
        }

        // Case B: Clicking an in-page section link (#courses, #about) while on /home
        if (isHomePage && href && href.includes('#')) {
          const hashIndex = href.indexOf('#');
          const targetId = href.substring(hashIndex);
          const targetElem = document.querySelector(targetId);
          if (targetElem) {
            e.preventDefault();
            targetElem.scrollIntoView({ behavior: 'smooth' });
            history.pushState(null, '', targetId);
            return;
          }
        }

        // Case C: Re-clicking current subpage -> scroll to top
        if (
          (currentPath.includes('roadmap') && navTarget === 'roadmap') ||
          (currentPath.includes('articles') && navTarget === 'articles') ||
          (currentPath.includes('contact') && navTarget === 'contact')
        ) {
          e.preventDefault();
          window.scrollTo({ top: 0, behavior: 'smooth' });
          return;
        }
      });
    });

    // 5. Scroll spy on /home for section links
    if (isHomePage) {
      const sections = [
        { id: 'courses', nav: 'courses' },
        { id: 'about', nav: 'about' }
      ];

      window.addEventListener('scroll', () => {
        const scrollPos = window.scrollY + 140;
        let matched = false;

        sections.forEach((sec) => {
          const el = document.getElementById(sec.id);
          if (el && el.offsetTop <= scrollPos && (el.offsetTop + el.offsetHeight) > scrollPos) {
            matched = true;
            header.querySelectorAll('.dropdown-link, .nav-item').forEach((l) => l.classList.remove('active'));
            highlightTree(sec.nav);
          }
        });

        if (!matched && window.scrollY < 300) {
          header.querySelectorAll('.dropdown-link, .nav-dropdown').forEach((l) => l.classList.remove('active'));
          const homeLink = header.querySelector('[data-nav="home"]');
          if (homeLink) homeLink.classList.add('active');
        }
      }, { passive: true });
    }

    // 6. Smooth landing when navigating with hash from another page
    if (window.location.hash) {
      setTimeout(() => {
        const target = document.querySelector(window.location.hash);
        if (target) {
          target.scrollIntoView({ behavior: 'smooth' });
        }
      }, 150);
    }

    // 7. Ru-Vi Language Switcher Toast
    const langBtn = header.querySelector('.language-btn');
    if (langBtn) {
      langBtn.addEventListener('click', () => {
        showToast('🇷🇺 Tiếng Nga & Tiếng Việt: Nội dung song ngữ đang được tối ưu hóa!');
      });
    }
  }

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
