/**
 * The Ghee Project — main.js
 * Shared JavaScript for all pages
 *
 * Covers:
 *  1. Scroll-triggered fade-up animations (IntersectionObserver)
 *  2. Staggered child animations
 *  3. Mobile nav hamburger toggle
 *  4. Smooth scroll for anchor links
 *  5. Sticky nav active state on scroll
 *  6. Active nav link highlighting by section
 *  7. Google Apps Script form handler (wholesale + order forms)
 *  8. Shop variant selector & filter
 *  9. FAQ accordion (wholesale page)
 * 10. Scroll-to-top button
 */

/* ─────────────────────────────────────────────
   1. SCROLL ANIMATIONS — fade-up + stagger
───────────────────────────────────────────── */

const animObserver = new IntersectionObserver(
  (entries) => {
    entries.forEach((entry) => {
      if (entry.isIntersecting) {
        entry.target.classList.add('visible');
        animObserver.unobserve(entry.target); // animate once only
      }
    });
  },
  { threshold: 0.12, rootMargin: '0px 0px -40px 0px' }
);

document.querySelectorAll('.fade-up, .stagger > *').forEach((el) => {
  animObserver.observe(el);
});

/* ─────────────────────────────────────────────
   2. MOBILE NAV — hamburger toggle
───────────────────────────────────────────── */

const hamburger = document.querySelector('.nav-hamburger');
const navLinks  = document.querySelector('.nav-links');
const navOverlay = document.querySelector('.nav-overlay');

if (hamburger && navLinks) {
  hamburger.addEventListener('click', () => {
    const isOpen = navLinks.classList.toggle('nav-open');
    hamburger.setAttribute('aria-expanded', isOpen);
    hamburger.classList.toggle('active', isOpen);
    if (navOverlay) navOverlay.classList.toggle('visible', isOpen);
    document.body.style.overflow = isOpen ? 'hidden' : '';
  });

  // Close on overlay click
  if (navOverlay) {
    navOverlay.addEventListener('click', closeNav);
  }

  // Close on nav link click (single-page navigation)
  navLinks.querySelectorAll('a').forEach((link) => {
    link.addEventListener('click', closeNav);
  });

  // Close on Escape key
  document.addEventListener('keydown', (e) => {
    if (e.key === 'Escape') closeNav();
  });
}

function closeNav() {
  if (!navLinks) return;
  navLinks.classList.remove('nav-open');
  if (hamburger) {
    hamburger.setAttribute('aria-expanded', 'false');
    hamburger.classList.remove('active');
  }
  if (navOverlay) navOverlay.classList.remove('visible');
  document.body.style.overflow = '';
}

/* ─────────────────────────────────────────────
   3. STICKY NAV — add scrolled class on scroll
───────────────────────────────────────────── */

const siteNav = document.querySelector('.site-nav');

if (siteNav) {
  const handleNavScroll = () => {
    siteNav.classList.toggle('scrolled', window.scrollY > 60);
  };
  window.addEventListener('scroll', handleNavScroll, { passive: true });
  handleNavScroll(); // run once on load
}

/* ─────────────────────────────────────────────
   4. ACTIVE NAV LINK — highlight by scroll position
───────────────────────────────────────────── */

const navAnchors = document.querySelectorAll('.nav-links a[href^="#"]');
const sections   = Array.from(navAnchors)
  .map((a) => document.querySelector(a.getAttribute('href')))
  .filter(Boolean);

if (sections.length > 0) {
  const sectionObserver = new IntersectionObserver(
    (entries) => {
      entries.forEach((entry) => {
        if (entry.isIntersecting) {
          const id = entry.target.getAttribute('id');
          navAnchors.forEach((a) => {
            a.classList.toggle('active', a.getAttribute('href') === `#${id}`);
          });
        }
      });
    },
    { rootMargin: '-40% 0px -55% 0px' }
  );
  sections.forEach((s) => sectionObserver.observe(s));
}

/* ─────────────────────────────────────────────
   5. SMOOTH SCROLL — anchor links
───────────────────────────────────────────── */

document.querySelectorAll('a[href^="#"]').forEach((anchor) => {
  anchor.addEventListener('click', (e) => {
    const target = document.querySelector(anchor.getAttribute('href'));
    if (!target) return;
    e.preventDefault();
    const navHeight = siteNav ? siteNav.offsetHeight : 0;
    const top = target.getBoundingClientRect().top + window.scrollY - navHeight - 16;
    window.scrollTo({ top, behavior: 'smooth' });
  });
});

/* ─────────────────────────────────────────────
   6. GOOGLE APPS SCRIPT FORM HANDLER
   Handles both the wholesale enquiry and
   the shop / order forms.
───────────────────────────────────────────── */

/**
 * SETUP REQUIRED:
 * Replace 'YOUR_SCRIPT_URL' in each HTML form's data-script-url attribute
 * with your deployed Google Apps Script Web App URL.
 * See SETUP.md for the full 5-minute setup guide.
 */

document.querySelectorAll('form[data-script-url]').forEach((form) => {
  form.addEventListener('submit', async (e) => {
    e.preventDefault();

    const scriptUrl = form.dataset.scriptUrl;
    const statusEl  = form.querySelector('.form-status');
    const submitBtn = form.querySelector('button[type="submit"]');

    // Config warning — URL not set yet
    if (!scriptUrl || scriptUrl === 'YOUR_SCRIPT_URL') {
      if (statusEl) {
        statusEl.textContent = 'Form not yet connected. See SETUP.md to link your Google Sheet.';
        statusEl.className = 'form-status form-status--warn';
      }
      return;
    }

    // Disable button during submission
    if (submitBtn) {
      submitBtn.disabled = true;
      submitBtn.textContent = 'Sending…';
    }

    const formData = new FormData(form);
    const payload  = {};
    formData.forEach((value, key) => { payload[key] = value; });

    // Add metadata
    payload._page      = window.location.pathname;
    payload._timestamp = new Date().toISOString();

    try {
      // WHY application/x-www-form-urlencoded:
      //   Google Apps Script redirects POST requests from script.google.com to
      //   script.googleusercontent.com. During that redirect the raw POST body
      //   (text/plain or application/json) is dropped and e.postData becomes null.
      //   Encoding the JSON as a named form field ("payload=...") lets Apps Script
      //   read it from e.parameter.payload, which survives the redirect reliably.
      //   application/x-www-form-urlencoded is also a CORS-safe content type,
      //   so it works correctly with mode:'no-cors'.
      await fetch(scriptUrl, {
        method:  'POST',
        headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
        body:    'payload=' + encodeURIComponent(JSON.stringify(payload)),
        mode:    'no-cors',
      });

      // no-cors always returns an opaque response (status 0) — treat any resolve as success
      if (statusEl) {
        statusEl.textContent = 'Message sent — we\'ll be in touch shortly.';
        statusEl.className = 'form-status form-status--success';
      }
      form.reset();
    } catch (err) {
      if (statusEl) {
        statusEl.textContent = 'Something went wrong. Please email us at hello@thegeeproject.com.au';
        statusEl.className = 'form-status form-status--error';
      }
    } finally {
      if (submitBtn) {
        submitBtn.disabled = false;
        submitBtn.textContent = submitBtn.dataset.label || 'Send Enquiry';
      }
    }
  });
});

/* ─────────────────────────────────────────────
   7. FAQ ACCORDION (wholesale page)
───────────────────────────────────────────── */

document.querySelectorAll('.faq-item').forEach((item) => {
  const trigger = item.querySelector('.faq-question');
  const answer  = item.querySelector('.faq-answer');
  if (!trigger || !answer) return;

  trigger.setAttribute('aria-expanded', 'false');

  trigger.addEventListener('click', () => {
    const isOpen = item.classList.toggle('open');
    trigger.setAttribute('aria-expanded', isOpen);

    if (isOpen) {
      answer.style.maxHeight = answer.scrollHeight + 'px';
    } else {
      answer.style.maxHeight = '0';
    }

    // Close siblings
    const siblings = item.parentElement.querySelectorAll('.faq-item');
    siblings.forEach((sib) => {
      if (sib !== item && sib.classList.contains('open')) {
        sib.classList.remove('open');
        sib.querySelector('.faq-question')?.setAttribute('aria-expanded', 'false');
        const sibAnswer = sib.querySelector('.faq-answer');
        if (sibAnswer) sibAnswer.style.maxHeight = '0';
      }
    });
  });
});

/* ─────────────────────────────────────────────
   8. SHOP — filter buttons
───────────────────────────────────────────── */

const filterBtns = document.querySelectorAll('.shop-filter-btn');
const shopCards  = document.querySelectorAll('.shop-card');

if (filterBtns.length > 0) {
  filterBtns.forEach((btn) => {
    btn.addEventListener('click', () => {
      filterBtns.forEach((b) => b.classList.remove('active'));
      btn.classList.add('active');

      const filter = btn.dataset.filter;

      shopCards.forEach((card) => {
        if (filter === 'all' || card.dataset.category === filter) {
          card.style.display = '';
          // Re-trigger fade animation
          card.classList.remove('visible');
          requestAnimationFrame(() => {
            requestAnimationFrame(() => card.classList.add('visible'));
          });
        } else {
          card.style.display = 'none';
        }
      });
    });
  });
}

/* ─────────────────────────────────────────────
   9. SCROLL-TO-TOP BUTTON
───────────────────────────────────────────── */

const scrollTopBtn = document.querySelector('.scroll-top');

if (scrollTopBtn) {
  window.addEventListener(
    'scroll',
    () => {
      scrollTopBtn.classList.toggle('visible', window.scrollY > 500);
    },
    { passive: true }
  );

  scrollTopBtn.addEventListener('click', () => {
    window.scrollTo({ top: 0, behavior: 'smooth' });
  });
}

/* ─────────────────────────────────────────────
   10. COOKIE BANNER (simple, non-tracking)
───────────────────────────────────────────── */

const cookieBanner = document.querySelector('.cookie-banner');
const cookieAccept = document.querySelector('.cookie-accept');

if (cookieBanner && !sessionStorage.getItem('cookieAck')) {
  // Show after brief delay so it doesn't flash on page load
  setTimeout(() => cookieBanner.classList.add('visible'), 1500);
}

if (cookieAccept) {
  cookieAccept.addEventListener('click', () => {
    sessionStorage.setItem('cookieAck', '1');
    if (cookieBanner) cookieBanner.classList.remove('visible');
  });
}
