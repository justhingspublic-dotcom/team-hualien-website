/**
 * Team Hualien Website - Main JavaScript
 */

document.addEventListener('DOMContentLoaded', function() {
  // ========================================
  // Mobile Menu Toggle
  // ========================================
  const menuToggle = document.getElementById('menuToggle');
  const mainNav = document.getElementById('mainNav');

  if (menuToggle && mainNav) {
    menuToggle.addEventListener('click', function() {
      mainNav.classList.toggle('active');
      this.classList.toggle('active');

      // Toggle hamburger animation
      const spans = this.querySelectorAll('span');
      if (this.classList.contains('active')) {
        spans[0].style.transform = 'rotate(45deg) translate(5px, 5px)';
        spans[1].style.opacity = '0';
        spans[2].style.transform = 'rotate(-45deg) translate(7px, -6px)';
      } else {
        spans[0].style.transform = 'none';
        spans[1].style.opacity = '1';
        spans[2].style.transform = 'none';
      }
    });
  }

  // Close mobile menu when clicking on a link
  const navLinks = document.querySelectorAll('.nav-link');
  navLinks.forEach(link => {
    link.addEventListener('click', function() {
      if (window.innerWidth < 768) {
        mainNav.classList.remove('active');
        menuToggle.classList.remove('active');
        const spans = menuToggle.querySelectorAll('span');
        spans[0].style.transform = 'none';
        spans[1].style.opacity = '1';
        spans[2].style.transform = 'none';
      }
    });
  });

  // ========================================
  // Header Scroll Effect
  // ========================================
  const header = document.getElementById('header');
  let lastScroll = 0;

  window.addEventListener('scroll', function() {
    const currentScroll = window.pageYOffset;

    if (currentScroll > 100) {
      header.classList.add('scrolled');
    } else {
      header.classList.remove('scrolled');
    }

    lastScroll = currentScroll;
  });

  // ========================================
  // Back to Top Button
  // ========================================
  const backToTop = document.getElementById('backToTop');

  if (backToTop) {
    window.addEventListener('scroll', function() {
      if (window.pageYOffset > 300) {
        backToTop.classList.add('visible');
      } else {
        backToTop.classList.remove('visible');
      }
    });

    backToTop.addEventListener('click', function() {
      window.scrollTo({
        top: 0,
        behavior: 'smooth'
      });
    });
  }

  // ========================================
  // Smooth Scroll for Anchor Links
  // ========================================
  document.querySelectorAll('a[href^="#"]').forEach(anchor => {
    anchor.addEventListener('click', function(e) {
      const href = this.getAttribute('href');
      if (href !== '#') {
        e.preventDefault();
        const target = document.querySelector(href);
        if (target) {
          const headerHeight = document.getElementById('header').offsetHeight;
          const targetPosition = target.getBoundingClientRect().top + window.pageYOffset - headerHeight;
          window.scrollTo({
            top: targetPosition,
            behavior: 'smooth'
          });
        }
      }
    });
  });

  // ========================================
  // Video Modal
  // ========================================
  const playBtn = document.getElementById('playVideo');
  const videoModal = document.getElementById('videoModal');
  const videoFrame = document.getElementById('videoFrame');
  const closeModal = document.querySelector('.close-modal');

  if (playBtn && videoModal) {
    playBtn.addEventListener('click', function() {
      // Replace with actual video URL
      videoFrame.src = 'https://www.youtube.com/embed/dQw4w9WgXcQ?autoplay=1';
      videoModal.style.display = 'flex';
      document.body.style.overflow = 'hidden';
    });

    closeModal.addEventListener('click', function() {
      videoModal.style.display = 'none';
      videoFrame.src = '';
      document.body.style.overflow = '';
    });

    videoModal.addEventListener('click', function(e) {
      if (e.target === videoModal) {
        videoModal.style.display = 'none';
        videoFrame.src = '';
        document.body.style.overflow = '';
      }
    });

    // Close with ESC key
    document.addEventListener('keydown', function(e) {
      if (e.key === 'Escape' && videoModal.style.display === 'flex') {
        videoModal.style.display = 'none';
        videoFrame.src = '';
        document.body.style.overflow = '';
      }
    });
  }

  // ========================================
  // Scroll Animation (Intersection Observer)
  // ========================================
  const observerOptions = {
    root: null,
    rootMargin: '0px',
    threshold: 0.1
  };

  const animateOnScroll = new IntersectionObserver(function(entries, observer) {
    entries.forEach(entry => {
      if (entry.isIntersecting) {
        entry.target.classList.add('animate-in');
        observer.unobserve(entry.target);
      }
    });
  }, observerOptions);

  // Observe elements
  document.querySelectorAll('.card, .initiative-card, .stat-item, .partner-logo').forEach(el => {
    el.style.opacity = '0';
    el.style.transform = 'translateY(20px)';
    el.style.transition = 'opacity 0.6s ease, transform 0.6s ease';
    animateOnScroll.observe(el);
  });

  // Add animation class styles
  const style = document.createElement('style');
  style.textContent = `
    .animate-in {
      opacity: 1 !important;
      transform: translateY(0) !important;
    }
  `;
  document.head.appendChild(style);

  // ========================================
  // Statistics Counter Animation
  // ========================================
  const statNumbers = document.querySelectorAll('.stat-number');

  const counterObserver = new IntersectionObserver(function(entries) {
    entries.forEach(entry => {
      if (entry.isIntersecting) {
        const target = entry.target;
        const countTo = target.getAttribute('data-count');

        if (countTo) {
          let count = 0;
          const increment = countTo / 50;
          const timer = setInterval(() => {
            count += increment;
            if (count >= countTo) {
              target.textContent = countTo.includes('+') ? countTo : countTo;
              clearInterval(timer);
            } else {
              target.textContent = Math.floor(count);
            }
          }, 30);
        }

        counterObserver.unobserve(target);
      }
    });
  }, { threshold: 0.5 });

  statNumbers.forEach(stat => {
    counterObserver.observe(stat);
  });

  // ========================================
  // Search Functionality
  // ========================================
  const searchForm = document.querySelector('.search-box');
  if (searchForm) {
    searchForm.addEventListener('submit', function(e) {
      e.preventDefault();
      const searchInput = this.querySelector('input');
      const query = searchInput.value.trim();
      if (query) {
        // Implement search logic or redirect to search results page
        console.log('Searching for:', query);
        alert('搜尋功能開發中：' + query);
      }
    });
  }

  // ========================================
  // Contact Form Validation
  // ========================================
  const contactForm = document.getElementById('contactForm');
  if (contactForm) {
    contactForm.addEventListener('submit', function(e) {
      e.preventDefault();

      const name = document.getElementById('name').value.trim();
      const email = document.getElementById('email').value.trim();
      const subject = document.getElementById('subject').value.trim();
      const message = document.getElementById('message').value.trim();

      // Simple validation
      if (!name || !email || !subject || !message) {
        alert('請填寫所有必填欄位');
        return;
      }

      // Email validation
      const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
      if (!emailRegex.test(email)) {
        alert('請輸入有效的電子郵件地址');
        return;
      }

      // Submit form (replace with actual submission logic)
      console.log('Form submitted:', { name, email, subject, message });
      alert('感謝您的留言，我們會盡快回覆！');
      this.reset();
    });
  }

  // ========================================
  // Dropdown Menu for Touch Devices
  // ========================================
  const dropdownItems = document.querySelectorAll('.nav-item.has-dropdown');

  dropdownItems.forEach(item => {
    const link = item.querySelector('.nav-link');

    link.addEventListener('click', function(e) {
      if (window.innerWidth < 768) {
        e.preventDefault();
        item.classList.toggle('dropdown-open');
      }
    });
  });

  // ========================================
  // Image Lazy Loading
  // ========================================
  if ('loading' in HTMLImageElement.prototype) {
    // Native lazy loading supported
    document.querySelectorAll('img[data-src]').forEach(img => {
      img.src = img.dataset.src;
    });
  } else {
    // Fallback for older browsers
    const lazyImages = document.querySelectorAll('img[data-src]');

    const lazyLoad = new IntersectionObserver(function(entries) {
      entries.forEach(entry => {
        if (entry.isIntersecting) {
          const img = entry.target;
          img.src = img.dataset.src;
          img.classList.add('loaded');
          lazyLoad.unobserve(img);
        }
      });
    });

    lazyImages.forEach(img => lazyLoad.observe(img));
  }

  // ========================================
  // Gallery Lightbox
  // ========================================
  const galleryItems = document.querySelectorAll('.gallery-item');

  galleryItems.forEach(item => {
    item.addEventListener('click', function() {
      const img = this.querySelector('img');
      if (img) {
        const lightbox = document.createElement('div');
        lightbox.className = 'lightbox';
        lightbox.style.cssText = 'position: fixed; top: 0; left: 0; width: 100%; height: 100%; background: rgba(0,0,0,0.9); display: flex; align-items: center; justify-content: center; z-index: 9999; cursor: pointer;';

        const lightboxImg = document.createElement('img');
        lightboxImg.src = img.src;
        lightboxImg.style.cssText = 'max-width: 90%; max-height: 90%; object-fit: contain;';

        lightbox.appendChild(lightboxImg);
        document.body.appendChild(lightbox);
        document.body.style.overflow = 'hidden';

        lightbox.addEventListener('click', function() {
          this.remove();
          document.body.style.overflow = '';
        });
      }
    });
  });

  // ========================================
  // Language Switcher
  // ========================================
  const langSwitch = document.querySelector('.lang-switch');
  if (langSwitch) {
    langSwitch.addEventListener('click', function(e) {
      e.preventDefault();
      const currentLang = document.documentElement.lang;
      const newLang = currentLang === 'zh-TW' ? 'en' : 'zh-TW';

      // Get current page path
      const currentPath = window.location.pathname;
      let newPath;

      if (currentPath.includes('-en.html')) {
        newPath = currentPath.replace('-en.html', '.html');
      } else {
        newPath = currentPath.replace('.html', '-en.html');
      }

      // Handle index page
      if (currentPath.endsWith('/') || currentPath.endsWith('/index.html')) {
        newPath = newLang === 'en' ? 'index-en.html' : 'index.html';
      }

      window.location.href = newPath;
    });
  }

  // ========================================
  // Active Navigation Link
  // ========================================
  const currentPage = window.location.pathname.split('/').pop() || 'index.html';

  document.querySelectorAll('.nav-link').forEach(link => {
    const href = link.getAttribute('href');
    if (href) {
      const linkPage = href.split('/').pop();
      if (linkPage === currentPage || (currentPage === 'index.html' && href === 'index.html')) {
        link.classList.add('active');
      }
    }
  });

  // ========================================
  // Print Function
  // ========================================
  window.printPage = function() {
    window.print();
  };

  // ========================================
  // Share Functions
  // ========================================
  window.shareOnFacebook = function() {
    const url = encodeURIComponent(window.location.href);
    window.open(`https://www.facebook.com/sharer/sharer.php?u=${url}`, '_blank', 'width=600,height=400');
  };

  window.shareOnTwitter = function() {
    const url = encodeURIComponent(window.location.href);
    const title = encodeURIComponent(document.title);
    window.open(`https://twitter.com/intent/tweet?url=${url}&text=${title}`, '_blank', 'width=600,height=400');
  };

  window.shareOnLine = function() {
    const url = encodeURIComponent(window.location.href);
    window.open(`https://social-plugins.line.me/lineit/share?url=${url}`, '_blank', 'width=600,height=400');
  };

  window.copyLink = function() {
    navigator.clipboard.writeText(window.location.href).then(() => {
      alert('連結已複製到剪貼簿！');
    });
  };

  // ========================================
  // Global Map Functions - MIT REAP Style
  // ========================================

  // Region Filter Functionality
  const regionSelect = document.getElementById('regionSelect');
  const personMarkers = document.querySelectorAll('.person-marker');

  if (regionSelect) {
    regionSelect.addEventListener('change', function() {
      const selectedRegion = this.value;

      personMarkers.forEach(marker => {
        const markerRegion = marker.getAttribute('data-region');

        if (selectedRegion === 'all' || markerRegion === selectedRegion) {
          marker.style.opacity = '1';
          marker.style.transform = 'translate(-50%, -100%) scale(1)';
          marker.style.pointerEvents = 'auto';
        } else {
          marker.style.opacity = '0.2';
          marker.style.transform = 'translate(-50%, -100%) scale(0.7)';
          marker.style.pointerEvents = 'none';
        }

        // Always keep Taiwan/Hualien visible
        if (marker.classList.contains('featured')) {
          marker.style.opacity = '1';
          marker.style.transform = 'translate(-50%, -100%) scale(1)';
          marker.style.pointerEvents = 'auto';
        }
      });
    });
  }

  // Hover effect for featured marker
  const featuredMarker = document.querySelector('.person-marker.featured');
  if (featuredMarker) {
    // Featured marker is always visible and highlighted
    featuredMarker.style.zIndex = '30';
  }

  // Handle top bar on scroll
  const topBar = document.querySelector('.header-top-bar');
  const headerWithTopBar = document.querySelector('.header.has-top-bar');

  if (topBar && headerWithTopBar) {
    window.addEventListener('scroll', function() {
      if (window.pageYOffset > 40) {
        topBar.style.transform = 'translateY(-100%)';
        headerWithTopBar.style.top = '0';
      } else {
        topBar.style.transform = 'translateY(0)';
        headerWithTopBar.style.top = '40px';
      }
    });
  }
});
