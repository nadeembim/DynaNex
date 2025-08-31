// DynaNex User Manual - Interactive JavaScript

document.addEventListener('DOMContentLoaded', function() {
    // Mobile Navigation Toggle
    const mobileMenu = document.getElementById('mobile-menu');
    const navMenu = document.querySelector('.nav-menu');
    
    if (mobileMenu) {
        mobileMenu.addEventListener('click', function() {
            mobileMenu.classList.toggle('active');
            navMenu.classList.toggle('active');
        });
    }

    // Fast Custom Smooth Scrolling for Navigation Links
    function fastSmoothScroll(target, duration = 800) {
        const targetPosition = target.offsetTop - 80; // Account for navbar
        const startPosition = window.pageYOffset;
        const distance = targetPosition - startPosition;
        let startTime = null;

        function animation(currentTime) {
            if (startTime === null) startTime = currentTime;
            const timeElapsed = currentTime - startTime;
            const run = ease(timeElapsed, startPosition, distance, duration);
            window.scrollTo(0, run);
            if (timeElapsed < duration) requestAnimationFrame(animation);
        }

        // Easing function for smooth animation (easeInOutCubic)
        function ease(t, b, c, d) {
            t /= d / 2;
            if (t < 1) return c / 2 * t * t * t + b;
            t -= 2;
            return c / 2 * (t * t * t + 2) + b;
        }

        requestAnimationFrame(animation);
    }

    // Apply fast smooth scrolling to all anchor links (nav + footer)
    document.querySelectorAll('a[href^="#"]').forEach(link => {
        link.addEventListener('click', function(e) {
            const href = this.getAttribute('href');
            
            if (href.startsWith('#')) {
                e.preventDefault();
                const targetSection = document.querySelector(href);
                
                if (targetSection) {
                    // Use custom fast smooth scroll
                    fastSmoothScroll(targetSection, 600); // Faster 600ms instead of default
                    
                    // Update active nav link (only for nav links)
                    if (this.classList.contains('nav-link')) {
                        document.querySelectorAll('.nav-link').forEach(navLink => {
                            navLink.classList.remove('active');
                        });
                        this.classList.add('active');
                    }
                    
                    // Close mobile menu if open
                    if (mobileMenu && mobileMenu.classList.contains('active')) {
                        mobileMenu.classList.remove('active');
                        navMenu.classList.remove('active');
                    }
                }
            }
        });
    });

    // Active Navigation Highlight on Scroll - Simple and fast
    const sections = document.querySelectorAll('section[id]');
    const navLinks = document.querySelectorAll('.nav-link[href^="#"]');

    function updateActiveNav() {
        const scrollPosition = window.scrollY + 100;
        
        sections.forEach(section => {
            const sectionTop = section.offsetTop;
            const sectionHeight = section.offsetHeight;
            const sectionId = section.getAttribute('id');
            
            if (scrollPosition >= sectionTop && scrollPosition < sectionTop + sectionHeight) {
                navLinks.forEach(link => {
                    link.classList.toggle('active', link.getAttribute('href') === `#${sectionId}`);
                });
            }
        });
    }

    // Simple scroll listener - works on all devices
    window.addEventListener('scroll', updateActiveNav, { passive: true });

    // Animate Elements on Scroll - Simple and fast
    const observerOptions = {
        threshold: 0.1,
        rootMargin: '0px 0px -50px 0px'
    };

    const observer = new IntersectionObserver(function(entries) {
        entries.forEach(entry => {
            if (entry.isIntersecting) {
                entry.target.style.opacity = '1';
                entry.target.style.transform = 'translateY(0)';
                observer.unobserve(entry.target);
            }
        });
    }, observerOptions);

    // Observe elements for animation - Smooth gliding effects
    document.querySelectorAll('.feature-card, .step, .doc-card').forEach(element => {
        element.style.opacity = '0';
        element.style.transform = 'translateY(40px)';
        element.style.transition = 'opacity 0.8s cubic-bezier(0.4, 0, 0.2, 1), transform 0.8s cubic-bezier(0.4, 0, 0.2, 1)';
        observer.observe(element);
    });

    // Hero Stats Counter Animation - Smooth counting
    function animateCounter(element, target, duration = 2000) {
        const start = 0;
        const startTime = performance.now();
        
        function updateCounter(currentTime) {
            const elapsed = currentTime - startTime;
            const progress = Math.min(elapsed / duration, 1);
            
            // Easing function for smooth animation
            const easeOutQuart = 1 - Math.pow(1 - progress, 4);
            
            if (target === '∞') {
                element.textContent = '∞';
            } else if (target.includes('%')) {
                const value = Math.floor(easeOutQuart * parseInt(target));
                element.textContent = value + '%';
            } else {
                const value = Math.floor(easeOutQuart * parseInt(target));
                element.textContent = value;
            }
            
            if (progress < 1) {
                requestAnimationFrame(updateCounter);
            } else {
                element.textContent = target;
            }
        }
        
        requestAnimationFrame(updateCounter);
    }

    // Trigger counter animation when hero section is visible
    const heroObserver = new IntersectionObserver(function(entries) {
        entries.forEach(entry => {
            if (entry.isIntersecting) {
                const statNumbers = entry.target.querySelectorAll('.stat-number');
                statNumbers.forEach(stat => {
                    const target = stat.getAttribute('data-target') || stat.textContent;
                    animateCounter(stat, target);
                });
                heroObserver.unobserve(entry.target);
            }
        });
    }, { threshold: 0.5 });

    const heroSection = document.querySelector('.hero');
    if (heroSection) {
        // Set data attributes for animation targets
        const statNumbers = heroSection.querySelectorAll('.stat-number');
        statNumbers.forEach(stat => {
            stat.setAttribute('data-target', stat.textContent);
            stat.textContent = '0';
        });
        
        heroObserver.observe(heroSection);
    }

    // Progress Bar Animation
    const progressObserver = new IntersectionObserver(function(entries) {
        entries.forEach(entry => {
            if (entry.isIntersecting) {
                const progressFill = entry.target.querySelector('.progress-fill');
                if (progressFill) {
                    setTimeout(() => {
                        progressFill.style.width = '85%';
                    }, 500);
                }
                progressObserver.unobserve(entry.target);
            }
        });
    }, { threshold: 0.5 });

    const floatingCard = document.querySelector('.floating-card');
    if (floatingCard) {
        progressObserver.observe(floatingCard);
    }

    // Add parallax effect to hero background - Simple and fast
    function updateParallax() {
        const scrolled = window.pageYOffset;
        const heroBackground = document.querySelector('.hero-background');
        
        if (heroBackground && scrolled < window.innerHeight) {
            heroBackground.style.transform = `translateY(${scrolled * -0.5}px)`;
        }
    }

    // Simple parallax - works on desktop and tablets
    if (window.innerWidth > 480) {
        window.addEventListener('scroll', updateParallax, { passive: true });
    }

    // Feature card hover effects - Only on non-touch devices
    if (!('ontouchstart' in window)) {
        document.querySelectorAll('.feature-card').forEach(card => {
            card.addEventListener('mouseenter', function() {
                this.style.transform = 'translateY(-8px) scale(1.02)';
            });
            
            card.addEventListener('mouseleave', function() {
                this.style.transform = 'translateY(0) scale(1)';
            });
        });

        // Doc card hover effects - Only on non-touch devices
        document.querySelectorAll('.doc-card').forEach(card => {
            card.addEventListener('mouseenter', function() {
                const arrow = this.querySelector('.doc-arrow');
                if (arrow) {
                    arrow.style.transform = 'translateX(5px)';
                }
            });
            
            card.addEventListener('mouseleave', function() {
                const arrow = this.querySelector('.doc-arrow');
                if (arrow) {
                    arrow.style.transform = 'translateX(0)';
                }
            });
        });
    }

    // Add loading animation
    window.addEventListener('load', function() {
        document.body.classList.add('loaded');
        
        // Animate hero content - Elegant entrance
        const heroContent = document.querySelector('.hero-content');
        if (heroContent) {
            heroContent.style.animation = 'fadeInUp 1.2s cubic-bezier(0.4, 0, 0.2, 1) ease-out';
        }
        
        const heroVisual = document.querySelector('.hero-visual');
        if (heroVisual) {
            heroVisual.style.animation = 'fadeInUp 1.2s cubic-bezier(0.4, 0, 0.2, 1) ease-out 0.3s both';
        }
    });

    // Add keyboard navigation support
    document.addEventListener('keydown', function(e) {
        if (e.key === 'Tab') {
            document.body.classList.add('keyboard-navigation');
        }
        // Close mobile menu on Escape key
        if (e.key === 'Escape' && mobileMenu && mobileMenu.classList.contains('active')) {
            mobileMenu.classList.remove('active');
            navMenu.classList.remove('active');
        }
    });

    document.addEventListener('click', function() {
        document.body.classList.remove('keyboard-navigation');
    });

    // Touch and mobile improvements
    let touchStartY = 0;
    let touchEndY = 0;

    // Better touch handling for mobile navigation
    if (mobileMenu && navMenu) {
        // Close menu when clicking outside
        document.addEventListener('click', function(e) {
            if (navMenu.classList.contains('active') && 
                !navMenu.contains(e.target) && 
                !mobileMenu.contains(e.target)) {
                mobileMenu.classList.remove('active');
                navMenu.classList.remove('active');
            }
        });

        // Touch gesture for closing menu
        document.addEventListener('touchstart', function(e) {
            touchStartY = e.changedTouches[0].screenY;
        });

        document.addEventListener('touchend', function(e) {
            touchEndY = e.changedTouches[0].screenY;
            handleSwipe();
        });

        function handleSwipe() {
            if (touchEndY < touchStartY - 50 && navMenu.classList.contains('active')) {
                // Swipe up to close menu
                mobileMenu.classList.remove('active');
                navMenu.classList.remove('active');
            }
        }
    }

    // Improve button touch feedback
    document.querySelectorAll('.btn').forEach(button => {
        button.addEventListener('touchstart', function() {
            this.style.transform = 'scale(0.98)';
        });

        button.addEventListener('touchend', function() {
            this.style.transform = 'scale(1)';
        });
    });

    // Prevent zoom on double-tap for better mobile UX
    let lastTouchEnd = 0;
    document.addEventListener('touchend', function(event) {
        const now = (new Date()).getTime();
        if (now - lastTouchEnd <= 300) {
            event.preventDefault();
        }
        lastTouchEnd = now;
    }, false);

    // Performance optimization: Lazy load images
    const images = document.querySelectorAll('img[data-src]');
    const imageObserver = new IntersectionObserver((entries, observer) => {
        entries.forEach(entry => {
            if (entry.isIntersecting) {
                const img = entry.target;
                img.src = img.dataset.src;
                img.classList.remove('lazy');
                imageObserver.unobserve(img);
            }
        });
    });

    images.forEach(img => imageObserver.observe(img));
});

// Old modal code removed - new simple modal is inline in HTML

// Add CSS animations via JavaScript
const style = document.createElement('style');
style.textContent = `
    @keyframes fadeInUp {
        from {
            opacity: 0;
            transform: translateY(30px);
        }
        to {
            opacity: 1;
            transform: translateY(0);
        }
    }

    .loaded .hero-content,
    .loaded .hero-visual {
        animation-fill-mode: both;
    }

    .nav-menu.active {
        display: flex;
        flex-direction: column;
        position: absolute;
        top: 100%;
        left: 0;
        right: 0;
        background: white;
        box-shadow: var(--shadow-lg);
        padding: var(--space-lg);
        gap: var(--space-md);
        border-radius: 0 0 var(--radius-lg) var(--radius-lg);
    }

    .nav-toggle.active .bar:nth-child(2) {
        opacity: 0;
    }

    .nav-toggle.active .bar:nth-child(1) {
        transform: translateY(8px) rotate(45deg);
    }

    .nav-toggle.active .bar:nth-child(3) {
        transform: translateY(-8px) rotate(-45deg);
    }

    .keyboard-navigation *:focus {
        outline: 2px solid var(--primary-blue);
        outline-offset: 2px;
    }

    @media (max-width: 768px) {
        .nav-menu {
            display: none;
        }
        
        .nav-menu.active {
            display: flex;
            flex-direction: column;
            position: absolute;
            top: 100%;
            left: 0;
            right: 0;
            background: rgba(0, 0, 0, 0.95);
            backdrop-filter: blur(20px);
            border-radius: 0 0 var(--radius-lg) var(--radius-lg);
            box-shadow: var(--shadow-xl);
            padding: var(--space-lg);
            gap: var(--space-lg);
            z-index: 1000;
            animation: slideDown 0.3s ease-out;
        }
        
        .nav-menu.active .nav-link {
            color: var(--primary-white);
            font-size: 1.1rem;
            padding: var(--space-md);
            border-radius: var(--radius-md);
            transition: all 0.2s ease;
        }
        
        .nav-menu.active .nav-link:hover,
        .nav-menu.active .nav-link.active {
            background: var(--accent-blue);
            transform: translateX(5px);
        }
        
        @keyframes slideDown {
            from {
                opacity: 0;
                transform: translateY(-10px);
            }
            to {
                opacity: 1;
                transform: translateY(0);
            }
        }
    }
`;
document.head.appendChild(style);