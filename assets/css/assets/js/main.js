// Azra Property — Shared JS

// Header scroll
const hdr = document.getElementById('hdr');
if (hdr) {
  const solid = hdr.dataset.solid === 'true';
  if (!solid) {
    window.addEventListener('scroll', () => {
      hdr.classList.toggle('on', scrollY > 48);
    }, { passive: true });
  }
}

// Reveal on scroll
const obs = new IntersectionObserver(entries => {
  entries.forEach((e, i) => {
    if (e.isIntersecting) {
      setTimeout(() => e.target.classList.add('in'), i * 55);
      obs.unobserve(e.target);
    }
  });
}, { threshold: 0.07 });
document.querySelectorAll('.rv').forEach(el => obs.observe(el));

// Logo fallback
document.querySelectorAll('.logo-img').forEach(img => {
  img.addEventListener('error', () => {
    img.style.display = 'none';
    const fb = img.nextElementSibling;
    if (fb) fb.style.display = 'flex';
  });
});

// Active nav link
const path = window.location.pathname.split('/').pop() || 'index.html';
document.querySelectorAll('nav a').forEach(a => {
  const href = a.getAttribute('href');
  if (href === path || (path === 'index.html' && href === '/') || (path === '' && href === 'index.html')) {
    a.classList.add('active');
  }
});
