const sections = document.querySelectorAll('.carousel-section');
let currentSection = 0;

function showSection(idx) {
  if (idx < 0) idx = sections.length - 1;
  if (idx >= sections.length) idx = 0;
  sections[currentSection].classList.remove('active');
  sections[idx].classList.add('active');
  currentSection = idx;
}

function prevSection() {
  showSection(currentSection - 1);
}

function nextSection() {
  showSection(currentSection + 1);
}

// Optional: Keyboard navigation
document.addEventListener('keydown', function(e) {
  if (e.key === 'ArrowLeft') prevSection();
  if (e.key === 'ArrowRight') nextSection();
});

// Archive dropdown logic
document.getElementById('archive-select').addEventListener('change', function() {
  var iframe = document.getElementById('archive-iframe');
  if (this.value) {
    iframe.src = this.value;
    iframe.style.display = 'block';
  } else {
    iframe.src = '';
    iframe.style.display = 'none';
  }
});

// Readings dropdown logic
const readingsSelect = document.getElementById('readings-select');
if (readingsSelect) {
  readingsSelect.addEventListener('change', function() {
    const iframe = document.getElementById('readings-iframe');
    if (this.value) {
      iframe.src = this.value;
      iframe.style.display = 'block';
    } else {
      iframe.src = '';
      iframe.style.display = 'none';
    }
  });
}

document.getElementById('expand-section-btn').addEventListener('click', function() {
  const carousel = document.querySelector('.carousel-container');
  carousel.classList.add('expanded-mode');
  this.classList.add('expanded-hide');

  // Add close button if not present
  if (!document.getElementById('close-expanded-btn')) {
    const btn = document.createElement('button');
    btn.id = 'close-expanded-btn';
    btn.innerHTML = '&times;';
    btn.title = 'Close expanded view';
    btn.className = 'fullscreen-close-btn';
    btn.onclick = function () {
      carousel.classList.remove('expanded-mode');
      document.getElementById('expand-section-btn').classList.remove('expanded-hide');
      btn.remove();
    };
    carousel.appendChild(btn);
  }
});

function showContextsSubpage(which) {
  const subpages = document.querySelectorAll('.contexts-subpage');
  subpages.forEach(div => div.style.display = 'none');
  const active = document.getElementById('contexts-' + which);
  if (active) active.style.display = 'block';
}

function showAdvocacySubpage(which) {
  const subpages = document.querySelectorAll('.advocacy-subpage');
  subpages.forEach(div => div.style.display = 'none');
  const active = document.getElementById('advocacy-' + which);
  if (active) active.style.display = 'block';
}

