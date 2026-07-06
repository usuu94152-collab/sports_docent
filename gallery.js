/* 카드뉴스 라이트박스 + 좌우 이동 버튼 — 공통 스크립트 */
(function () {
  var figs = Array.prototype.slice.call(document.querySelectorAll('.slides figure'));
  if (!figs.length) return;

  /* ---- 좌우 이동 버튼 ---- */
  var slides = document.querySelector('.slides');
  if (slides) {
    var wrap = document.createElement('div');
    wrap.className = 'slides-wrap';
    slides.parentNode.insertBefore(wrap, slides);
    wrap.appendChild(slides);

    var prevBtn = document.createElement('button');
    prevBtn.className = 'slide-nav prev';
    prevBtn.setAttribute('aria-label', '이전 이미지');
    prevBtn.innerHTML = '‹';
    var nextBtn = document.createElement('button');
    nextBtn.className = 'slide-nav next';
    nextBtn.setAttribute('aria-label', '다음 이미지');
    nextBtn.innerHTML = '›';
    wrap.appendChild(prevBtn);
    wrap.appendChild(nextBtn);

    function step(dir) {
      var fig = slides.querySelector('figure');
      var amount = fig ? fig.clientWidth + 12 : slides.clientWidth;
      slides.scrollBy({ left: dir * amount, behavior: 'smooth' });
    }
    function refresh() {
      prevBtn.classList.toggle('is-off', slides.scrollLeft <= 2);
      nextBtn.classList.toggle('is-off',
        slides.scrollLeft + slides.clientWidth >= slides.scrollWidth - 2);
    }
    prevBtn.addEventListener('click', function () { step(-1); });
    nextBtn.addEventListener('click', function () { step(1); });
    slides.addEventListener('scroll', refresh);
    window.addEventListener('resize', refresh);
    refresh();
  }

  var srcs = figs.map(function (f) {
    var img = f.querySelector('img');
    return img ? img.getAttribute('src') : '';
  });

  var box = document.createElement('div');
  box.className = 'lightbox';
  box.innerHTML =
    '<button class="close" aria-label="닫기">&times;</button>' +
    '<button class="nav prev" aria-label="이전">&#8249;</button>' +
    '<img alt="확대 이미지">' +
    '<button class="nav next" aria-label="다음">&#8250;</button>' +
    '<div class="counter"></div>';
  document.body.appendChild(box);

  var imgEl = box.querySelector('img');
  var counter = box.querySelector('.counter');
  var cur = 0;

  function show(i) {
    cur = (i + srcs.length) % srcs.length;
    imgEl.src = srcs[cur];
    counter.textContent = (cur + 1) + ' / ' + srcs.length;
  }
  function open(i) { show(i); box.classList.add('open'); }
  function close() { box.classList.remove('open'); }

  figs.forEach(function (f, i) {
    f.addEventListener('click', function () { open(i); });
  });
  box.querySelector('.close').addEventListener('click', close);
  box.querySelector('.prev').addEventListener('click', function (e) { e.stopPropagation(); show(cur - 1); });
  box.querySelector('.next').addEventListener('click', function (e) { e.stopPropagation(); show(cur + 1); });
  box.addEventListener('click', function (e) { if (e.target === box) close(); });
  document.addEventListener('keydown', function (e) {
    if (!box.classList.contains('open')) return;
    if (e.key === 'Escape') close();
    else if (e.key === 'ArrowLeft') show(cur - 1);
    else if (e.key === 'ArrowRight') show(cur + 1);
  });
})();
