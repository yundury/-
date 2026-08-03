(function () {
  const heroCircle = document.getElementById("hero-circle");
  const loadingScreen = document.getElementById("loading-screen");
  const app = document.getElementById("app");
  const modal = document.getElementById("detail-modal");
  const modalBody = document.getElementById("modal-body");
  const modalClose = document.getElementById("modal-close");
  const modalBackdrop = modal.querySelector(".modal-backdrop");

  // 캐릭터 5명을 정오각형 모양으로 중앙 이미지 둘레에 배치
  const RADIUS_PERCENT = 38;

  function avatarPosition(index, total) {
    const angleDeg = -90 + index * (360 / total);
    const angleRad = (angleDeg * Math.PI) / 180;
    const left = 50 + RADIUS_PERCENT * Math.cos(angleRad);
    const top = 50 + RADIUS_PERCENT * Math.sin(angleRad);
    return { left, top };
  }

  // 이미지 로드 실패(아직 업로드 전 등) 시 자리표시로 자연스럽게 대체
  const ON_IMG_ERROR = `this.style.display='none'; this.nextElementSibling.hidden=false;`;

  function renderHeroCircle() {
    const centerHtml = APP_CONFIG.centerImage
      ? `<img src="${APP_CONFIG.centerImage}" alt="${APP_CONFIG.title}" onerror="${ON_IMG_ERROR}" />
         <div class="hero-center-placeholder" hidden>${APP_CONFIG.title}<br />(중앙 이미지 준비 중)</div>`
      : `<div class="hero-center-placeholder">${APP_CONFIG.title}<br />(중앙 이미지 준비 중)</div>`;

    const avatarsHtml = CHARACTERS.map((c, i) => {
      const { left, top } = avatarPosition(i, CHARACTERS.length);
      const inner = c.image
        ? `<img src="${c.image}" alt="${c.name}" onerror="${ON_IMG_ERROR}" /><span hidden>${c.emoji}</span>`
        : `<span>${c.emoji}</span>`;
      return `
        <button class="hero-avatar" data-id="${c.id}"
          style="left:${left}%; top:${top}%; background:${c.color}33; color:${c.color}"
          aria-label="${c.name}">
          ${inner}
        </button>`;
    }).join("");

    heroCircle.innerHTML = `
      <div class="hero-center">${centerHtml}</div>
      ${avatarsHtml}
    `;
  }

  function openDetail(id) {
    const c = CHARACTERS.find((item) => item.id === id);
    if (!c) return;

    modalBody.innerHTML = `
      <div class="modal-avatar" style="background:${c.color}33; color:${c.color}">
        <span>${c.emoji}</span>
      </div>
      <h2 class="modal-name">${c.name}</h2>
      <p class="modal-tagline">${c.tagline}</p>
      <div class="modal-meta">
        <span>${c.age}</span>
        <span>${c.role}</span>
      </div>
      <div class="modal-section-title">소개</div>
      <p class="modal-description">${c.description}</p>
      <div class="modal-section-title">특징</div>
      <ul class="modal-features">
        ${c.features.map((f) => `<li>${f}</li>`).join("")}
      </ul>
    `;

    modal.hidden = false;
    document.body.style.overflow = "hidden";
  }

  function closeDetail() {
    modal.hidden = true;
    document.body.style.overflow = "";
  }

  heroCircle.addEventListener("click", (e) => {
    const avatar = e.target.closest(".hero-avatar");
    if (!avatar) return;
    openDetail(Number(avatar.dataset.id));
  });

  modalClose.addEventListener("click", closeDetail);
  modalBackdrop.addEventListener("click", closeDetail);

  renderHeroCircle();

  // 로딩 화면을 잠깐 보여준 뒤 메인 화면 표시
  window.addEventListener("load", () => {
    setTimeout(() => {
      app.hidden = false;
      setTimeout(() => loadingScreen.remove(), 550);
    }, 1200);
  });

  // 서비스 워커 등록 (PWA 오프라인 지원)
  if ("serviceWorker" in navigator) {
    window.addEventListener("load", () => {
      navigator.serviceWorker.register("sw.js").catch(() => {});
    });
  }
})();
