(function () {
  const grid = document.getElementById("character-grid");
  const loadingScreen = document.getElementById("loading-screen");
  const app = document.getElementById("app");
  const modal = document.getElementById("detail-modal");
  const modalBody = document.getElementById("modal-body");
  const modalClose = document.getElementById("modal-close");
  const modalBackdrop = modal.querySelector(".modal-backdrop");

  function renderGrid() {
    grid.innerHTML = CHARACTERS.map(
      (c) => `
      <button class="character-card" data-id="${c.id}">
        <div class="character-avatar" style="background:${c.color}33; color:${c.color}">
          <span>${c.emoji}</span>
        </div>
        <div class="character-name">${c.name}</div>
        <div class="character-tagline">${c.tagline}</div>
      </button>`
    ).join("");
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

  grid.addEventListener("click", (e) => {
    const card = e.target.closest(".character-card");
    if (!card) return;
    openDetail(Number(card.dataset.id));
  });

  modalClose.addEventListener("click", closeDetail);
  modalBackdrop.addEventListener("click", closeDetail);

  renderGrid();

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
