"""
config.py를 매번 열어서 고치지 않아도 되도록, 지역/검색키워드/필터/기준 리뷰수를
입력하는 화면을 로컬 웹페이지로 띄워주는 실행기입니다. 뒤에서는
naver_restaurant_scraper.py의 크롤링 코드가 파이썬으로 그대로 돌아가고,
화면만 사용자님 컴퓨터의 기본 브라우저에 새 탭으로 예쁘게 뜹니다.

실행 방법: python launcher.py (또는 start.bat 더블클릭)
"""

import importlib
import http.server
import json
import os
import socketserver
import subprocess
import sys
import threading
import time
import webbrowser

import config
import naver_restaurant_scraper as scraper

PORT = 8899

_state_lock = threading.Lock()
_state = {
    "running": False,
    "log": "",
    "done": False,
    "rows": [],
    "output_path": None,
    "error": None,
}
_stop_event = None


class _LogWriter:
    """print() 출력을 웹페이지 로그창으로 보내기 위한 도우미."""

    def write(self, text):
        if text:
            with _state_lock:
                _state["log"] += text
        return len(text)

    def flush(self):
        pass


def _build_output_name(district, groups, brand, min_reviews):
    name_parts = [p.replace("/", "_") for p in (district, "_".join(groups), brand) if p]
    base_name = "_".join(name_parts) if name_parts else "검색결과"
    return f"{base_name}_리뷰{min_reviews}개이상.xlsx"


def _run_scraper(district, groups_text, brand, min_reviews):
    global _stop_event

    # config.py를 다시 읽어온다 - 이 서버는 여러 번 "실행"을 눌러도 계속 켜져
    # 있으므로, 중간에 config.py를 고쳤다면 이번 실행부터는 반영해야 한다.
    importlib.reload(config)
    config.DISTRICT = district
    config.PRODUCT_GROUPS = [g.strip() for g in groups_text.replace(",", "\n").split("\n") if g.strip()]
    config.BRAND = brand
    config.MIN_REVIEW_COUNT = min_reviews
    config.OUTPUT_FILE = _build_output_name(district, config.PRODUCT_GROUPS, brand, min_reviews)

    old_stdout = sys.stdout
    sys.stdout = _LogWriter()
    try:
        output_path, rows = scraper.main(stop_event=_stop_event)
        with _state_lock:
            _state["output_path"] = output_path
            _state["rows"] = rows
    except Exception as e:
        with _state_lock:
            _state["error"] = str(e)
        print(f"\n오류가 발생했습니다: {e}\n")
    finally:
        sys.stdout = old_stdout
        with _state_lock:
            _state["running"] = False
            _state["done"] = True


def _open_file(path):
    if sys.platform.startswith("win"):
        os.startfile(path)
    elif sys.platform == "darwin":
        subprocess.Popen(["open", path])
    else:
        subprocess.Popen(["xdg-open", path])


def _graceful_shutdown(server):
    """탭을 닫거나 종료 버튼을 눌렀을 때 프로그램을 끈다.

    한창 검색 중이었다면 무작정 끄지 않고, "중지" 버튼을 누른 것과 똑같이
    지금까지 모은 것만 저장하고 끝날 때까지 잠깐 기다린 뒤에 끈다 (그래야
    실수로 탭을 닫아도 진행 중이던 결과를 잃지 않는다).
    """
    with _state_lock:
        running = _state["running"]
    if running and _stop_event is not None:
        _stop_event.set()
        for _ in range(1200):  # 최대 약 2분 대기
            with _state_lock:
                if not _state["running"]:
                    break
            time.sleep(0.1)
    server.shutdown()


_PAGE_HTML = """<!doctype html>
<html lang="ko">
<head>
<meta charset="utf-8">
<title>Up Project 08. AX</title>
<style>
  * { box-sizing: border-box; }
  body {
    margin: 0;
    background: #f7f7f5;
    color: #16181d;
    font-family: "Apple SD Gothic Neo", "Malgun Gothic", "Pretendard", sans-serif;
    padding: 36px 20px 60px;
  }
  .wrap { max-width: 1080px; margin: 0 auto; }
  .card {
    background: #ffffff;
    border: 1px solid #e2e2de;
    border-radius: 14px;
    padding: 36px 44px;
    box-shadow: 0 18px 40px -24px rgba(0,0,0,0.18);
  }
  .title {
    font-size: 20px;
    font-weight: 800;
    margin-bottom: 4px;
  }
  .credit {
    font-size: 11px;
    color: #a3a49d;
    margin-bottom: 10px;
  }
  .subtitle {
    font-size: 13px;
    color: #8b8c86;
    margin-bottom: 26px;
  }
  .fields-grid {
    display: grid;
    grid-template-columns: 1fr 1fr;
    column-gap: 36px;
    row-gap: 22px;
  }
  .field label {
    display: block;
    font-size: 12.5px;
    font-weight: 700;
    color: #53565c;
    margin-bottom: 7px;
  }
  .field input {
    width: 100%;
    border: 1px solid #e2e2de;
    border-radius: 10px;
    padding: 11px 14px;
    font-size: 14px;
    color: #16181d;
    font-family: inherit;
  }
  .field input:focus {
    outline: none;
    border-color: #3554d1;
  }
  .field .note {
    font-size: 11.5px;
    color: #8b8c86;
    margin-top: 6px;
  }
  .field-wide { grid-column: 1 / -1; }
  .filter-pills {
    display: flex;
    flex-wrap: wrap;
    gap: 8px;
  }
  .pill {
    border: 1px solid #dbe1f7;
    background: #eef1fb;
    color: #3554d1;
    border-radius: 999px;
    padding: 8px 16px;
    font-size: 12.5px;
    font-weight: 700;
    cursor: pointer;
    font-family: inherit;
    transition: background 0.12s, color 0.12s, border-color 0.12s;
  }
  .pill:hover { background: #e2e7fa; }
  .pill.active {
    background: #3554d1;
    border-color: #3554d1;
    color: #ffffff;
  }
  .actions {
    display: flex;
    justify-content: flex-end;
    gap: 8px;
    margin: 28px 0 4px;
  }
  button {
    border: none;
    border-radius: 999px;
    padding: 10px 22px;
    font-size: 13px;
    font-weight: 700;
    cursor: pointer;
    font-family: inherit;
  }
  button:disabled { cursor: default; opacity: 0.55; }
  #stopBtn { background: #f7e8e5; color: #a33f2b; }
  #runBtn { background: #3554d1; color: #ffffff; }
  #runBtn:hover:not(:disabled) { background: #2c46b3; }
  #stopBtn:hover:not(:disabled) { background: #f0dbd6; }

  .log-head {
    display: flex;
    justify-content: space-between;
    align-items: center;
    margin-top: 18px;
    margin-bottom: 6px;
  }
  .log-head .log-title { font-size: 12.5px; font-weight: 700; color: #53565c; }
  #copyLogBtn {
    background: #f2f2ef;
    color: #53565c;
    font-size: 11.5px;
    padding: 5px 12px;
  }
  #copyLogBtn:hover { background: #e7e7e3; }

  .log {
    background: #ffffff;
    border: 1px solid #e2e2de;
    border-radius: 12px;
    padding: 14px 16px;
    font-size: 12px;
    line-height: 1.75;
    color: #4b4c48;
    white-space: pre-wrap;
    height: 220px;
    overflow-y: auto;
  }

  .results {
    margin-top: 18px;
    display: none;
  }
  .results-head {
    display: flex;
    align-items: baseline;
    justify-content: space-between;
    margin-bottom: 10px;
  }
  .results-head .count { font-size: 13.5px; font-weight: 700; }
  .results-head .path { font-size: 11.5px; color: #8b8c86; }
  .table-wrap {
    border: 1px solid #e2e2de;
    border-radius: 12px;
    overflow: auto;
    max-height: 360px;
    background: #ffffff;
  }
  table { border-collapse: collapse; width: 100%; table-layout: fixed; font-size: 12px; }
  th, td {
    padding: 9px 12px;
    text-align: left;
    border-bottom: 1px solid #eeeeec;
    white-space: nowrap;
    overflow: hidden;
    text-overflow: ellipsis;
  }
  th {
    position: sticky; top: 0;
    background: #f2f3f8;
    color: #33352f;
    font-weight: 700;
  }
  th:nth-child(1), td:nth-child(1) { width: 15%; }
  th:nth-child(2), td:nth-child(2) { width: 9%; }
  th:nth-child(3), td:nth-child(3) { width: 10%; }
  th:nth-child(4), td:nth-child(4) { width: 7%; }
  th:nth-child(5), td:nth-child(5) { width: 49%; }
  th:nth-child(6), td:nth-child(6) { width: 10%; }
  td.num { text-align: right; font-variant-numeric: tabular-nums; }
  td.wrap-cell { white-space: normal; overflow: visible; text-overflow: clip; line-height: 1.6; }
  a.maplink { color: #3554d1; text-decoration: none; }
  a.maplink:hover { text-decoration: underline; }

  .bottom-actions {
    display: flex;
    justify-content: flex-end;
    gap: 8px;
    margin-top: 14px;
  }
  #openBtn { background: #eef1fb; color: #3554d1; }

  .topbar {
    display: flex;
    justify-content: flex-end;
    margin-bottom: 10px;
  }
  #quitBtn {
    background: transparent;
    color: #8b8c86;
    font-size: 11.5px;
    font-weight: 600;
    padding: 4px 8px;
    border-radius: 6px;
  }
  #quitBtn:hover { background: #ebebe8; color: #53565c; }
</style>
</head>
<body>
<div class="wrap">
  <div class="topbar">
    <button id="quitBtn">프로그램 종료</button>
  </div>
  <div class="card">
    <div class="title">Up Project 08. AX</div>
    <div class="credit">만든이: 김채윤, 김윤경, 정민정, 김선태, 홍지우 바이어</div>
    <div class="subtitle">희망지역 · 검색키워드 · 필터를 조합해서 네이버 지도에서 리뷰 많은 곳을 찾아드려요.</div>

    <div class="fields-grid">
      <div class="field">
        <label>희망지역</label>
        <input id="district" type="text">
        <div class="note">ex. 서울, 부산, 강남구 등</div>
      </div>
      <div class="field">
        <label>검색키워드</label>
        <input id="searchKeyword" type="text">
        <div class="note">*(선택)상품군/브랜드 등 원하는 내용을 입력할 수 있습니다. ex. 평양냉면, 파스타, 한정식 등</div>
      </div>
      <div class="field field-wide">
        <label>필터</label>
        <div class="filter-pills">
          <button type="button" class="pill" data-value="미쉐린">미쉐린</button>
          <button type="button" class="pill" data-value="새로오픈한 맛집">새로오픈한 맛집</button>
          <button type="button" class="pill" data-value="많이찾는 맛집">많이찾는 맛집</button>
          <button type="button" class="pill" data-value="리뷰많은 맛집">리뷰많은 맛집</button>
          <button type="button" class="pill" data-value="현대백화점">현대백화점</button>
          <button type="button" class="pill" data-value="신세계백화점">신세계백화점</button>
        </div>
        <div class="note">*(선택)네이버 지도 검색시 사용되는 필터입니다. 선택시 해당 필터가 적용된 결과값이 도출됩니다.</div>
      </div>
      <div class="field">
        <label>기준리뷰수</label>
        <input id="minReviews" type="text">
        <div class="note">*기준 리뷰수 이상 리뷰가 달린 브랜드만 수집합니다. (비워두면 전부 수집)</div>
      </div>
    </div>

    <div class="actions">
      <button id="stopBtn" disabled>중지</button>
      <button id="runBtn">실행</button>
    </div>
  </div>

  <div class="log-head">
    <div class="log-title">진행 로그</div>
    <button id="copyLogBtn">로그 복사</button>
  </div>
  <div id="log" class="log"></div>

  <div id="results" class="results">
    <div class="results-head">
      <div class="count" id="resultCount"></div>
      <div class="path" id="resultPath"></div>
    </div>
    <div class="table-wrap">
      <table>
        <thead>
          <tr>
            <th>브랜드명</th><th>키워드</th><th>카테고리</th><th>리뷰수</th><th>대표 리뷰</th><th>네이버지도주소</th>
          </tr>
        </thead>
        <tbody id="resultBody"></tbody>
      </table>
    </div>
    <div class="bottom-actions">
      <button id="openBtn">Excel 파일 열기</button>
    </div>
  </div>
</div>

<script>
  const districtEl = document.getElementById("district");
  const searchKeywordEl = document.getElementById("searchKeyword");
  const minReviewsEl = document.getElementById("minReviews");
  const pillEls = document.querySelectorAll(".pill");
  const runBtn = document.getElementById("runBtn");
  const stopBtn = document.getElementById("stopBtn");
  const logEl = document.getElementById("log");
  const copyLogBtn = document.getElementById("copyLogBtn");
  const resultsEl = document.getElementById("results");
  const resultCountEl = document.getElementById("resultCount");
  const resultPathEl = document.getElementById("resultPath");
  const resultBodyEl = document.getElementById("resultBody");
  const openBtn = document.getElementById("openBtn");
  const quitBtn = document.getElementById("quitBtn");

  let polling = null;
  let shownDone = false;

  function escapeHtml(s) {
    return String(s).replace(/[&<>"']/g, c => ({
      "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;"
    }[c]));
  }

  function renderResults(data) {
    resultsEl.style.display = "block";
    resultCountEl.textContent = `결과: ${data.rows.length}곳`;
    resultPathEl.textContent = data.output_path ? data.output_path : "";
    resultBodyEl.innerHTML = data.rows.map(r => `
      <tr>
        <td>${escapeHtml(r["브랜드명"] || "")}</td>
        <td>${escapeHtml(r["키워드"] || "")}</td>
        <td>${escapeHtml(r["카테고리"] || "")}</td>
        <td class="num">${escapeHtml(r["리뷰수"] || 0)}</td>
        <td class="wrap-cell">${escapeHtml(r["대표 리뷰"] || "")}</td>
        <td>${r["네이버지도주소"] ? `<a class="maplink" href="${escapeHtml(r["네이버지도주소"])}" target="_blank" rel="noopener">지도 열기</a>` : ""}</td>
      </tr>
    `).join("");
  }

  function poll() {
    fetch("/status").then(r => r.json()).then(data => {
      logEl.textContent = data.log;
      logEl.scrollTop = logEl.scrollHeight;

      runBtn.disabled = data.running;
      runBtn.textContent = data.running ? "실행 중..." : "실행";
      stopBtn.disabled = !data.running;

      if (data.done && !data.running && !shownDone) {
        shownDone = true;
        renderResults(data);
      }
    }).catch(() => {});
  }

  pillEls.forEach(btn => {
    btn.addEventListener("click", () => btn.classList.toggle("active"));
  });

  runBtn.addEventListener("click", () => {
    const district = districtEl.value.trim();
    const searchKeyword = searchKeywordEl.value.trim();
    const filters = Array.from(document.querySelectorAll(".pill.active")).map(b => b.dataset.value);
    if (!district && filters.length === 0 && !searchKeyword) {
      alert("희망지역 / 필터 / 검색키워드 중 하나는 입력해주세요.");
      return;
    }
    shownDone = false;
    resultsEl.style.display = "none";
    logEl.textContent = "";

    // 필터와 검색키워드는 따로따로 검색하는 게 아니라, 네이버 지도에 실제로
    // 입력하듯이 하나로 합친 검색어로 검색한다 (예: "서울 미쉐린 파스타").
    const combinedKeyword = [...filters, searchKeyword].filter(Boolean).join(" ");

    fetch("/start", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        district: district,
        groups: combinedKeyword,
        brand: "",
        min_reviews: minReviewsEl.value.trim(),
      }),
    }).then(r => r.json()).then(data => {
      if (!data.ok) alert(data.error || "시작하지 못했습니다.");
    });
  });

  stopBtn.addEventListener("click", () => {
    fetch("/stop", { method: "POST" });
  });

  openBtn.addEventListener("click", () => {
    fetch("/open-file", { method: "POST" });
  });

  copyLogBtn.addEventListener("click", () => {
    navigator.clipboard.writeText(logEl.textContent).then(() => {
      const original = copyLogBtn.textContent;
      copyLogBtn.textContent = "복사됨!";
      setTimeout(() => { copyLogBtn.textContent = original; }, 1200);
    }).catch(() => {
      alert("복사에 실패했습니다. 로그 내용을 직접 드래그해서 복사해주세요.");
    });
  });

  let quitting = false;

  quitBtn.addEventListener("click", () => {
    quitting = true;
    fetch("/shutdown", { method: "POST" }).finally(() => {
      document.body.innerHTML = "<div style='padding:60px;text-align:center;color:#8b8c86;font-family:sans-serif;'>프로그램을 종료했습니다. 이 탭은 닫으셔도 됩니다.</div>";
      clearInterval(polling);
    });
  });

  // 탭/창을 그냥 닫아도 자동으로 프로그램이 꺼지도록 한다. sendBeacon은 페이지가
  // 닫히는 도중에도 요청이 확실히 전달되도록 브라우저가 보장해주는 방식이다.
  // (검색 중이었다면 서버 쪽에서 "중지"와 똑같이 지금까지 모은 것만 저장하고 끈다.)
  window.addEventListener("pagehide", () => {
    if (!quitting) {
      navigator.sendBeacon("/shutdown");
    }
  });

  polling = setInterval(poll, 1000);
  poll();
</script>
</body>
</html>
"""


class _Handler(http.server.BaseHTTPRequestHandler):
    def log_message(self, format, *args):
        pass

    def _send_json(self, obj, status=200):
        body = json.dumps(obj, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def _read_json_body(self):
        length = int(self.headers.get("Content-Length", 0) or 0)
        raw = self.rfile.read(length) if length else b""
        try:
            return json.loads(raw.decode("utf-8")) if raw else {}
        except Exception:
            return {}

    def do_GET(self):
        if self.path == "/" or self.path.startswith("/?"):
            body = _PAGE_HTML.encode("utf-8")
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            self.wfile.write(body)
        elif self.path == "/status":
            with _state_lock:
                snapshot = dict(_state)
            self._send_json(snapshot)
        else:
            self.send_response(404)
            self.end_headers()

    def do_POST(self):
        global _stop_event
        payload = self._read_json_body()

        if self.path == "/start":
            with _state_lock:
                already_running = _state["running"]
            if already_running:
                self._send_json({"ok": False, "error": "이미 실행 중입니다."})
                return

            district = (payload.get("district") or "").strip()
            groups_text = (payload.get("groups") or "").strip()
            brand = (payload.get("brand") or "").strip()
            min_reviews_raw = (payload.get("min_reviews") or "").strip()
            if not district and not groups_text and not brand:
                self._send_json({"ok": False, "error": "희망지역 / 필터 / 검색키워드 중 하나는 입력해주세요."})
                return
            try:
                min_reviews = int(min_reviews_raw.replace(",", "")) if min_reviews_raw else 0
            except ValueError:
                self._send_json({"ok": False, "error": "기준 리뷰 수는 숫자로 입력해주세요."})
                return

            with _state_lock:
                _state.update({"running": True, "log": "", "done": False, "rows": [], "output_path": None, "error": None})
            _stop_event = threading.Event()
            thread = threading.Thread(target=_run_scraper, args=(district, groups_text, brand, min_reviews), daemon=True)
            thread.start()
            self._send_json({"ok": True})

        elif self.path == "/stop":
            if _stop_event is not None:
                _stop_event.set()
            self._send_json({"ok": True})

        elif self.path == "/open-file":
            with _state_lock:
                path = _state.get("output_path")
            if path and os.path.exists(path):
                try:
                    _open_file(path)
                except Exception:
                    pass
            self._send_json({"ok": True})

        elif self.path == "/shutdown":
            self._send_json({"ok": True})
            threading.Thread(target=_graceful_shutdown, args=(self.server,), daemon=True).start()

        else:
            self.send_response(404)
            self.end_headers()


def main():
    url = f"http://127.0.0.1:{PORT}/"
    try:
        httpd = socketserver.TCPServer(("127.0.0.1", PORT), _Handler)
    except OSError:
        # 이미 이 프로그램이 실행 중인 것으로 보고, 새로 띄우지 않고 브라우저 탭만 연다.
        webbrowser.open(url)
        return

    webbrowser.open(url)
    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        pass


if __name__ == "__main__":
    main()
