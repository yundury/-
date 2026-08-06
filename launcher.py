"""
config.py를 매번 열어서 고치지 않아도 되도록, 지역/키워드/브랜드/기준 리뷰수를
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


_PAGE_HTML = """<!doctype html>
<html lang="ko">
<head>
<meta charset="utf-8">
<title>네이버 맛집 리뷰 검색기</title>
<style>
  * { box-sizing: border-box; }
  body {
    margin: 0;
    background: #f7f7f5;
    color: #16181d;
    font-family: "Apple SD Gothic Neo", "Malgun Gothic", "Pretendard", sans-serif;
    padding: 36px 20px 60px;
  }
  .wrap { max-width: 640px; margin: 0 auto; }
  .card {
    background: #ffffff;
    border: 1px solid #e2e2de;
    border-radius: 14px;
    padding: 32px 36px;
    box-shadow: 0 18px 40px -24px rgba(0,0,0,0.18);
  }
  .title {
    font-size: 17px;
    font-weight: 800;
    margin-bottom: 22px;
  }
  .field { margin-bottom: 16px; }
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
    padding: 10px 12px;
    font-size: 13.5px;
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
  .actions {
    display: flex;
    justify-content: flex-end;
    gap: 8px;
    margin: 20px 0 4px;
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

  .log {
    margin-top: 18px;
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
  table { border-collapse: collapse; width: 100%; font-size: 12px; }
  th, td {
    padding: 9px 12px;
    text-align: left;
    border-bottom: 1px solid #eeeeec;
    white-space: nowrap;
  }
  th {
    position: sticky; top: 0;
    background: #f2f3f8;
    color: #33352f;
    font-weight: 700;
  }
  td.num { text-align: right; font-variant-numeric: tabular-nums; }
  td.wrap-cell { white-space: normal; max-width: 260px; }
  a.maplink { color: #3554d1; text-decoration: none; }
  a.maplink:hover { text-decoration: underline; }

  .bottom-actions {
    display: flex;
    justify-content: flex-end;
    gap: 8px;
    margin-top: 14px;
  }
  #openBtn { background: #eef1fb; color: #3554d1; }
  #quitBtn { background: #f2f2ef; color: #53565c; }
</style>
</head>
<body>
<div class="wrap">
  <div class="card">
    <div class="title">&#128465; 독수리오형제 Project</div>

    <div class="field">
      <label>희망지역</label>
      <input id="district" type="text">
      <div class="note">ex. 서울, 부산, 강남구 등</div>
    </div>
    <div class="field">
      <label>키워드</label>
      <input id="keyword" type="text">
      <div class="note">ex. 한식 맛집, 새로오픈한맛집 등</div>
    </div>
    <div class="field">
      <label>브랜드</label>
      <input id="brand" type="text">
      <div class="note">ex. 스타벅스, 교촌치킨 등 (특정 브랜드만 찾을 때)</div>
    </div>
    <div class="field">
      <label>기준리뷰수</label>
      <input id="minReviews" type="text">
      <div class="note">*기준 리뷰수 이상 리뷰가 달린 브랜드만 수집합니다. (비워두면 전부 수집)</div>
    </div>

    <div class="actions">
      <button id="stopBtn" disabled>중지</button>
      <button id="runBtn">실행</button>
    </div>
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
      <button id="quitBtn">종료</button>
      <button id="openBtn">Excel 파일 열기</button>
    </div>
  </div>
</div>

<script>
  const districtEl = document.getElementById("district");
  const keywordEl = document.getElementById("keyword");
  const brandEl = document.getElementById("brand");
  const minReviewsEl = document.getElementById("minReviews");
  const runBtn = document.getElementById("runBtn");
  const stopBtn = document.getElementById("stopBtn");
  const logEl = document.getElementById("log");
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

  runBtn.addEventListener("click", () => {
    const district = districtEl.value.trim();
    const keyword = keywordEl.value.trim();
    const brand = brandEl.value.trim();
    if (!district && !keyword && !brand) {
      alert("희망지역 / 키워드 / 브랜드 중 하나는 입력해주세요.");
      return;
    }
    shownDone = false;
    resultsEl.style.display = "none";
    logEl.textContent = "";

    fetch("/start", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        district: district,
        groups: keyword,
        brand: brand,
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

  quitBtn.addEventListener("click", () => {
    fetch("/shutdown", { method: "POST" }).finally(() => {
      document.body.innerHTML = "<div style='padding:60px;text-align:center;color:#8b8c86;font-family:sans-serif;'>프로그램을 종료했습니다. 이 탭은 닫으셔도 됩니다.</div>";
      clearInterval(polling);
    });
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
                self._send_json({"ok": False, "error": "희망지역 / 키워드 / 브랜드 중 하나는 입력해주세요."})
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
            threading.Thread(target=self.server.shutdown, daemon=True).start()

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
