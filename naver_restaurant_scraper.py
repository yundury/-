"""
네이버 지도에서 지역+음식종류로 검색한 뒤, 검색 결과 목록을 스크롤해서 모으고
(목록에 이미 나오는 리뷰 수로 1차로 걸러낸 다음), 살아남은 가게만 상세 페이지에
들어가서 주소/카테고리/AI 브리핑/지도 링크까지 채워 엑셀로 저장하는 스크립트.

네이버 지역검색 API는 더 이상 쓰지 않는다 (검색어당 5개 제한이 있고, 어차피
목록 화면에 리뷰 수가 그대로 보이기 때문). API 키도 필요 없다.

실행 방법은 README.md를 참고하세요.
모든 설정값(지역, 찾을 음식 종류, 최소 리뷰 수 등)은 config.py에서 바꿉니다.
"""

import os
import re
import time
import random
from urllib.parse import quote

from playwright.sync_api import sync_playwright
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

import config


# 브라우저 로그인/쿠키 정보를 저장해두는 폴더.
# 매번 새 브라우저(신규 방문자)인 것처럼 접속하면 네이버가 자동화로 의심하기 쉬워서,
# 같은 브라우저 프로필을 계속 재사용해 "이전에도 왔던 사람"처럼 보이게 한다.
PROFILE_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "browser_profile")

# 최종 엑셀 컬럼 순서
HEADERS = ["가게이름", "상품군", "카테고리", "주소", "네이버지도 주소", "리뷰수", "브랜드설명"]

# 엑셀 서식에 쓸 값들
FONT_NAME = "나눔바른고딕"
HEADER_FILL = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")


def _district_only(address):
    """'서울 강남구 도산대로1길 10' -> '서울 강남구'처럼 구까지만 남긴다."""
    parts = address.split()
    return " ".join(parts[:2]) if len(parts) >= 2 else address


def _parse_review_count(text):
    """'리뷰 1.1만', '리뷰 1,234' 같은 표기에서 숫자를 뽑아낸다.

    네이버 플레이스는 리뷰가 많으면 '1.1만'처럼 만/천 단위로 줄여서 보여준다.
    """
    match = re.search(r"리뷰\s*([\d,]+(?:\.\d+)?)\s*(만|천)?", text)
    if not match:
        return None

    value = float(match.group(1).replace(",", ""))
    unit = match.group(2)
    if unit == "만":
        value *= 10000
    elif unit == "천":
        value *= 1000

    return int(value)


# 가게 이름 뒤에 공백 없이 바로 붙어 나올 수 있는 '예약/톡톡/쿠폰' 같은 배지 글자들.
# 목록 항목 첫 줄에서 이런 배지가 이름에 섞여 나오면 뒤에서부터 잘라낸다.
_KNOWN_BADGE_WORDS = ["예약", "톡톡", "쿠폰", "포장주문", "발견"]


def _strip_trailing_badges(name):
    changed = True
    while changed:
        changed = False
        for badge in _KNOWN_BADGE_WORDS:
            if name.endswith(badge):
                name = name[: -len(badge)]
                changed = True
    return name.strip()


def _parse_list_item(text):
    """검색 결과 '목록'에 나오는 항목 하나의 텍스트에서 (가게이름, 리뷰수)를 뽑아낸다.

    목록의 각 항목은 보통 첫 줄이 가게 이름이고, 어딘가에 '리뷰 N' 형태로
    리뷰 수가 나온다. 리뷰 수가 없는 항목(광고/필터 칩 등)은 걸러진다.
    """
    review_count = _parse_review_count(text)
    if review_count is None:
        return None

    first_line = _strip_trailing_badges(text.strip().split("\n")[0].strip())
    if not first_line:
        return None

    return first_line, review_count


def _extract_name_and_category(body_text):
    """상세 페이지 글자에서 '별점' 바로 앞 두 줄(가게이름, 카테고리)을 뽑아낸다.

    네이버 플레이스 상세 화면은 보통 '...\\n가게이름\\n카테고리\\n별점\\n4.85리뷰...'
    순서로 나온다.
    """
    idx = body_text.find("별점")
    if idx == -1:
        return "", ""

    lines = [l.strip() for l in body_text[:idx].split("\n") if l.strip()]
    if not lines:
        return "", ""

    category = lines[-1]
    name = lines[-2] if len(lines) >= 2 else ""
    return name, category


def _extract_address(body_text):
    """상세 페이지 글자에서 '주소' 라벨 바로 다음 줄을 주소로 가져온다."""
    match = re.search(r"주소\n([^\n]+)", body_text)
    return match.group(1).strip() if match else ""


def _parse_ai_briefing(text, max_items=2):
    """'AI 브리핑' 라벨 아래에 나오는 요약 문장을 최대 max_items개까지 가져온다.

    각 요약 문장 끝에는 보통 '닉네임 +2' 같은 출처 표시가 붙는데,
    그 표시를 기준으로 문장 단위를 나눈다.
    """
    match = re.search(r"AI\s*브리핑", text)
    if not match:
        return None, "'AI 브리핑' 글자를 페이지에서 못 찾았습니다 (스크롤해도 안 나타났을 수 있음)."

    section = text[match.end():match.end() + 1500]
    debug_context = f"'AI 브리핑' 뒤 1500자: {section!r}"

    # '실험 단계로 정확하지 않을 수 있어요' / '~정리한 정보는 다음과 같습니다' 같은
    # 고정 안내 문구는 실제 요약 내용이 아니므로 제거하고 시작한다.
    section = re.sub(r".*?정리한 정보는 다음과 같습니다\.?", "", section, count=1, flags=re.S)

    # 각 요약 문장은 보통 끝에 '닉네임 +숫자' 형태의 출처 표시가 붙는다.
    bullets = re.findall(r"(.+?)\s*\S+\s*\+\d+", section, flags=re.S)
    bullets = [b.replace("\n", " ").strip() for b in bullets if b.strip()]

    if not bullets:
        # 출처 표시를 못 찾으면, 줄바꿈 기준으로 대충이라도 나눠본다
        bullets = [b.strip() for b in section.split("\n") if b.strip()]

    if not bullets:
        return None, debug_context

    return " / ".join(bullets[:max_items]), debug_context


def _wait_for_list_or_entry(page, timeout_ms=15000, poll_ms=300):
    """검색 후 '목록'이 뜨는지, 검색결과가 하나뿐이라 '상세 페이지'로 바로
    넘어가는지는 미리 알 수 없다. 고정 시간을 기다리는 대신, 둘 중 무엇이
    먼저 나타나는지 짧은 간격으로 계속 확인해서 그때그때 판단한다.

    반환값: "list" | "entry" | "timeout"
    """
    elapsed = 0
    while elapsed < timeout_ms:
        if page.frame_locator("#searchIframe").locator("li").first.is_visible():
            return "list"
        if page.frame_locator("#entryIframe").locator("body").is_visible():
            return "entry"
        page.wait_for_timeout(poll_ms)
        elapsed += poll_ms
    return "timeout"


def _current_list_item_texts(search_frame):
    """지금 화면에 있는 목록 항목들 중 '리뷰'가 포함된 것만 텍스트로 가져온다."""
    items = search_frame.locator("li").all()
    texts = []
    for it in items:
        try:
            t = it.inner_text(timeout=2000)
        except Exception:
            continue
        if "리뷰" in t:
            texts.append(t)
    return texts


def _go_to_next_page(search_frame):
    """목록 맨 아래 페이지 번호(1 2 3 4 5 >) 중 '다음 페이지' 화살표를 눌러본다.

    정확한 버튼 모양을 확인할 방법이 없어서 몇 가지 후보를 순서대로 시도한다.
    성공하면 True, 다음 페이지가 없거나 버튼을 못 찾으면 False를 돌려준다.
    """
    candidates = [
        lambda: search_frame.get_by_role("link", name=re.compile("다음")),
        lambda: search_frame.locator("a[aria-label*='다음']"),
        lambda: search_frame.locator("a:has-text('다음페이지')"),
        lambda: search_frame.locator("a.eUTV2"),  # 네이버 지도 페이지네이션에 흔히 쓰이는 클래스(추정)
    ]
    for make_locator in candidates:
        try:
            locator = make_locator()
            if locator.count() > 0 and locator.first.is_enabled():
                locator.first.click(timeout=3000)
                return True
        except Exception:
            continue
    return False


def _hover_over_list(page, search_frame):
    """가게 목록 항목 위로 마우스를 옮긴다. 목록 바깥(지도 등)에서 휠을 돌리면
    목록이 아니라 다른 곳이 스크롤돼서, 실제 목록 카드의 화면 좌표를 찾아 그 위로
    옮긴 뒤 스크롤해야 한다. 좌표를 못 구하면 대략적인 위치로 대신한다.
    """
    try:
        box = search_frame.locator("li").first.bounding_box()
        if box:
            page.mouse.move(box["x"] + box["width"] / 2, box["y"] + box["height"] / 2)
            return
    except Exception:
        pass
    page.mouse.move(200, 500)


def _scroll_current_page(page, search_frame, max_scrolls=25, stable_limit=3):
    """한 페이지 안에서는 무한 스크롤로 계속 더 불러와진다. 더 이상 새 항목이
    안 늘어날 때까지(또는 max_scrolls번 스크롤할 때까지) 마우스 휠로 내리면서 모은다.
    """
    _hover_over_list(page, search_frame)

    texts = _current_list_item_texts(search_frame)
    stable_rounds = 0
    for _ in range(max_scrolls):
        _hover_over_list(page, search_frame)  # 스크롤되며 목록 위치가 바뀔 수 있어 매번 다시 확인
        page.mouse.wheel(0, 2000)
        page.wait_for_timeout(700)
        new_texts = _current_list_item_texts(search_frame)

        if len(new_texts) <= len(texts):
            stable_rounds += 1
            if stable_rounds >= stable_limit:
                break
        else:
            stable_rounds = 0

        texts = new_texts

    return texts


def _collect_list_items_across_pages(page, max_pages):
    """검색 결과 목록은 한 페이지 안에서 무한 스크롤로 꽤 많이 불러와지고,
    그 스크롤이 끝나면 '페이지 번호(1,2,3...)'로 다음 목록으로 넘어가는 구조다.
    한 페이지를 끝까지 스크롤해서 다 모은 뒤, '다음 페이지' 버튼을 눌러가며 반복한다.
    """
    search_frame = page.frame_locator("#searchIframe")
    all_texts = []

    for page_num in range(1, max_pages + 1):
        page.wait_for_timeout(800)  # 페이지 전환 후 목록이 그려질 시간을 준다
        page_texts = _scroll_current_page(page, search_frame)
        all_texts.extend(page_texts)
        print(f"    {page_num}페이지: {len(page_texts)}개 항목")

        if page_num < max_pages:
            moved = _go_to_next_page(search_frame)
            if not moved:
                print(f"    다음 페이지 버튼을 찾지 못했습니다 ({page_num}페이지에서 멈춤).")
                break
            page.wait_for_timeout(1200)

    return all_texts


def collect_candidates():
    """config.py에 있는 지역 x 음식종류 조합마다 네이버 지도에서 검색 -> 목록 스크롤
    -> 리뷰 1000개 이상인 후보만 골라 모은다. (browser context/page는 이 함수 안에서
    새로 열고 닫는다.)
    """
    candidates = {}
    debug_shown = False

    with sync_playwright() as p:
        context, page = _open_browser(p)

        for group in config.PRODUCT_GROUPS:
            query = f"{config.DISTRICT} {group}"
            print(f"'{query}' 검색 중...")
            page.goto(f"https://map.naver.com/p/search/{quote(query)}", timeout=30000)

            state = _wait_for_list_or_entry(page)

            if state == "entry":
                # 검색 결과가 통째로 1건뿐이라 목록 없이 바로 상세 페이지로 들어간 경우
                try:
                    entry_frame = page.frame_locator("#entryIframe")
                    body_text = entry_frame.locator("body").inner_text(timeout=10000)
                    review_count = _parse_review_count(body_text)
                    name, _ = _extract_name_and_category(body_text)
                    if review_count is not None and review_count >= config.MIN_REVIEW_COUNT and name:
                        candidates.setdefault(name, {"가게이름": name, "상품군": group, "리뷰수": review_count})
                except Exception as e:
                    print(f"  -> 상세 페이지를 읽지 못했습니다: {e}")
            elif state == "list":
                item_texts = _collect_list_items_across_pages(page, max_pages=config.MAX_LIST_PAGES)

                if not debug_shown:
                    print("  ---- (목록이 잘 읽히는지 확인용) 처음 3개 항목 원문 ----")
                    for t in item_texts[:3]:
                        print("  >>>", repr(t))
                    print("  --------------------------------------------------------")
                    debug_shown = True

                found_in_query = 0
                for text in item_texts:
                    parsed = _parse_list_item(text)
                    if not parsed:
                        continue
                    name, review_count = parsed
                    if review_count < config.MIN_REVIEW_COUNT:
                        continue
                    if name not in candidates:
                        candidates[name] = {"가게이름": name, "상품군": group, "리뷰수": review_count}
                        found_in_query += 1
                print(f"  -> 목록 {len(item_texts)}개 중 리뷰 {config.MIN_REVIEW_COUNT}개 이상 신규 {found_in_query}곳")
            else:
                print("  -> 목록도 상세 페이지도 뜨지 않았습니다 (타임아웃). 건너뜁니다.")

            time.sleep(random.uniform(2.0, 4.0))

        context.close()

    return list(candidates.values())


def _try_get_extra_info(page, name):
    """가게 이름으로 다시 검색해 상세 페이지에 들어가서 주소/카테고리/AI 브리핑/
    지도 링크를 읽어온다. (리뷰 수는 목록 단계에서 이미 구했으므로 여기서는 안 읽는다.)

    반환값: (정보 dict 또는 None, 실패 원인을 설명하는 디버그 문자열 또는 None)
    """
    # 주소 전체를 검색어로 쓰면 실제 사람이 잘 안 쓰는 특이한 검색어라 자동화로
    # 의심받기 쉬워서, 가게이름 + 지역구 정도로 짧고 자연스러운 검색어를 사용한다.
    query = quote(f"{name} {config.DISTRICT}")
    page.goto(f"https://map.naver.com/p/search/{query}", timeout=30000)

    # 이름이 비슷한 가게가 여러 곳이면 '목록'이 뜨고, 검색결과가 하나뿐이면
    # 목록 없이 바로 '상세 페이지'로 넘어간다. 어느 쪽인지 지켜보다가
    # 목록이 뜬 경우에만 첫 번째 항목을 클릭한다.
    state = _wait_for_list_or_entry(page)
    if state == "list":
        try:
            search_frame = page.frame_locator("#searchIframe")
            search_frame.locator("li").filter(has_text=name).first.click(timeout=15000)
            page.wait_for_timeout(2500)
        except Exception:
            pass

    try:
        entry_frame = page.frame_locator("#entryIframe")
        # 상세 페이지(entryIframe)도 내용이 다 그려질 때까지 넉넉하게 기다린다.
        body_text = entry_frame.locator("body").inner_text(timeout=15000)

        # 'AI 브리핑' 제목/안내 문구는 스크롤 없이도 바로 뜨지만, 실제 요약
        # 문장(동그라미 항목들)은 한 번 더 로딩돼야 나온다. 그 실제 내용이 뜰 때만
        # 나오는 "정리한 정보는 다음과 같습니다" 문구가 보일 때까지 스크롤한다.
        # (창 크기에 따라 위치가 달라질 수 있어 정보 패널의 실제 화면 좌표를 찾아 그 위에서 스크롤한다.)
        try:
            box = entry_frame.locator("body").bounding_box()
            if box:
                page.mouse.move(box["x"] + box["width"] / 2, box["y"] + box["height"] / 2)
            else:
                page.mouse.move(220, 400)
        except Exception:
            page.mouse.move(220, 400)
        for _ in range(8):
            if "정리한 정보는 다음과 같습니다" in body_text:
                break
            page.mouse.wheel(0, 1200)
            page.wait_for_timeout(600)
            body_text = entry_frame.locator("body").inner_text(timeout=5000)
    except Exception as e:
        frame_names = [f.name or "(이름없음)" for f in page.frames]

        if any("captcha" in fn.lower() for fn in frame_names):
            debug = (
                "네이버가 이 접속을 '자동화 프로그램'으로 의심해서 캡차(사람 인증) 화면을 띄웠습니다.\n"
                "config.py에서 HEADLESS = False 로 바꾼 뒤 다시 실행해서, 뜨는 브라우저 창에서\n"
                "캡차를 직접 한 번 풀어주세요. 이 프로그램은 브라우저 정보를 저장해두기 때문에,\n"
                "한 번 풀고 나면 다음 실행부터는 안 떠야 정상입니다. 그래도 계속 뜨면,\n"
                "요청 속도를 더 늦추거나(가게 수를 줄이거나) 시간을 두고 다시 시도해보세요.\n"
                f"현재 페이지 URL: {page.url}"
            )
        else:
            debug = (
                f"entryIframe을 찾지 못했습니다 ({e}).\n"
                f"현재 페이지 URL: {page.url}\n"
                f"현재 페이지에 있는 프레임 이름 목록: {frame_names}"
            )
        return None, debug

    _, category = _extract_name_and_category(body_text)
    address = _district_only(_extract_address(body_text))
    description, briefing_debug = _parse_ai_briefing(body_text)

    info = {
        "카테고리": category,
        "주소": address,
        "네이버지도 주소": page.url,
        "브랜드설명": description or "",
        "_브리핑디버그": briefing_debug,
    }
    return info, None


def get_extra_info(page, name, max_attempts=2):
    """캡차가 아닌 실패는 타이밍 문제일 수 있으므로 한 번 더 재시도한다."""
    debug_info = None
    for attempt in range(1, max_attempts + 1):
        info, debug_info = _try_get_extra_info(page, name)
        if info is not None:
            return info, None

        if debug_info and "캡차" in debug_info:
            return None, debug_info

        if attempt < max_attempts:
            page.wait_for_timeout(2000)

    return None, debug_info


def _open_browser(p):
    """launch_persistent_context: 매번 새 브라우저가 아니라 browser_profile 폴더에
    쿠키/방문 기록을 저장해두고 재사용한다. 캡차를 한 번 풀면 그 기록이 남아서
    다음 실행부터는 덜 의심받는다.
    """
    context = p.chromium.launch_persistent_context(
        PROFILE_DIR,
        headless=config.HEADLESS,
        user_agent=(
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0 Safari/537.36"
        ),
        viewport={"width": 1400, "height": 900},
        args=["--disable-blink-features=AutomationControlled"],
    )
    # 자동화 브라우저임을 알리는 대표적인 신호(navigator.webdriver)를 숨긴다.
    context.add_init_script(
        "Object.defineProperty(navigator, 'webdriver', {get: () => undefined})"
    )
    page = context.new_page()
    return context, page


def _display_width(text):
    """열 너비/줄 수를 어림잡기 위한 글자 폭 계산. 한글 등 넓은 글자는 2칸으로 센다."""
    width = 0
    for ch in str(text):
        width += 2 if ord(ch) > 0x1100 else 1
    return width


def save_excel(rows, output_file):
    """결과를 엑셀로 저장하면서 열 너비/행 높이 자동 맞춤, 헤더 색상, 폰트를 적용한다."""
    wb = Workbook()
    ws = wb.active
    ws.title = "결과"
    ws.append(HEADERS)

    for row in rows:
        ws.append([row.get(h, "") for h in HEADERS])

    desc_col = HEADERS.index("브랜드설명") + 1
    link_col = HEADERS.index("네이버지도 주소") + 1

    # 1행(헤더): 회색 배경 + 지정 폰트 + 굵게
    for cell in ws[1]:
        cell.font = Font(name=FONT_NAME, bold=True)
        cell.fill = HEADER_FILL

    # 본문: 폰트 적용, 설명 칸은 줄바꿈 허용, 지도 링크 칸은 클릭 가능한 하이퍼링크로
    for row_cells in ws.iter_rows(min_row=2):
        for cell in row_cells:
            if cell.column == link_col and cell.value:
                cell.hyperlink = cell.value
                cell.font = Font(name=FONT_NAME, color="0563C1", underline="single")
            else:
                cell.font = Font(name=FONT_NAME)
            if cell.column == desc_col:
                cell.alignment = Alignment(wrap_text=True, vertical="top")

    # 열 너비 자동 맞춤 (글자 폭 기준 근사치 - 엑셀의 진짜 자동맞춤과 100% 같지는 않음)
    for col_index, header in enumerate(HEADERS, start=1):
        max_width = max(
            [_display_width(header)] + [_display_width(row.get(header, "")) for row in rows]
        )
        cap = 50 if header in ("브랜드설명", "네이버지도 주소") else 40
        ws.column_dimensions[get_column_letter(col_index)].width = min(max_width + 4, cap)

    # 행 높이 자동 맞춤 (설명 칸이 몇 줄로 접힐지 어림잡아 계산)
    desc_col_width = ws.column_dimensions[get_column_letter(desc_col)].width or 50
    for row_index, row in enumerate(rows, start=2):
        desc_text = str(row.get("브랜드설명", ""))
        lines_needed = max(1, -(-_display_width(desc_text) // int(desc_col_width)))
        ws.row_dimensions[row_index].height = max(15, lines_needed * 15)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    wb.save(output_file)


def main():
    candidates = collect_candidates()
    if not candidates:
        print("검색된 가게가 없습니다. config.py의 검색 조건을 확인하세요.")
        return

    print(f"1차로 찾은 (리뷰 {config.MIN_REVIEW_COUNT}개 이상) 후보: {len(candidates)}곳")

    final_rows = []

    with sync_playwright() as p:
        context, page = _open_browser(p)

        debug_shown = 0
        briefing_debug_shown = 0
        for i, place in enumerate(candidates, 1):
            print(f"[{i}/{len(candidates)}] 상세정보 확인 중: {place['가게이름']} (리뷰 {place['리뷰수']})")
            try:
                info, debug_info = get_extra_info(page, place["가게이름"])
            except Exception as e:
                info, debug_info = None, f"예외 발생: {e}"

            if info is None:
                print("  -> 상세정보를 찾지 못했습니다. 건너뜁니다.")
                if debug_info and debug_shown < 2:
                    print("  ========== 디버그 정보 (이 부분을 복사해서 보내주세요) ==========")
                    print(debug_info)
                    print("  =================================================================")
                    debug_shown += 1
                continue

            if briefing_debug_shown < 2:
                print("  ---- (브랜드설명이 맞는지 확인용) AI 브리핑 주변 텍스트 ----")
                print(f"  현재 추출된 브랜드설명: {info['브랜드설명']!r}")
                print(f"  {info['_브리핑디버그']}")
                print("  ------------------------------------------------------------")
                briefing_debug_shown += 1

            final_rows.append({
                "가게이름": place["가게이름"],
                "상품군": place["상품군"],
                "카테고리": info["카테고리"],
                "주소": info["주소"],
                "네이버지도 주소": info["네이버지도 주소"],
                "리뷰수": place["리뷰수"],
                "브랜드설명": info["브랜드설명"],
            })

            # 너무 빠르게 계속 요청하지 않도록 잠깐 대기
            # (사이트에 부담을 주지 않고, 자동화로 의심받지 않기 위함)
            time.sleep(random.uniform(3.0, 6.0))

        context.close()

    if not final_rows:
        print(f"리뷰 {config.MIN_REVIEW_COUNT}개 이상인 가게가 없습니다.")
        return

    final_rows.sort(key=lambda r: r["리뷰수"], reverse=True)
    save_excel(final_rows, config.OUTPUT_FILE)
    print(f"완료! {len(final_rows)}곳을 '{config.OUTPUT_FILE}' 파일로 저장했습니다.")


if __name__ == "__main__":
    main()
