"""
네이버 지역검색 API로 특정 지역의 가게 목록을 모으고,
각 가게의 네이버 플레이스(지도) 페이지에서 리뷰 개수/AI 브리핑을 가져온 뒤,
리뷰 개수가 설정한 값 이상인 가게만 엑셀 파일로 저장하는 스크립트.

실행 방법은 README.md를 참고하세요.
모든 설정값(API 키, 지역, 찾을 음식 종류, 최소 리뷰 수 등)은 config.py에서 바꿉니다.
"""

import os
import re
import time
import random
from urllib.parse import quote

import requests
from playwright.sync_api import sync_playwright
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

import config


LOCAL_SEARCH_URL = "https://openapi.naver.com/v1/search/local.json"

# 브라우저 로그인/쿠키 정보를 저장해두는 폴더.
# 매번 새 브라우저(신규 방문자)인 것처럼 접속하면 네이버가 자동화로 의심하기 쉬워서,
# 같은 브라우저 프로필을 계속 재사용해 "이전에도 왔던 사람"처럼 보이게 한다.
PROFILE_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "browser_profile")

# config.py의 PRODUCT_GROUPS에서 고른 상품군마다 지역검색에 같이 붙일 검색어.
SEARCH_KEYWORDS_BY_GROUP = {
    "한식": ["한식", "맛집"],
    "중식": ["중식", "중국집"],
    "일식": ["일식", "스시", "돈카츠"],
    "양식": ["양식", "레스토랑"],
    "에스닉": ["아시안음식", "베트남음식", "태국음식", "인도음식"],
    "베이커리": ["베이커리", "빵집"],
    "디저트": ["디저트", "카페"],
}

# 최종 엑셀 컬럼 순서
HEADERS = ["가게이름", "상품군", "카테고리", "주소", "네이버지도 주소", "리뷰수", "브랜드설명"]

# 엑셀 서식에 쓸 값들
FONT_NAME = "나눔바른고딕"
HEADER_FILL = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")


def _classify_category(category_raw):
    """네이버가 주는 카테고리 문자열(예: '한식>닭갈비', '음식점>한식>육류,고기요리>돼지고기구이')에서
    상품군(한식/중식/양식/일식/에스닉/베이커리/디저트)과 세부 카테고리(닭갈비 등)를 뽑아낸다.
    """
    segments = [s.strip() for s in category_raw.split(">") if s.strip() and s.strip() != "음식점"]

    keyword_map = [
        ("한식", ["한식"]),
        ("중식", ["중식", "중국음식"]),
        ("일식", ["일식", "일본음식"]),
        ("양식", ["양식", "이탈리안", "프렌치", "스테이크"]),
        ("베이커리", ["베이커리", "빵집", "제과"]),
        ("디저트", ["디저트", "카페"]),
        ("에스닉", ["아시아음식", "베트남", "태국", "인도음식", "멕시코", "중동음식", "에스닉"]),
    ]

    product_group = "기타"
    matched_segment = None
    for seg in segments:
        for group, keywords in keyword_map:
            if any(k in seg for k in keywords):
                product_group = group
                matched_segment = seg
                break
        if matched_segment:
            break

    detail_segments = [s for s in segments if s != matched_segment]
    detail = detail_segments[-1] if detail_segments else ""
    return product_group, detail


def _district_only(address):
    """'서울특별시 강남구 논현로152길 36 1층' -> '서울특별시 강남구'처럼 구까지만 남긴다."""
    parts = address.split()
    return " ".join(parts[:2]) if len(parts) >= 2 else address


def search_restaurants():
    """네이버 지역검색 API로 config.PRODUCT_GROUPS에 해당하는 가게 후보를 모은다."""
    headers = {
        "X-Naver-Client-Id": config.NAVER_CLIENT_ID,
        "X-Naver-Client-Secret": config.NAVER_CLIENT_SECRET,
    }

    results = {}

    search_keywords = []
    for group in config.PRODUCT_GROUPS:
        search_keywords.extend(SEARCH_KEYWORDS_BY_GROUP.get(group, [group]))

    queries = [
        f"{config.DISTRICT} {neighborhood} {keyword}"
        for neighborhood in config.NEIGHBORHOODS
        for keyword in search_keywords
    ]

    for query in queries:
        params = {
            "query": query,
            "display": 5,  # 네이버 지역검색 API 특성상 한 번에 최대 5개까지만 받을 수 있음
            "start": 1,
            "sort": "comment",  # 리뷰(댓글) 많은 순으로 정렬해서 상위 5개를 받음
        }
        try:
            resp = requests.get(LOCAL_SEARCH_URL, headers=headers, params=params, timeout=10)
            resp.raise_for_status()
        except requests.RequestException as e:
            print(f"[검색 실패] '{query}': {e}")
            continue

        items = resp.json().get("items", [])
        for item in items:
            name = re.sub("<.*?>", "", item.get("title", ""))
            category_raw = item.get("category", "")
            address = item.get("roadAddress") or item.get("address", "")

            product_group, detail_category = _classify_category(category_raw)
            if product_group not in config.PRODUCT_GROUPS:
                continue

            key = (name, address)
            if key not in results:
                results[key] = {
                    "가게이름": name,
                    "상품군": product_group,
                    "카테고리": detail_category,
                    "주소": _district_only(address),
                }

        time.sleep(0.3)  # 초당 API 호출 제한을 지키기 위한 짧은 대기

    print(f"지역검색으로 찾은 후보: {len(results)}곳")
    return list(results.values())


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


def _parse_ai_briefing(text, max_items=2):
    """'AI 요약'/'AI 브리핑' 라벨 주변의 짧은 소개 문구를 최대 max_items개까지 가져온다.

    'AI 브리핑' 아래에는 보통 요약 문장이 몇 개 나오고, 문장마다 끝에
    '닉네임 +2' 같은 출처 표시가 붙는다. 그 표시를 기준으로 문장 단위를 나눈다.
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


def _try_get_place_details(page, name):
    """실제로 검색 -> (필요하면) 목록 클릭 -> 상세 페이지 읽기를 한 번 시도한다.

    반환값: (상세정보 dict 또는 None, 실패 원인을 설명하는 디버그 문자열 또는 None)
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

        # 'AI 브리핑'은 스크롤을 어느 정도 내려야 나타나는 지연 로딩 영역이라,
        # 왼쪽 정보 패널 위에서 마우스 휠을 내리면서 나타날 때까지 기다린다.
        page.mouse.move(220, 400)
        for _ in range(6):
            if "브리핑" in body_text:
                break
            page.mouse.wheel(0, 1200)
            page.wait_for_timeout(500)
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

    review_count = _parse_review_count(body_text)
    if review_count is None:
        debug = (
            f"'리뷰' 뒤에 오는 숫자를 페이지에서 못 찾았습니다.\n"
            f"현재 페이지 URL: {page.url}\n"
            f"entryIframe에서 읽은 글자 (앞부분 800자):\n{body_text[:800]}"
        )
        return None, debug

    description, briefing_debug = _parse_ai_briefing(body_text)

    details = {
        "리뷰수": review_count,
        "네이버지도 주소": page.url,
        "브랜드설명": description or "",
        "_브리핑디버그": briefing_debug,
    }
    return details, None


def get_place_details(page, name, max_attempts=2):
    """네이버 지도에서 가게 이름으로 검색해 들어간 뒤, 리뷰 수/AI 브리핑/지도 링크를 읽어온다.

    타이밍이 꼬여서 실패하는 경우를 대비해 캡차가 아닌 실패는 한 번 더 재시도한다.

    반환값: (상세정보 dict 또는 None, 실패 원인을 설명하는 디버그 문자열 또는 None)
    """
    debug_info = None
    for attempt in range(1, max_attempts + 1):
        details, debug_info = _try_get_place_details(page, name)
        if details is not None:
            return details, None

        if debug_info and "캡차" in debug_info:
            # 캡차는 다시 시도해도 똑같이 막히므로 바로 포기하고 사용자에게 알린다.
            return None, debug_info

        if attempt < max_attempts:
            page.wait_for_timeout(2000)

    return None, debug_info


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
    candidates = search_restaurants()
    if not candidates:
        print("검색된 가게가 없습니다. config.py의 검색 조건을 확인하세요.")
        return

    final_rows = []

    with sync_playwright() as p:
        # launch_persistent_context: 매번 새 브라우저가 아니라 browser_profile 폴더에
        # 쿠키/방문 기록을 저장해두고 재사용한다. 캡차를 한 번 풀면 그 기록이 남아서
        # 다음 실행부터는 덜 의심받는다.
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

        debug_shown = 0
        briefing_debug_shown = 0
        for i, place in enumerate(candidates, 1):
            print(f"[{i}/{len(candidates)}] 확인 중: {place['가게이름']}")
            try:
                details, debug_info = get_place_details(page, place["가게이름"])
            except Exception as e:
                details, debug_info = None, f"예외 발생: {e}"

            if details is None:
                print("  -> 정보를 찾지 못했습니다. 건너뜁니다.")
                if debug_info and debug_shown < 2:
                    print("  ========== 디버그 정보 (이 부분을 복사해서 보내주세요) ==========")
                    print(debug_info)
                    print("  =================================================================")
                    debug_shown += 1
                continue

            review_count = details["리뷰수"]
            print(f"  -> 리뷰 수: {review_count}")

            if briefing_debug_shown < 2:
                print("  ---- (브랜드설명이 맞는지 확인용) AI 요약 라벨 주변 텍스트 ----")
                print(f"  현재 추출된 브랜드설명: {details['브랜드설명']!r}")
                print(f"  {details['_브리핑디버그']}")
                print("  ---------------------------------------------------------------")
                briefing_debug_shown += 1

            if review_count >= config.MIN_REVIEW_COUNT:
                final_rows.append({
                    "가게이름": place["가게이름"],
                    "상품군": place["상품군"],
                    "카테고리": place["카테고리"],
                    "주소": place["주소"],
                    "네이버지도 주소": details["네이버지도 주소"],
                    "리뷰수": review_count,
                    "브랜드설명": details["브랜드설명"],
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
