"""
네이버 지역검색 API로 특정 지역의 한식당 목록을 모으고,
각 가게의 네이버 플레이스(지도) 페이지에서 리뷰 개수를 가져온 뒤,
리뷰 개수가 설정한 값 이상인 가게만 엑셀 파일로 저장하는 스크립트.

실행 방법은 README.md를 참고하세요.
모든 설정값(API 키, 지역, 최소 리뷰 수 등)은 config.py에서 바꿉니다.
"""

import os
import re
import time
import random
from urllib.parse import quote

import requests
import pandas as pd
from playwright.sync_api import sync_playwright

import config


LOCAL_SEARCH_URL = "https://openapi.naver.com/v1/search/local.json"

# 브라우저 로그인/쿠키 정보를 저장해두는 폴더.
# 매번 새 브라우저(신규 방문자)인 것처럼 접속하면 네이버가 자동화로 의심하기 쉬워서,
# 같은 브라우저 프로필을 계속 재사용해 "이전에도 왔던 사람"처럼 보이게 한다.
PROFILE_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "browser_profile")


def search_restaurants():
    """네이버 지역검색 API로 한식당 후보 목록을 모은다. (이름/주소 기준 중복 제거)"""
    headers = {
        "X-Naver-Client-Id": config.NAVER_CLIENT_ID,
        "X-Naver-Client-Secret": config.NAVER_CLIENT_SECRET,
    }

    results = {}

    queries = [
        f"{config.DISTRICT} {neighborhood} {keyword}"
        for neighborhood in config.NEIGHBORHOODS
        for keyword in config.KEYWORDS
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
            category = item.get("category", "")
            address = item.get("roadAddress") or item.get("address", "")

            if config.CATEGORY_FILTER not in category:
                continue

            key = (name, address)
            if key not in results:
                results[key] = {
                    "가게이름": name,
                    "주소": address,
                    "카테고리": category,
                }

        time.sleep(0.3)  # 초당 API 호출 제한을 지키기 위한 짧은 대기

    print(f"지역검색으로 찾은 한식당 후보: {len(results)}곳")
    return list(results.values())


def _parse_review_count(text):
    """'리뷰 1.1만', '리뷰 1,234', '방문자리뷰 966' 같은 표기에서 숫자를 뽑아낸다.

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


def _try_get_review_count(page, name):
    """실제로 검색 -> (필요하면) 목록 클릭 -> 상세 페이지 읽기를 한 번 시도한다."""
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

    return review_count, None


def get_review_count(page, name, address, max_attempts=2):
    """네이버 지도에서 가게 이름으로 검색해 들어간 뒤, 리뷰 개수를 읽어온다.

    타이밍이 꼬여서 실패하는 경우를 대비해 캡차가 아닌 실패는 한 번 더 재시도한다.

    반환값: (리뷰개수 또는 None, 실패 원인을 설명하는 디버그 문자열 또는 None)
    """
    debug_info = None
    for attempt in range(1, max_attempts + 1):
        review_count, debug_info = _try_get_review_count(page, name)
        if review_count is not None:
            return review_count, None

        if debug_info and "캡차" in debug_info:
            # 캡차는 다시 시도해도 똑같이 막히므로 바로 포기하고 사용자에게 알린다.
            return None, debug_info

        if attempt < max_attempts:
            page.wait_for_timeout(2000)

    return None, debug_info


def main():
    candidates = search_restaurants()
    if not candidates:
        print("검색된 한식당이 없습니다. config.py의 검색 조건을 확인하세요.")
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
        for i, place in enumerate(candidates, 1):
            print(f"[{i}/{len(candidates)}] 리뷰 수 확인 중: {place['가게이름']}")
            try:
                review_count, debug_info = get_review_count(page, place["가게이름"], place["주소"])
            except Exception as e:
                review_count, debug_info = None, f"예외 발생: {e}"

            if review_count is None:
                print("  -> 리뷰 수를 찾지 못했습니다. 건너뜁니다.")
                if debug_info and debug_shown < 2:
                    print("  ========== 디버그 정보 (이 부분을 복사해서 보내주세요) ==========")
                    print(debug_info)
                    print("  =================================================================")
                    debug_shown += 1
                continue

            print(f"  -> 리뷰 수: {review_count}")

            if review_count >= config.MIN_REVIEW_COUNT:
                final_rows.append({
                    "가게이름": place["가게이름"],
                    "주소": place["주소"],
                    "카테고리": place["카테고리"],
                    "리뷰개수": review_count,
                })

            # 너무 빠르게 계속 요청하지 않도록 잠깐 대기
            # (사이트에 부담을 주지 않고, 자동화로 의심받지 않기 위함)
            time.sleep(random.uniform(3.0, 6.0))

        context.close()

    if not final_rows:
        print(f"리뷰 {config.MIN_REVIEW_COUNT}개 이상인 가게가 없습니다.")
        return

    df = pd.DataFrame(final_rows, columns=["가게이름", "주소", "카테고리", "리뷰개수"])
    df = df.sort_values("리뷰개수", ascending=False)
    df.to_excel(config.OUTPUT_FILE, index=False)
    print(f"완료! {len(df)}곳을 '{config.OUTPUT_FILE}' 파일로 저장했습니다.")


if __name__ == "__main__":
    main()
