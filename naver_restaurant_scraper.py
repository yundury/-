"""
네이버 지역검색 API로 특정 지역의 한식당 목록을 모으고,
각 가게의 네이버 플레이스(지도) 페이지에서 리뷰 개수를 가져온 뒤,
리뷰 개수가 설정한 값 이상인 가게만 엑셀 파일로 저장하는 스크립트.

실행 방법은 README.md를 참고하세요.
모든 설정값(API 키, 지역, 최소 리뷰 수 등)은 config.py에서 바꿉니다.
"""

import re
import time
import random
from urllib.parse import quote

import requests
import pandas as pd
from playwright.sync_api import sync_playwright

import config


LOCAL_SEARCH_URL = "https://openapi.naver.com/v1/search/local.json"


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


def get_review_count(page, name, address):
    """네이버 지도에서 가게 이름+주소로 검색해 들어간 뒤, 리뷰 개수를 읽어온다.

    네이버 플레이스는 '방문자리뷰'와 '블로그리뷰'를 따로 보여주고
    하나로 합친 '총 리뷰 수'는 제공하지 않는 경우가 많아,
    이 함수는 두 값을 더한 것을 '리뷰개수'로 취급한다.

    반환값: (리뷰개수 또는 None, 실패 원인을 설명하는 디버그 문자열 또는 None)
    """
    query = quote(f"{name} {address}")
    page.goto(f"https://map.naver.com/p/search/{query}", timeout=30000)
    page.wait_for_timeout(2500)

    try:
        search_frame = page.frame_locator("#searchIframe")
        # 검색 결과 목록에서 가게 이름이 포함된 첫 번째 항목 클릭
        search_frame.locator("li").filter(has_text=name).first.click(timeout=8000)
        page.wait_for_timeout(2000)
    except Exception:
        # 검색 결과가 1건이라 목록 없이 바로 상세 페이지로 들어간 경우일 수 있음
        pass

    try:
        entry_frame = page.frame_locator("#entryIframe")
        body_text = entry_frame.locator("body").inner_text(timeout=8000)
    except Exception as e:
        frame_names = [f.name or "(이름없음)" for f in page.frames]
        debug = (
            f"entryIframe을 찾지 못했습니다 ({e}).\n"
            f"현재 페이지 URL: {page.url}\n"
            f"현재 페이지에 있는 프레임 이름 목록: {frame_names}"
        )
        return None, debug

    visitor = re.search(r"방문자\s*리뷰\s*([\d,]+)", body_text)
    blog = re.search(r"블로그\s*리뷰\s*([\d,]+)", body_text)

    if not visitor and not blog:
        debug = (
            f"'방문자리뷰'/'블로그리뷰' 글자를 페이지에서 못 찾았습니다.\n"
            f"현재 페이지 URL: {page.url}\n"
            f"entryIframe에서 읽은 글자 (앞부분 800자):\n{body_text[:800]}"
        )
        return None, debug

    visitor_count = int(visitor.group(1).replace(",", "")) if visitor else 0
    blog_count = int(blog.group(1).replace(",", "")) if blog else 0

    return visitor_count + blog_count, None


def main():
    candidates = search_restaurants()
    if not candidates:
        print("검색된 한식당이 없습니다. config.py의 검색 조건을 확인하세요.")
        return

    final_rows = []

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=config.HEADLESS)
        page = browser.new_page(
            user_agent=(
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0 Safari/537.36"
            )
        )

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

            # 너무 빠르게 계속 요청하지 않도록 잠깐 대기 (사이트에 부담을 주지 않기 위함)
            time.sleep(random.uniform(1.5, 3.0))

        browser.close()

    if not final_rows:
        print(f"리뷰 {config.MIN_REVIEW_COUNT}개 이상인 가게가 없습니다.")
        return

    df = pd.DataFrame(final_rows, columns=["가게이름", "주소", "카테고리", "리뷰개수"])
    df = df.sort_values("리뷰개수", ascending=False)
    df.to_excel(config.OUTPUT_FILE, index=False)
    print(f"완료! {len(df)}곳을 '{config.OUTPUT_FILE}' 파일로 저장했습니다.")


if __name__ == "__main__":
    main()
