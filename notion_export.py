"""
크롤링 결과(브랜드명/키워드/카테고리/리뷰수/대표 리뷰/네이버지도주소)를 노션(Notion)
데이터베이스(표)에도 올려주는 선택 기능. 엑셀 저장과는 별개로 동작하며, 엑셀
저장은 이 기능을 켜든 안 켜든 항상 그대로 됩니다.

사용하려면 config.py에서 NOTION_ENABLED = True 로 바꾸고, NOTION_TOKEN /
NOTION_PARENT_PAGE_ID 값을 채워야 합니다. 설정 방법은 README.md의
"9. (선택) 노션에 결과 올리기" 부분을 참고하세요.

노션 API는 이 프로젝트를 만든 환경(샌드박스)에서 접속이 막혀 있어서 실제로
호출해보며 테스트하지 못했습니다. 처음 몇 번 실행해보고 오류가 나면 그 메시지를
그대로 복사해서 보내주세요 - 정확한 원인에 맞춰 고쳐드리겠습니다.
"""

import json
import os
import time
import urllib.error
import urllib.request

NOTION_API_BASE = "https://api.notion.com/v1"
NOTION_VERSION = "2022-06-28"

# 노션에 만들 데이터베이스 제목
DATABASE_TITLE = "네이버 맛집 결과"

# 어느 부모 페이지 밑에 데이터베이스를 만들었는지 기억해두는 파일.
# (매번 새 데이터베이스를 만들지 않고, 한 번 만든 걸 계속 재사용하기 위함)
_DB_ID_CACHE_FILE = os.path.join(
    os.path.dirname(os.path.abspath(__file__)), ".notion_database_id.json"
)

# 노션 rich_text 한 블록에 들어갈 수 있는 글자 수 제한에 여유를 두고 자른다.
_MAX_TEXT_LEN = 1900


def _notion_request(method, path, token, body=None):
    url = f"{NOTION_API_BASE}{path}"
    data = json.dumps(body).encode("utf-8") if body is not None else None
    req = urllib.request.Request(url, data=data, method=method)
    req.add_header("Authorization", f"Bearer {token}")
    req.add_header("Notion-Version", NOTION_VERSION)
    req.add_header("Content-Type", "application/json")
    try:
        with urllib.request.urlopen(req, timeout=15) as resp:
            return json.loads(resp.read().decode("utf-8"))
    except urllib.error.HTTPError as e:
        err_body = e.read().decode("utf-8", errors="replace")
        raise RuntimeError(f"노션 API 오류 ({e.code}): {err_body}") from e
    except urllib.error.URLError as e:
        raise RuntimeError(f"노션에 접속하지 못했습니다: {e}") from e


def _load_db_cache():
    if os.path.exists(_DB_ID_CACHE_FILE):
        try:
            with open(_DB_ID_CACHE_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}
    return {}


def _save_db_cache(cache):
    with open(_DB_ID_CACHE_FILE, "w", encoding="utf-8") as f:
        json.dump(cache, f, ensure_ascii=False, indent=2)


def _get_or_create_database(token, parent_page_id):
    """지정한 노션 페이지 밑에, 결과를 담을 데이터베이스가 이미 있으면 재사용하고
    없으면 새로 만든다. (parent_page_id별로 따로 기억하므로, 다른 페이지를
    쓰면 그 페이지 밑에 새로 하나 더 만들어진다.)
    """
    cache = _load_db_cache()
    if parent_page_id in cache:
        return cache[parent_page_id]

    result = _notion_request(
        "POST",
        "/databases",
        token,
        body={
            "parent": {"type": "page_id", "page_id": parent_page_id},
            "title": [{"type": "text", "text": {"content": DATABASE_TITLE}}],
            "properties": {
                "브랜드명": {"title": {}},
                "키워드": {"rich_text": {}},
                "카테고리": {"rich_text": {}},
                "리뷰수": {"number": {}},
                "대표 리뷰": {"rich_text": {}},
                "네이버지도주소": {"url": {}},
            },
        },
    )
    database_id = result["id"]
    cache[parent_page_id] = database_id
    _save_db_cache(cache)
    return database_id


def _truncate(text):
    text = str(text or "")
    return text[:_MAX_TEXT_LEN]


def _row_to_properties(row):
    map_url = row.get("네이버지도주소") or None
    return {
        "브랜드명": {"title": [{"text": {"content": _truncate(row.get("브랜드명", ""))}}]},
        "키워드": {"rich_text": [{"text": {"content": _truncate(row.get("키워드", ""))}}]},
        "카테고리": {"rich_text": [{"text": {"content": _truncate(row.get("카테고리", ""))}}]},
        "리뷰수": {"number": row.get("리뷰수", 0)},
        "대표 리뷰": {"rich_text": [{"text": {"content": _truncate(row.get("대표 리뷰", ""))}}]},
        "네이버지도주소": {"url": map_url},
    }


def _find_existing_page(token, database_id, brand_name):
    result = _notion_request(
        "POST",
        f"/databases/{database_id}/query",
        token,
        body={"filter": {"property": "브랜드명", "title": {"equals": brand_name}}},
    )
    results = result.get("results") or []
    return results[0]["id"] if results else None


def push_rows_to_notion(rows, token, parent_page_id):
    """rows(엑셀에 저장한 것과 같은 딕셔너리 리스트)를 노션 데이터베이스에 올린다.
    이미 같은 브랜드명의 행이 있으면 업데이트하고, 없으면 새로 추가한다.

    엑셀 저장이 이미 끝난 뒤에 호출되는 걸 전제로 하며, 여기서 오류가 나도
    이미 저장된 엑셀 파일에는 영향이 없다 (호출하는 쪽에서 try/except로 감싼다).
    """
    if not token or not parent_page_id:
        print("[노션] NOTION_TOKEN / NOTION_PARENT_PAGE_ID가 비어 있어서 노션 업로드를 건너뜁니다.")
        return

    print(f"[노션] 결과를 노션 표에 올리는 중... (총 {len(rows)}곳)")
    database_id = _get_or_create_database(token, parent_page_id)

    created, updated, failed = 0, 0, 0
    for row in rows:
        brand_name = row.get("브랜드명", "")
        try:
            existing_id = _find_existing_page(token, database_id, brand_name)
            properties = _row_to_properties(row)
            if existing_id:
                _notion_request("PATCH", f"/pages/{existing_id}", token, body={"properties": properties})
                updated += 1
            else:
                _notion_request(
                    "POST",
                    "/pages",
                    token,
                    body={"parent": {"database_id": database_id}, "properties": properties},
                )
                created += 1
        except Exception as e:
            failed += 1
            print(f"[노션] '{brand_name}' 업로드 실패: {e}")

        time.sleep(0.35)  # 노션 API 요청 속도 제한(초당 약 3건)을 넘지 않도록 살짝 쉰다

    print(f"[노션] 완료: 새로 추가 {created}곳 / 업데이트 {updated}곳 / 실패 {failed}곳")
