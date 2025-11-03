# -*- coding: utf-8 -*-
"""
PDMS 학습기업 스크레이퍼 (지역만 고정, 지사/참여유형명 자동 읽기)
- '지역'은 개발자가 목록(또는 일부)만 지정
- 각 지역 선택 후 '지사' 드롭다운 옵션을 동적으로 읽어서 전부 순회
- '참여유형명'(또는 '참여유형')도 화면에서 동적으로 읽어서 전부 순회
- 라벨/iframe/로딩대기/페이지네이션/디버그 스냅샷 대응
"""

import time, sys
from typing import List, Dict
import pandas as pd
from playwright.sync_api import sync_playwright, Page, Frame, Locator

URL = "https://pdms.ncs.go.kr/cmn/pub/opr/retrieveOprLrnEntrprList.do"

# ▶ 돌릴 지역만 골라서 넣으세요(처음엔 1~2개로 테스트 권장)
REGIONS = [
    "서울", "경기", "인천", "강원", "충북", "충남",
    "대전", "경북", "대구", "전북", "경남", "울산",
    "부산", "광주", "전남", "제주", "세종"
]

# ▶ 안전장치(테스트용): 각 조합당 페이지네이션 최대 몇 페이지까지?
MAX_PAGES_PER_COMBO = 9999  # 제한 없음. 테스트 땐 3 정도로 줄이세요.


# ----------------- 유틸 -----------------
def find_target_frame(page: Page) -> Page | Frame:
    """폼이 iframe 안에 있으면 프레임을 리턴"""
    if page.get_by_role("combobox").count() > 0:
        return page
    for f in page.frames:
        try:
            if f.get_by_role("combobox").count() > 0:
                return f
        except Exception:
            pass
    return page

def get_combobox_by_label(ctx: Page | Frame, label_text: str) -> Locator:
    """접근성 라벨명으로 select를 찾되, 실패 시 label 형제 select fallback"""
    cb = ctx.get_by_role("combobox", name=label_text)
    if cb.count() == 0:
        label = ctx.locator(f"label:has-text('{label_text}')").first
        if label.count() == 0:
            raise RuntimeError(f"[셀렉터] 라벨 '{label_text}'을 찾지 못했습니다.")
        for_id = label.get_attribute("for")
        sel = ctx.locator(f"select#{for_id}") if for_id else label.locator("xpath=following-sibling::select[1]")
        if sel.count() == 0:
            raise RuntimeError(f"[셀렉터] 라벨 '{label_text}'에 연결된 select를 찾지 못했습니다.")
        return sel
    return cb.first

def wait_options_loaded(select_el: Locator, min_count: int = 2, timeout_ms: int = 10000):
    """선택박스 옵션이 min_count 이상 채워질 때까지 대기"""
    select_el.wait_for(state="attached", timeout=timeout_ms)
    start = time.time()
    while True:
        try:
            if select_el.locator("option").count() >= min_count:
                return
        except Exception:
            pass
        if (time.time() - start) * 1000 > timeout_ms:
            raise TimeoutError("[대기초과] 옵션 로딩이 완료되지 않았습니다.")
        time.sleep(0.2)

def list_options_text(select_el: Locator) -> list[str]:
    """select의 모든 option 텍스트 리스트 반환(공백 제거)"""
    opts = select_el.locator("option")
    return [opts.nth(i).inner_text().strip() for i in range(opts.count()) if opts.nth(i).inner_text().strip()]

def click_search_or_query(ctx: Page | Frame):
    """검색/조회 버튼 클릭(둘 중 있는 쪽)"""
    btn = ctx.get_by_role("button", name="검색")
    if btn.count(): btn.first.click(); return
    btn = ctx.get_by_role("button", name="조회")
    if btn.count(): btn.first.click(); return
    cand = ctx.locator("input[type=button][value='검색'], input[type=button][value='조회']")
    if cand.count(): cand.first.click(); return
    raise RuntimeError("[셀렉터] '검색/조회' 버튼을 찾지 못했습니다.")

def locate_result_table(ctx: Page | Frame) -> Locator:
    """헤더 키워드로 가장 그럴듯한 테이블 선택"""
    tables = ctx.locator("table")
    best, best_score = None, -1
    for i in range(min(20, tables.count())):
        t = tables.nth(i)
        head_text = t.locator("thead").inner_text(timeout=1000) if t.locator("thead").count() else ""
        score = sum(1 for kw in ["기업", "신청", "유형", "주소", "종목"] if kw in head_text)
        if score > best_score:
            best, best_score = t, score
    return best if best is not None else tables.first

def extract_rows_from_table(table: Locator) -> List[Dict]:
    """테이블에서 기본 5열 구조로 파싱(NO/기업명/유형/주소/종목)"""
    rows: List[Dict] = []
    trs = table.locator("tbody tr")
    for i in range(trs.count()):
        tds = trs.nth(i).locator("td")
        c = tds.count()
        if c == 0: continue
        safe = lambda idx: tds.nth(idx).inner_text().strip() if idx < c else ""
        rows.append({
            "NO": safe(0),
            "학습기업명": safe(1),
            "참여신청유형": safe(2),
            "주소": safe(3),
            "종목": safe(4),
        })
    return rows

def click_next_if_possible(ctx: Page | Frame) -> bool:
    nxt = ctx.get_by_role("link", name="다음")
    if nxt.count() > 0 and nxt.first.is_enabled():
        nxt.first.click(); return True
    return False


# ----------------- 메인 -----------------
def run():
    all_rows: List[Dict] = []

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True, args=["--no-sandbox"])
        context = browser.new_context(locale="ko-KR")
        page = context.new_page()

        page.goto(URL, timeout=120_000)
        page.wait_for_load_state("domcontentloaded")

        ctx_form = find_target_frame(page)

        region_cb = get_combobox_by_label(ctx_form, "지역")
        branch_cb = get_combobox_by_label(ctx_form, "지사")
        # '참여유형명' / '참여유형' 어느 쪽이든 시도
        try:
            type_cb = get_combobox_by_label(ctx_form, "참여유형명")
        except Exception:
            type_cb = get_combobox_by_label(ctx_form, "참여유형")

        # 초기 옵션 준비
        wait_options_loaded(region_cb, 2, 10000)
        print("[옵션목록] 지역:", list_options_text(region_cb))

        for region in REGIONS:
            # 지역 선택
            region_cb.select_option(label=region)
            time.sleep(0.4)

            # 지역 선택 후 '지사' 옵션 로딩
            wait_options_loaded(branch_cb, 2, 10000)
            branch_opts = [x for x in list_options_text(branch_cb) if x != "선택"]
            print(f"[옵션목록] 지사({region}):", branch_opts)

            # 참여유형명 옵션(항상 동적 읽기)
            wait_options_loaded(type_cb, 2, 10000)
            type_opts = [x for x in list_options_text(type_cb) if x != "선택"]
            print(f"[옵션목록] 참여유형명({region}):", type_opts)

            for branch in branch_opts:
                for typ in type_opts:
                    # 지사/유형 선택
                    branch_cb.select_option(label=branch); time.sleep(0.2)
                    try:
                        type_cb.select_option(label=typ)
                    except Exception:
                        print(f"[경고] 유형 '{typ}' 선택 불가 → 건너뜀"); continue

                    # 검색
                    click_search_or_query(ctx_form)
                    page.wait_for_load_state("networkidle"); time.sleep(0.6)

                    # 첫 페이지
                    table = locate_result_table(ctx_form)
                    rows = extract_rows_from_table(table)
                    for r in rows:
                        r.update({"지역": region, "지사": branch, "참여유형명": typ})
                    all_rows.extend(rows)

                    # 페이지네이션
                    page_count = 1
                    while page_count < MAX_PAGES_PER_COMBO:
                        if not click_next_if_possible(ctx_form):
                            break
                        page.wait_for_load_state("networkidle"); time.sleep(0.4)
                        table = locate_result_table(ctx_form)
                        rows = extract_rows_from_table(table)
                        for r in rows:
                            r.update({"지역": region, "지사": branch, "참여유형명": typ})
                        all_rows.extend(rows)
                        page_count += 1

        browser.close()

    df = pd.DataFrame(all_rows).drop_duplicates()
    df.to_csv("pdms_learning_companies.csv", index=False, encoding="utf-8-sig")
    print(f"✅ Saved: pdms_learning_companies.csv (rows={len(df)})")


if __name__ == "__main__":
    try:
        run()
    except Exception as e:
        print("❌ ERROR:", e, file=sys.stderr)
        raise
