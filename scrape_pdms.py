# -*- coding: utf-8 -*-
"""
PDMS 학습기업 스크레이퍼 (안정화 버전)
- 지역만 고정, 지사/참여유형명은 화면에서 '동적으로 읽어 전체 순회'
- 라벨 의존 X: 옵션 텍스트로 select를 분류(지역/지사/유형)
- select_option 실패 시 locator.evaluate()로 JS fallback
- 검색/페이지네이션 뒤에도 재바인딩 로직 동일 적용
"""

import time, sys
from typing import List, Dict, Tuple
import pandas as pd
from playwright.sync_api import sync_playwright, Page, Frame, Locator

URL = "https://pdms.ncs.go.kr/cmn/pub/opr/retrieveOprLrnEntrprList.do"

# 처음엔 1~2개만 테스트 후 확장 권장
REGIONS = [
    "서울","경기","인천","강원","충북","충남",
    "대전","경북","대구","전북","경남","울산",
    "부산","광주","전남","제주","세종"
]

MAX_PAGES_PER_COMBO = 3   # 테스트 시 3 등으로 축소

# --- 유틸 ---
REGION_KEYWORDS = set(REGIONS)
TYPE_KEYWORDS = {
    "공동훈련센터형","고숙련마이스터","대학(2년/4년) 연계형","IPP","전문대 재학단계",
    "P-Tech","도제학교","단독기업","민간 자율형",
    "연계형 일학습병행(채용예정자)","연계형 일학습병행(재직자향상)","단기과정",
    "첨단산업 아카데미(전문대)","첨단산업 아카데미(IPP)","경력개발 고도화",
    "구직자 취업연계형","외국인유학생(연수)","특화대학"
}

def find_target_frame(page: Page) -> Page | Frame:
    if page.get_by_role("combobox").count() > 0:
        return page
    for f in page.frames:
        try:
            if f.get_by_role("combobox").count() > 0:
                return f
        except Exception:
            pass
    return page

def wait_options_loaded(select_el: Locator, min_count: int = 2, timeout_ms: int = 10000):
    select_el.wait_for(state="attached", timeout=timeout_ms)
    try: select_el.wait_for(state="visible", timeout=timeout_ms)
    except: pass
    start = time.time()
    while True:
        try:
            if select_el.locator("option").count() >= min_count:
                return
        except Exception:
            pass
        if (time.time() - start) * 1000 > timeout_ms:
            raise TimeoutError("[대기초과] 옵션 로딩 미완료")
        time.sleep(0.2)

def options_text(select_el: Locator) -> List[str]:
    opts = select_el.locator("option")
    out = []
    n = opts.count()
    for i in range(n):
        t = opts.nth(i).inner_text().strip()
        if t:
            out.append(t)
    return out

def classify_selects(ctx: Page | Frame) -> Tuple[Locator, Locator, Locator]:
    """
    현재 화면의 select 3개(지역/지사/유형)를 '옵션 내용'으로 식별
    - 지역: REGIONS의 키워드가 여러 개 포함
    - 유형: TYPE_KEYWORDS가 포함
    - 나머지: 지사
    """
    selects = ctx.locator("select")
    cnt = selects.count()
    if cnt < 3:
        # combobox 역할 기반으로 보조
        comboboxes = ctx.get_by_role("combobox")
        if comboboxes.count() >= 3:
            selects = comboboxes
            cnt = selects.count()

    region_sel = None
    type_sel = None
    branch_sel = None
    scores = []

    for i in range(cnt):
        sel = selects.nth(i)
        try:
            wait_options_loaded(sel, 2, 5000)
        except Exception:
            continue
        texts = options_text(sel)
        txt_set = set(texts)
        # 점수 계산
        region_score = len(REGION_KEYWORDS.intersection(txt_set))
        type_score = len(TYPE_KEYWORDS.intersection(txt_set))
        scores.append((i, region_score, type_score, texts))
    # 지역 = region_score 최댓값
    scores_sorted = sorted(scores, key=lambda x: (x[1], x[2]), reverse=True)
    if scores_sorted:
        region_idx = scores_sorted[0][0]
        region_sel = selects.nth(region_idx)
    # 유형 = type_score 최댓값(지역과 다른 것)
    type_candidates = sorted(scores, key=lambda x: x[2], reverse=True)
    for i, rs, ts, _ in type_candidates:
        if region_sel is not None and i == scores_sorted[0][0]:
            continue
        if ts > 0:
            type_sel = selects.nth(i); break
    # 지사 = 남은 것 중 하나
    for i in range(cnt):
        if region_sel is not None and i == scores_sorted[0][0]:
            continue
        if type_sel is not None and i == type_candidates[0][0]:
            continue
        branch_sel = selects.nth(i)
        break

    if not (region_sel and branch_sel and type_sel):
        raise RuntimeError("[식별오류] 지역/지사/유형 select를 모두 식별하지 못했습니다.")
    return region_sel, branch_sel, type_sel

def select_by_label_with_fallback(select_el: Locator, label_text: str):
    """기본 select_option → 실패 시 locator.evaluate()로 강제 선택"""
    try:
        select_el.select_option(label=label_text)
        return
    except Exception:
        select_el.evaluate(
            """
            (el, label) => {
              const opts = Array.from(el.options);
              const hit = opts.find(o => o.text.trim() === label);
              if (!hit) throw new Error("label not found: " + label);
              el.value = hit.value;
              el.dispatchEvent(new Event('change', {bubbles:true}));
            }
            """,
            label_text
        )

def click_search_or_query(ctx: Page | Frame):
    btn = ctx.get_by_role("button", name="검색")
    if btn.count(): btn.first.click(); return
    btn = ctx.get_by_role("button", name="조회")
    if btn.count(): btn.first.click(); return
    cand = ctx.locator("input[type=button][value='검색'], input[type=button][value='조회']")
    if cand.count(): cand.first.click(); return
    raise RuntimeError("[셀렉터] '검색/조회' 버튼을 찾지 못했습니다.")

def locate_result_table(ctx: Page | Frame) -> Locator:
    tables = ctx.locator("table")
    best, score_best = None, -1
    for i in range(min(20, tables.count())):
        t = tables.nth(i)
        head = t.locator("thead")
        head_text = head.inner_text(timeout=1000) if head.count() else ""
        score = sum(1 for kw in ["기업","신청","유형","주소","종목"] if kw in head_text)
        if score > score_best:
            best, score_best = t, score
    return best if best is not None else tables.first

def extract_rows_from_table(table: Locator) -> List[Dict]:
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

# --- 메인 ---
def run():
    all_rows: List[Dict] = []
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True, args=["--no-sandbox"])
        context = browser.new_context(locale="ko-KR")
        page = context.new_page()
        page.goto(URL, timeout=120_000)
        page.wait_for_load_state("domcontentloaded")

        # 최초 바인딩: 옵션 기반으로 구분
        ctx_form = find_target_frame(page)
        region_cb, branch_cb, type_cb = classify_selects(ctx_form)

        # 지역 옵션 확인
        wait_options_loaded(region_cb, 2, 10000)
        regions_on_site = [x for x in options_text(region_cb) if x != "선택"]
        print("[옵션목록] 지역:", regions_on_site)

        for region in REGIONS:
            if region not in regions_on_site:
                print(f"[스킵] 지역 '{region}' 없음"); continue

            # 매 루프마다 재바인딩 (옵션기반)
            ctx_form = find_target_frame(page)
            region_cb, branch_cb, type_cb = classify_selects(ctx_form)

            # 지역 선택
            select_by_label_with_fallback(region_cb, region)
            time.sleep(0.4)

            # 지사/유형 옵션 갱신 후 읽기
            wait_options_loaded(branch_cb, 2, 10000)
            branch_opts = [x for x in options_text(branch_cb) if x != "선택"]
            wait_options_loaded(type_cb, 2, 10000)
            type_opts = [x for x in options_text(type_cb) if x != "선택"]

            print(f"[옵션목록] 지사({region}):", branch_opts)
            print(f"[옵션목록] 참여유형명({region}):", type_opts)

            for branch in branch_opts:
                for typ in type_opts:
                    # 조합 시작 시 재바인딩
                    ctx_form = find_target_frame(page)
                    region_cb, branch_cb, type_cb = classify_selects(ctx_form)

                    select_by_label_with_fallback(region_cb, region); time.sleep(0.2)
                    wait_options_loaded(branch_cb, 2, 10000)
                    select_by_label_with_fallback(branch_cb, branch); time.sleep(0.2)
                    wait_options_loaded(type_cb, 2, 10000)
                    try:
                        select_by_label_with_fallback(type_cb, typ)
                    except Exception:
                        print(f"[경고] 유형 '{typ}' 선택 불가 → 스킵"); continue

                    # 검색
                    click_search_or_query(ctx_form)
                    page.wait_for_load_state("networkidle"); time.sleep(0.6)

                    # 테이블 파싱
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
