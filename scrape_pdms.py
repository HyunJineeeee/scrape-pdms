# -*- coding: utf-8 -*-
"""
PDMS 학습기업 스크레이퍼 (최종 안정화)
- 지역만 리스트로 주면, 지사/참여유형명은 화면에서 자동 읽어 '전체 순회'
- select는 우선 'title' 속성으로 직접 찾고, 안되면 '옵션 내용'으로 분류 후 폴백
- select_option 실패 시 locator.evaluate()으로 강제 선택(JS fallback)
- 검색/페이지네이션 뒤에도 매번 재바인딩
"""

import time, sys
from typing import List, Dict, Tuple, Optional
import pandas as pd
from playwright.sync_api import sync_playwright, Page, Frame, Locator

URL = "https://pdms.ncs.go.kr/cmn/pub/opr/retrieveOprLrnEntrprList.do"

# 테스트는 1~2개만 두고 돌려보고, 정상 확인 후 전체로 확대하세요.
REGIONS = ["서울", "경기", "인천", "강원", "충북", "충남", "대전", "경북", "대구", "전북", "경남", "울산", "부산", "광주", "전남", "제주", "세종"]

MAX_PAGES_PER_COMBO = 3  # 테스트 시 3 정도로 줄이세요.

REGION_KEYWORDS = set(REGIONS)
TYPE_KEYWORDS = {
    "공동훈련센터형","고숙련마이스터","대학(2년/4년) 연계형","IPP","전문대 재학단계",
    "P-Tech","도제학교","단독기업","민간 자율형",
    "연계형 일학습병행(채용예정자)","연계형 일학습병행(재직자향상)","단기과정",
    "첨단산업 아카데미(전문대)","첨단산업 아카데미(IPP)","경력개발 고도화",
    "구직자 취업연계형","외국인유학생(연수)","특화대학"
}

# ---------- 공통 유틸 ----------
def find_target_frame(page: Page) -> Page | Frame:
    # combobox 또는 select가 가장 많이 보이는 프레임을 선택
    candidates: List[Page|Frame] = [page] + page.frames
    best_ctx = page
    best_score = -1
    for ctx in candidates:
        try:
            score = ctx.locator("select").count()
            if score > best_score:
                best_ctx, best_score = ctx, score
        except Exception:
            continue
    return best_ctx

def wait_options_loaded(select_el: Locator, min_count: int = 2, timeout_ms: int = 10000):
    select_el.wait_for(state="attached", timeout=timeout_ms)
    try:
        select_el.wait_for(state="visible", timeout=timeout_ms)
    except:
        pass
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
        try:
            t = opts.nth(i).inner_text().strip()
            if t:
                out.append(t)
        except Exception:
            continue
    return out

def select_by_label_with_fallback(select_el: Locator, label_text: str):
    # 기본 시도
    try:
        select_el.select_option(label=label_text)
        return
    except Exception:
        # JS fallback
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
    if btn.count():
        btn.first.click(); return
    btn = ctx.get_by_role("button", name="조회")
    if btn.count():
        btn.first.click(); return
    alt = ctx.locator("input[type=button][value='검색'], input[type=button][value='조회']")
    if alt.count():
        alt.first.click(); return
    raise RuntimeError("[셀렉터] '검색/조회' 버튼을 찾지 못했습니다.")

def locate_result_table(ctx: Page | Frame) -> Locator:
    tables = ctx.locator("table")
    best, score_best = None, -1
    count = min(30, tables.count())
    for i in range(count):
        t = tables.nth(i)
        head = t.locator("thead")
        head_text = head.inner_text(timeout=500) if head.count() else ""
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

# ---------- 셀렉트 바인딩 (title 우선 + 옵션분류 폴백) ----------
def query_select_by_title(ctx: Page | Frame, keyword: str) -> Optional[Locator]:
    # title 속성에 '지역', '지사', '참여유형명/참여유형' 이 포함된 select 우선
    loc = ctx.locator(f"select[title*='{keyword}']")
    if loc.count():
        return loc.first
    # aria-label 사용 가능 시
    loc = ctx.locator(f"[aria-label*='{keyword}']")
    if loc.count():
        return loc.first
    return None

def classify_selects(ctx: Page | Frame) -> Tuple[Locator, Locator, Locator]:
    # 1) title 기반 우선 탐색
    region = query_select_by_title(ctx, "지역")
    branch = query_select_by_title(ctx, "지사")
    typ = query_select_by_title(ctx, "참여유형명") or query_select_by_title(ctx, "참여유형")
    if region and branch and typ:
        return region, branch, typ

    # 2) 폴백: 옵션 내용으로 분류
    selects = ctx.locator("select")
    cnt = selects.count()
    if cnt == 0:
        raise RuntimeError("[식별오류] select 요소가 없습니다.")

    cand = []
    for i in range(min(20, cnt)):
        sel = selects.nth(i)
        try:
            wait_options_loaded(sel, 2, 3000)
            texts = options_text(sel)
            cand.append((sel, texts))
        except Exception:
            continue

    region_sel = None
    type_sel = None
    branch_sel = None
    region_score_best = -1
    type_score_best = -1

    for sel, texts in cand:
        sset = set(texts)
        rscore = len(REGION_KEYWORDS.intersection(sset))
        tscore = len(TYPE_KEYWORDS.intersection(sset))
        if rscore > region_score_best:
            region_score_best = rscore; region_sel = sel
        if tscore > type_score_best:
            type_score_best = tscore; type_sel = sel

    # 남는 하나를 지사로
    for sel, _ in cand:
        if sel != region_sel and sel != type_sel:
            branch_sel = sel
            break

    if not (region_sel and branch_sel and type_sel):
        raise RuntimeError("[식별오류] 지역/지사/유형 select를 모두 식별하지 못했습니다.")
    return region_sel, branch_sel, type_sel

# ---------- 메인 ----------
def run():
    all_rows: List[Dict] = []

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True, args=["--no-sandbox"])
        context = browser.new_context(locale="ko-KR")
        page = context.new_page()
        page.goto(URL, timeout=120_000)
        page.wait_for_load_state("domcontentloaded")

        # 최초 바인딩
        ctx_form = find_target_frame(page)
        region_cb, branch_cb, type_cb = classify_selects(ctx_form)

        # 지역 옵션 확인
        wait_options_loaded(region_cb, 2, 10000)
        regions_on_site = [x for x in options_text(region_cb) if x != "선택"]
        print("[옵션목록] 지역:", regions_on_site)

        for region in REGIONS:
            if region not in regions_on_site:
                print(f"[스킵] 지역 '{region}' 없음"); continue

            # 재바인딩(매 루프)
            ctx_form = find_target_frame(page)
            region_cb, branch_cb, type_cb = classify_selects(ctx_form)

            # 지역 선택
            select_by_label_with_fallback(region_cb, region)
            time.sleep(0.4)

            # 지사/유형 옵션 갱신 후 확보
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

                    click_search_or_query(ctx_form)
                    page.wait_for_load_state("networkidle"); time.sleep(0.6)

                    table = locate_result_table(ctx_form)
                    rows = extract_rows_from_table(table)
                    for r in rows:
                        r.update({"지역": region, "지사": branch, "참여유형명": typ})
                    all_rows.extend(rows)

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
