# -*- coding: utf-8 -*-
"""
PDMS 학습기업 스크레이퍼 (안정 재바인딩)
- 지역만 고정, 지사/참여유형명은 페이지에서 자동 읽어 전체 순회
- 초기에 각 select의 id/name/인덱스를 '시그니처'로 저장 → 이후 재바인딩은 id/name 우선
- 라벨이 잠시 사라져도 동작
- select_option 실패 시 JS fallback
"""

import time, sys
from typing import List, Dict, Optional, Tuple
import pandas as pd
from playwright.sync_api import sync_playwright, Page, Frame, Locator

URL = "https://pdms.ncs.go.kr/cmn/pub/opr/retrieveOprLrnEntrprList.do"

REGIONS = [
    "서울","경기","인천","강원","충북","충남",
    "대전","경북","대구","전북","경남","울산",
    "부산","광주","전남","제주","세종"
]

MAX_PAGES_PER_COMBO = 3   # 테스트 시 3 등으로 축소

# ----------------- 공통 유틸 -----------------
def find_target_frame(page: Page) -> Page | Frame:
    # combobox가 보이는 프레임을 찾음
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
            raise TimeoutError("[대기초과] 옵션 로딩이 완료되지 않았습니다.")
        time.sleep(0.2)

def list_options_text(select_el: Locator) -> list[str]:
    opts = select_el.locator("option")
    out = []
    for i in range(opts.count()):
        t = opts.nth(i).inner_text().strip()
        if t: out.append(t)
    return out

def select_by_label_with_fallback(ctx: Page | Frame, select_el: Locator, label_text: str):
    try:
        select_el.select_option(label=label_text)
        return
    except Exception:
        # JS fallback (label 매칭)
        ctx.evaluate(
            """
            (select, label) => {
              const options = Array.from(select.options);
              const hit = options.find(o => o.text.trim() === label);
              if (!hit) throw new Error("label not found: " + label);
              select.value = hit.value;
              select.dispatchEvent(new Event('change', {bubbles:true}));
            }
            """,
            select_el, label_text
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
    best, best_score = None, -1
    for i in range(min(20, tables.count())):
        t = tables.nth(i)
        head_text = t.locator("thead").inner_text(timeout=1000) if t.locator("thead").count() else ""
        score = sum(1 for kw in ["기업","신청","유형","주소","종목"] if kw in head_text)
        if score > best_score:
            best, best_score = t, score
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

# ----------------- 시그니처(안정 재바인딩 핵심) -----------------
class SelectSignature:
    """select를 안정적으로 다시 찾기 위한 id/name/index 시그니처"""
    def __init__(self, css_by_id: Optional[str], css_by_name: Optional[str], index: int):
        self.css_by_id = css_by_id
        self.css_by_name = css_by_name
        self.index = index   # 같은 폼 내 select 순서(최후 폴백)

    def query(self, ctx: Page | Frame) -> Locator:
        if self.css_by_id:
            loc = ctx.locator(self.css_by_id)
            if loc.count(): return loc.first
        if self.css_by_name:
            loc = ctx.locator(self.css_by_name)
            if loc.count(): return loc.first
        # 마지막 폴백: 해당 폼 안의 select N번째
        selects = ctx.locator("select")
        if selects.count() > self.index:
            return selects.nth(self.index)
        # 정말 안되면 combobox로라도
        cbs = ctx.get_by_role("combobox")
        if cbs.count() > self.index:
            return cbs.nth(self.index)
        raise RuntimeError("select rebinding failed (id/name/index 모두 실패)")

def make_signature(ctx: Page | Frame, el: Locator) -> SelectSignature:
    # id / name 추출
    el_id = el.get_attribute("id")
    el_name = el.get_attribute("name")
    css_by_id = f"select#{el_id}" if el_id else None
    css_by_name = f"select[name='{el_name}']" if el_name else None
    # index 추정: 같은 부모(form/fieldset) 내 select 순서
    parent = el.locator("xpath=..")
    selects = parent.locator("xpath=.//select")
    index = 0
    try:
        total = selects.count()
        for i in range(total):
            if selects.nth(i).evaluate("e => e === this", arg=el.element_handle()):
                index = i
                break
    except Exception:
        # 부모 기준 실패 시 페이지 전체 기준으로
        selects = ctx.locator("select")
        total = selects.count()
        for i in range(total):
            try:
                if selects.nth(i).evaluate("e => e === this", arg=el.element_handle()):
                    index = i; break
            except Exception:
                pass
    return SelectSignature(css_by_id, css_by_name, index)

# ----------------- 메인 -----------------
def run():
    all_rows: List[Dict] = []

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True, args=["--no-sandbox"])
        context = browser.new_context(locale="ko-KR")
        page = context.new_page()
        page.goto(URL, timeout=120_000)
        page.wait_for_load_state("domcontentloaded")

        # 최초 프레임 및 콤보 바인딩(라벨 사용)
        ctx_form = find_target_frame(page)

        # 라벨명이 '참여유형명' 또는 '참여유형'일 수 있음 → 둘 다 시도
        def get_type_combo(ctx) -> Locator:
            cb = ctx.get_by_role("combobox", name="참여유형명")
            if cb.count(): return cb.first
            cb = ctx.get_by_role("combobox", name="참여유형")
            if cb.count(): return cb.first
            # 라벨 실패 시 첫 3개 select 중 세 번째가 유형일 가능성 높음(지역,지사 다음)
            return ctx.locator("select").nth(2)

        region_cb = ctx_form.get_by_role("combobox", name="지역").first \
                    if ctx_form.get_by_role("combobox", name="지역").count() else ctx_form.locator("select").nth(0)
        branch_cb = ctx_form.get_by_role("combobox", name="지사").first \
                    if ctx_form.get_by_role("combobox", name="지사").count() else ctx_form.locator("select").nth(1)
        type_cb = get_type_combo(ctx_form)

        # 시그니처 기록(이게 핵심!)
        sig_region = make_signature(ctx_form, region_cb)
        sig_branch = make_signature(ctx_form, branch_cb)
        sig_type   = make_signature(ctx_form, type_cb)

        # 지역 옵션 확인
        wait_options_loaded(region_cb, 2, 10000)
        regions_on_site = [x for x in list_options_text(region_cb) if x != "선택"]
        print("[옵션목록] 지역:", regions_on_site)

        # 지역 루프
        for region in REGIONS:
            if region not in regions_on_site:
                print(f"[스킵] 지역 '{region}' 옵션 없음"); continue

            # 재바인딩(라벨 대신 시그니처)
            ctx_form = find_target_frame(page)
            region_cb = sig_region.query(ctx_form)
            branch_cb = sig_branch.query(ctx_form)
            type_cb   = sig_type.query(ctx_form)

            # 지역 선택
            select_by_label_with_fallback(ctx_form, region_cb, region)
            time.sleep(0.4)

            # 지사/유형 옵션 로드
            wait_options_loaded(branch_cb, 2, 10000)
            branch_opts = [x for x in list_options_text(branch_cb) if x != "선택"]

            wait_options_loaded(type_cb, 2, 10000)
            type_opts = [x for x in list_options_text(type_cb) if x != "선택"]

            print(f"[옵션목록] 지사({region}):", branch_opts)
            print(f"[옵션목록] 참여유형명({region}):", type_opts)

            for branch in branch_opts:
                for typ in type_opts:
                    # 매 조합마다 재바인딩(시그니처 기반)
                    ctx_form = find_target_frame(page)
                    region_cb = sig_region.query(ctx_form)
                    branch_cb = sig_branch.query(ctx_form)
                    type_cb   = sig_type.query(ctx_form)

                    # 지역 다시 지정(의존성 유지)
                    select_by_label_with_fallback(ctx_form, region_cb, region)
                    time.sleep(0.2)

                    # 지사/유형 선택
                    wait_options_loaded(branch_cb, 2, 10000)
                    select_by_label_with_fallback(ctx_form, branch_cb, branch)
                    time.sleep(0.2)

                    wait_options_loaded(type_cb, 2, 10000)
                    try:
                        select_by_label_with_fallback(ctx_form, type_cb, typ)
                    except Exception:
                        print(f"[경고] 유형 '{typ}' 선택 불가 → 건너뜀"); continue

                    # 검색
                    click_search_or_query(ctx_form)
                    page.wait_for_load_state("networkidle"); time.sleep(0.6)

                    # 결과 수집
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
