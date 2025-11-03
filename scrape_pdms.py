# -*- coding: utf-8 -*-
"""
PDMS 학습기업 스크레이퍼
- 지역만 고정, 지사/참여유형명은 페이지에서 자동 읽어 전체 순회
- 콤보박스는 매번 '재조회'하여 detached locator 문제 해결
- visible/enable 대기 + JS fallback 선택
- 검색/페이지네이션 후에도 프레임/폼/콤보 재바인딩
"""

import time, sys
from typing import List, Dict
import pandas as pd
from playwright.sync_api import sync_playwright, Page, Frame, Locator

URL = "https://pdms.ncs.go.kr/cmn/pub/opr/retrieveOprLrnEntrprList.do"

REGIONS = [
    "서울","경기","인천","강원","충북","충남",
    "대전","경북","대구","전북","경남","울산",
    "부산","광주","전남","제주","세종"
]

MAX_PAGES_PER_COMBO = 9999   # 테스트 시 3 등으로 축소

# ---------- 공통 유틸 ----------
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

def get_combobox_by_label(ctx: Page | Frame, label_text: str) -> Locator:
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
    select_el.wait_for(state="attached", timeout=timeout_ms)
    # 보임/활성 대기
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
    """기본 select_option 시도 → 실패 시 JS로 label 매칭 선택"""
    try:
        select_el.select_option(label=label_text)
        return
    except Exception:
        # JS fallback
        ctx.evaluate(
            """
            (select, label) => {
              const opts = Array.from(select.options);
              const hit = opts.find(o => o.text.trim() === label);
              if (!hit) throw new Error("label not found: "+label);
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

def rebind_controls(page: Page) -> tuple[Page|Frame, Locator, Locator, Locator]:
    """검색/렌더링 후 프레임과 콤보를 다시 바인딩"""
    ctx_form = find_target_frame(page)
    region_cb = get_combobox_by_label(ctx_form, "지역")
    branch_cb = get_combobox_by_label(ctx_form, "지사")
    try:
        type_cb = get_combobox_by_label(ctx_form, "참여유형명")
    except Exception:
        type_cb = get_combobox_by_label(ctx_form, "참여유형")
    return ctx_form, region_cb, branch_cb, type_cb

# ---------- 메인 ----------
def run():
    all_rows: List[Dict] = []

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True, args=["--no-sandbox"])
        context = browser.new_context(locale="ko-KR")
        page = context.new_page()
        page.goto(URL, timeout=120_000)
        page.wait_for_load_state("domcontentloaded")

        ctx_form, region_cb, branch_cb, type_cb = rebind_controls(page)

        # 지역 옵션 확인
        wait_options_loaded(region_cb, 2, 10000)
        regions_on_site = [x for x in list_options_text(region_cb) if x != "선택"]
        print("[옵션목록] 지역:", regions_on_site)

        for region in REGIONS:
            if region not in regions_on_site:
                print(f"[스킵] 지역 '{region}' 옵션 없음")
                continue

            # 매 회차 재바인딩(안정성)
            ctx_form, region_cb, branch_cb, type_cb = rebind_controls(page)

            # 지역 선택
            select_by_label_with_fallback(ctx_form, region_cb, region)
            time.sleep(0.4)

            # 지사/유형 옵션 새로 읽기
            wait_options_loaded(branch_cb, 2, 10000)
            branch_opts = [x for x in list_options_text(branch_cb) if x != "선택"]
            wait_options_loaded(type_cb, 2, 10000)
            type_opts = [x for x in list_options_text(type_cb) if x != "선택"]

            print(f"[옵션목록] 지사({region}):", branch_opts)
            print(f"[옵션목록] 참여유형명({region}):", type_opts)

            for branch in branch_opts:
                for typ in type_opts:
                    # 매 조합마다 재바인딩 (지사/유형 선택 전에)
                    ctx_form, region_cb, branch_cb, type_cb = rebind_controls(page)

                    # 지역 다시 설정(조합 반복 중 DOM이 바뀌었을 수 있음)
                    try:
                        select_by_label_with_fallback(ctx_form, region_cb, region)
                    except Exception:
                        # 프레임 체인지 대비 한번 더
                        ctx_form, region_cb, branch_cb, type_cb = rebind_controls(page)
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

                    # 결과 테이블
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
