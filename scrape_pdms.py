# -*- coding: utf-8 -*-
"""
PDMS 학습기업 목록 스크레이퍼 (Playwright)
- 라벨 텍스트 기반으로 combobox(select) 안정 선택
- iframe 대응
- 지역 선택 후 지사 옵션 로딩 대기
- 페이지네이션 처리
- CSV 저장

사용법:
  python scrape_pdms.py
"""

import time
import sys
from typing import List, Dict, Optional
import pandas as pd
from playwright.sync_api import sync_playwright, Page, Frame, Locator

URL = "https://pdms.ncs.go.kr/cmn/pub/opr/retrieveOprLrnEntrprList.do"

# ▶ 필요한 필터만 넣으세요 (라벨 표시 문자열과 동일하게)
TARGETS = [
    {"region": "서울", "branch": "서울강남", "type": "공동훈련센터형"},
    # {"region": "경기", "branch": "경기남부", "type": "공동훈련센터형"},
]

# ===== Helper functions =====

def find_target_frame(page: Page) -> Page | Frame:
    """
    폼이 iframe 안에 있을 가능성 대비: 메인 페이지에서 못 찾으면 프레임들 순회
    """
    # 우선 메인에서 combobox를 찾아본다
    if page.get_by_role("combobox").count() > 0:
        return page
    # 프레임들 검사
    for f in page.frames:
        try:
            if f.get_by_role("combobox").count() > 0:
                return f
        except Exception:
            pass
    return page  # 그래도 없으면 일단 page 반환

def get_combobox_by_label(ctx: Page | Frame, label_text: str) -> Locator:
    """
    접근성 라벨 이름으로 combobox(select) 찾기
    ※ 라벨 텍스트가 정확히 일치해야 함 (페이지에 따라 "참여유형" vs "참여유형명" 차이 있을 수 있음)
    """
    cb = ctx.get_by_role("combobox", name=label_text)
    if cb.count() == 0:
        # 라벨-셀렉트 형제 관계로 탐색 (대안)
        label = ctx.locator(f"label:has-text('{label_text}')").first
        if label.count() == 0:
            raise RuntimeError(f"라벨 '{label_text}' 을(를) 찾지 못했습니다. 실제 라벨 문구를 확인해 주세요.")
        # for 속성으로 연결
        for_id = label.get_attribute("for")
        if for_id:
            sel = ctx.locator(f"select#{for_id}")
        else:
            # following-sibling::select fallback
            sel = label.locator("xpath=following-sibling::select[1]")
        if sel.count() == 0:
            raise RuntimeError(f"라벨 '{label_text}' 에 연결된 select를 찾지 못했습니다.")
        return sel
    return cb.first

def wait_options_loaded(select_el: Locator, min_count: int = 2, timeout_ms: int = 10000):
    """
    선택 박스의 옵션이 동적으로 채워질 때까지 대기
    (예: 지역 선택 후 지사 옵션 채워짐)
    """
    select_el.wait_for(state="attached", timeout=timeout_ms)
    # 옵션 수가 min_count 이상이 될 때까지 폴링
    start = time.time()
    while True:
        try:
            opt_count = select_el.locator("option").count()
            if opt_count >= min_count:
                return
        except Exception:
            pass
        if (time.time() - start) * 1000 > timeout_ms:
            raise TimeoutError("옵션 로딩 대기 시간 초과")
        time.sleep(0.2)

def extract_table_rows(ctx: Page | Frame) -> List[Dict]:
    """
    결과 테이블에서 행 파싱
    - 셀렉터는 페이지 구조에 따라 한번 확인 필요
    """
    rows: List[Dict] = []
    # 테이블 찾기: 접근성 텍스트/캡션이 없다면 가장 큰 결과 테이블을 지정
    # 아래는 예시 셀렉터 (상황에 맞게 한 번만 조정)
    table = ctx.locator("table").filter(has_text="학습기업").first
    if table.count() == 0:
        # fallback: 첫 번째 tbody 기준
        table = ctx.locator("table").first

    trs = table.locator("tbody tr")
    n = trs.count()
    for i in range(n):
        tds = trs.nth(i).locator("td")
        tdn = tds.count()
        if tdn == 0:
            continue
        # 컬럼 순서는 사이트에 맞춰 1~2회 조정 필요
        # 예: NO | 기업명 | 신청유형 | 주소 | 종목
        def safe(idx: int) -> str:
            return tds.nth(idx).inner_text().strip() if idx < tdn else ""

        row = {
            "NO": safe(0),
            "학습기업명": safe(1),
            "참여신청유형": safe(2),
            "주소": safe(3),
            "종목": safe(4),
        }
        rows.append(row)
    return rows

def click_search(ctx: Page | Frame):
    # '검색' 버튼 클릭 (role 기반이 가장 안정적)
    btn = ctx.get_by_role("button", name="검색")
    if btn.count() == 0:
        # fallback: input[type=button] with value=검색
        btn = ctx.locator("input[type=button][value='검색']")
    btn.first.click()

def click_next_if_possible(ctx: Page | Frame) -> bool:
    """
    페이지네이션 '다음'이 있으면 클릭하고 True, 없으면 False
    """
    nxt = ctx.get_by_role("link", name="다음")
    if nxt.count() > 0 and nxt.first.is_enabled():
        nxt.first.click()
        return True
    # 대안: 다음 페이지 번호 찾기 (필요시 보완)
    return False

# ===== Main =====

def run():
    all_rows: List[Dict] = []
    with sync_playwright() as p:
        # headless=True가 Actions에서 빠름
        browser = p.chromium.launch(headless=True, args=["--no-sandbox"])
        ctx = browser.new_context(locale="ko-KR")
        page = ctx.new_page()
        page.goto(URL, timeout=120_000)
        page.wait_for_load_state("domcontentloaded")

        # 폼이 프레임 안인지 판별
        ctx_form = find_target_frame(page)

        # 콤보박스 레퍼런스
        # 라벨 문구는 실제 화면에 보이는 텍스트와 정확히 일치시켜야 함
        region_cb = get_combobox_by_label(ctx_form, "지역")
        branch_cb = get_combobox_by_label(ctx_form, "지사")
        # 사이트마다 '참여유형' 혹은 '참여유형명' 등으로 표기 다를 수 있음
        try:
            type_cb = get_combobox_by_label(ctx_form, "참여유형")
        except RuntimeError:
            type_cb = get_combobox_by_label(ctx_form, "참여유형명")

        # 셀렉트박스 옵션 로딩 대기(최초)
        wait_options_loaded(region_cb, min_count=2, timeout_ms=10000)
        wait_options_loaded(branch_cb, min_count=2, timeout_ms=10000)

        for t in TARGETS:
            # 지역 선택
            region_cb.select_option(label=t["region"])
            # 지역 선택 이후 지사 옵션이 비동기로 갱신될 수 있으므로 대기
            time.sleep(0.4)
            wait_options_loaded(branch_cb, min_count=2, timeout_ms=10000)

            # 지사 선택
            branch_cb.select_option(label=t["branch"])
            time.sleep(0.3)

            # 참여유형 선택
            type_cb.select_option(label=t["type"])
            time.sleep(0.2)

            # 검색
            click_search(ctx_form)
            # 결과 로딩 대기
            page.wait_for_load_state("networkidle")
            time.sleep(0.6)

            # 페이지네이션 루프
            while True:
                rows = extract_table_rows(ctx_form)
                for r in rows:
                    r.update({"지역": t["region"], "지사": t["branch"], "참여유형": t["type"]})
                all_rows.extend(rows)

                # 다음 페이지로
                has_next = False
                try:
                    has_next = click_next_if_possible(ctx_form)
                except Exception:
                    has_next = False
                if not has_next:
                    break
                page.wait_for_load_state("networkidle")
                time.sleep(0.4)

        browser.close()

    # CSV 저장
    df = pd.DataFrame(all_rows).drop_duplicates()
    df.to_csv("pdms_learning_companies.csv", index=False, encoding="utf-8-sig")
    print(f"✅ Saved: pdms_learning_companies.csv ({len(df)} rows)")

if __name__ == "__main__":
    try:
        run()
    except Exception as e:
        # 디버깅을 돕는 스크린샷 저장
        # (Actions에서 실패 시 아티팩트로 확인하도록)
        # 주: 페이지 핸들이 여기선 없을 수 있으니 pass
        print("❌ ERROR:", e, file=sys.stderr)
        raise
