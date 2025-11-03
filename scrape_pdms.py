import time
import pandas as pd
from playwright.sync_api import sync_playwright

URL = "https://pdms.ncs.go.kr/cmn/pub/opr/retrieveOprLrnEntrprList.do"

# 원하는 필터를 여기에 설정(예: 지역=서울, 지사=서울강남, 참여유형=공동훈련센터형)
TARGETS = [
    {"region": "서울", "branch": "서울강남", "type": "공동훈련센터형"},
    # {"region": "경기", "branch": "경기남부", "type": "공동훈련센터형"},
]

def extract_table_rows(page):
    # 페이지 테이블을 pandas로 만들기 쉬운 형태로 긁어옴
    rows = []
    # 테이블 헤더/바디의 CSS 셀렉터는 사이트 변경에 따라 바뀔 수 있음 → 필요 시 수정
    table = page.locator("table:has-text('학습기업 목록')")  # 페이지 내 '학습기업 목록' 테이블
    trs = table.locator("tbody tr")
    for i in range(trs.count()):
        tds = trs.nth(i).locator("td")
        if tds.count() < 5:
            continue
        no = tds.nth(0).inner_text().strip()
        name = tds.nth(1).inner_text().strip()
        apply_type = tds.nth(2).inner_text().strip()
        addr = tds.nth(3).inner_text().strip()
        category = tds.nth(4).inner_text().strip()
        rows.append({"NO": no, "학습기업명": name, "참여신청유형": apply_type, "주소": addr, "종목": category})
    return rows

def run():
    all_rows = []
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        ctx = browser.new_context(locale="ko-KR")
        page = ctx.new_page()
        page.goto(URL, timeout=120000)
        page.wait_for_load_state("domcontentloaded")

        for target in TARGETS:
            # 1) 지역 선택
            page.locator("select:has-text('지역')").select_option(label=target["region"])
            time.sleep(0.5)

            # 2) 지사 선택
            page.locator("select:has-text('지사')").select_option(label=target["branch"])
            time.sleep(0.5)

            # 3) 참여유형 선택
            page.locator("select:has-text('참여유형명')").select_option(label=target["type"])
            time.sleep(0.3)

            # 4) 검색 버튼 클릭
            # 버튼 텍스트가 '검색' 또는 아이콘일 수 있어 role 기반으로 선택
            page.get_by_role("button", name="검색").click()
            page.wait_for_load_state("networkidle")
            time.sleep(0.8)

            # 5) 페이지네이션 돌면서 수집
            while True:
                rows = extract_table_rows(page)
                # 필드에 현재 선택된 필터를 추가해서 함께 저장
                for r in rows:
                    r.update({"지역": target["region"], "지사": target["branch"], "참여유형": target["type"]})
                all_rows.extend(rows)

                # 다음 페이지가 있으면 클릭
                next_btn = page.get_by_role("link", name="다음")  # 페이징 텍스트가 '다음' 기준 (필요시 수정)
                if next_btn.count() == 0 or not next_btn.is_enabled():
                    break
                next_btn.click()
                page.wait_for_load_state("networkidle")
                time.sleep(0.5)

        browser.close()

    # CSV 저장
    df = pd.DataFrame(all_rows).drop_duplicates()
    df.to_csv("pdms_learning_companies.csv", index=False, encoding="utf-8-sig")
    print(f"Saved: pdms_learning_companies.csv ({len(df)} rows)")

if __name__ == "__main__":
    run()
