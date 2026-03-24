# -*- coding: utf-8 -*-
"""
로또 엑셀 데이터를 읽어 static/data.json 으로 저장합니다.
GitHub Pages 등 정적 호스팅에서 서버 없이 사용할 수 있도록 합니다.

사용법: 프로젝트 폴더에서 실행
  python build_static_data.py

실행 전에 '로또 회차별 당첨번호_*.xlsx' 파일이 같은 폴더에 있어야 합니다.
"""
import json
import os
import re

# 프로젝트 루트를 작업 디렉터리로
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
os.chdir(BASE_DIR)

def fetch_full_draw_info(round_no):
    """특정 회차의 당첨번호 6개와 보너스 번호를 엑셀 파일에서 직접 가져옵니다. {round, numbers, bonus}"""
    try:
        import lotto_app
        import pandas as pd
        
        if not lotto_app.EXCEL_PATH:
            return None
            
        df = pd.read_excel(lotto_app.EXCEL_PATH, header=None)
        
        # 엑셀 뒷부분부터 역순 검색하여 해당 회차(round_no) 찾기
        for i in range(len(df)-1, -1, -1):
            row = df.iloc[i]
            try:
                # 첫 번째 열이 회차 번호인지 확인
                val = row.iloc[0]
                if pd.notna(val) and int(float(val)) == round_no:
                    nums = []
                    bonus = None
                    # 보통 2열~7열이 번호1~6, 8열이 보너스 번호
                    for idx in range(2, 9):
                        v = row.iloc[idx]
                        if pd.notna(v):
                            n = int(float(v))
                            if 1 <= n <= 45:
                                nums.append(n)
                    
                    if len(nums) >= 7:
                        # 앞에서 6개가 당첨번호, 7번째가 보너스
                        return {
                            "round": round_no, 
                            "numbers": nums[:6], 
                            "bonus": nums[6]
                        }
            except (ValueError, TypeError, IndexError):
                continue
    except Exception as e:
        print(f"엑셀에서 보너스 번호를 읽는 중 오류 발생: {e}")
    return None

# lotto_app 로드 시 데이터 로드까지 실행됨
import lotto_app

def main():
    total = len(lotto_app.DRAWS) if lotto_app.DRAWS else 0
    if total == 0:
        print("경고: 당첨 데이터가 0건입니다. 엑셀 파일 경로와 형식을 확인하세요.")
    # JSON은 키가 문자열이어야 함
    freq = {str(n): lotto_app.FREQ[n] for n in range(1, 46)}
    cooccur = {}
    for num in range(1, 46):
        cooccur[str(num)] = {
            str(other): int(lotto_app.COOCCUR[num][other])
            for other in range(1, 46)
            if other != num and lotto_app.COOCCUR[num][other] > 0
        }
    data = {
        "totalDraws": total,
        "freq": freq,
        "cooccur": cooccur,
    }
    
    # 최신 회차 보너스 번호 포함 데이터 가져오기 (API 호출)
    if total > 0:
        latest_draw_info = fetch_full_draw_info(total)
        if latest_draw_info:
            data["latest_draw"] = latest_draw_info
            
    # 로컬 Flask용
    static_dir = os.path.join(BASE_DIR, "static")
    os.makedirs(static_dir, exist_ok=True)
    static_path = os.path.join(static_dir, "data.json")
    with open(static_path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False)
    # GitHub Pages용 (docs 폴더)
    docs_dir = os.path.join(BASE_DIR, "docs")
    os.makedirs(docs_dir, exist_ok=True)
    docs_data_path = os.path.join(docs_dir, "data.json")
    with open(docs_data_path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False)
    index_src = os.path.join(static_dir, "index.html")
    index_dst = os.path.join(docs_dir, "index.html")
    index_src = os.path.join(static_dir, "index.html")
    index_dst = os.path.join(docs_dir, "index.html")
    if os.path.exists(index_src):
        with open(index_src, "r", encoding="utf-8") as f:
            html_content = f.read()
            
        # JS에 직접 데이터 하드코딩 (file:/// 로컬 파일 실행 지원용)
        json_data_str = json.dumps(data, ensure_ascii=False)
        
        # 중복된 추가 renderLatestDraw 코드 블록이 있다면 제거 (꼬임 방지)
        # static/index.html 끝 부분에 임의로 추가된 코드가 있다면 삭제 처리
        if "// ── 최신 당첨 번호 렌더링 ──────────────────────────────" in html_content:
            html_content = re.sub(
                r"// ── 최신 당첨 번호 렌더링 ──────────────────────────────.*?// ──────────────────────────────────────────────────────\s*", 
                "", 
                html_content, 
                flags=re.DOTALL
            )

        # 원래의 fetch 로직을 하드코딩된 변수로 교체
        fetch_regex = r"let lottoData = null;\s*fetch\('data\.json(\?t=[^']+)?'\)\s*\.then\([^)]+\)\s*\.then\(data => {\s*lottoData = data;\s*renderLatestDraw\(\);\s*}\)\s*\.catch\([^)]+\);"
        
        if re.search(fetch_regex, html_content):
            html_content = re.sub(fetch_regex, f"let lottoData = {json_data_str};\n    renderLatestDraw();", html_content)
        elif "let lottoData = null;" in html_content:
            # fallback
            html_content = html_content.replace(
                "let lottoData = null;", 
                f"let lottoData = {json_data_str};"
            )
            
        with open(index_dst, "w", encoding="utf-8") as g:
            g.write(html_content)
        print(f"저장 완료: {static_path}, {docs_data_path}, {index_dst} ({total}회차)")
    else:
        print(f"저장 완료: {static_path}, {docs_data_path} ({total}회차)")
    return total

if __name__ == "__main__":
    main()
