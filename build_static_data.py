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

# 프로젝트 루트를 작업 디렉터리로
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
os.chdir(BASE_DIR)

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
    if os.path.exists(index_src):
        with open(index_src, "r", encoding="utf-8") as f:
            with open(index_dst, "w", encoding="utf-8") as g:
                g.write(f.read())
        print(f"저장 완료: {static_path}, {docs_data_path}, {index_dst} ({total}회차)")
    else:
        print(f"저장 완료: {static_path}, {docs_data_path} ({total}회차)")
    return total

if __name__ == "__main__":
    main()
