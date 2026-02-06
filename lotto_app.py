# -*- coding: utf-8 -*-
"""
로또 회차별 당첨번호 기반 추천 웹 앱
- 구간별 번호 선택: 1~10, 11~20, 21~30, 31~40, 41~45
- 번호 클릭 시 해당 번호와 함께 가장 많이 나온 번호 추천 (실제 당첨 데이터 기반)

실행: pip install flask pandas openpyxl 후, python lotto_app.py
      브라우저에서 http://127.0.0.1:5000 접속
"""
import os
import warnings
from collections import defaultdict

# openpyxl 기본 스타일 경고 무시
warnings.filterwarnings("ignore", message="Workbook contains no default style", module="openpyxl")

import pandas as pd
from flask import Flask, send_from_directory, request, jsonify

app = Flask(__name__, static_folder="static")

# 프로젝트 폴더 기준 데이터 로드 (파일명 패턴으로 검색)
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
def find_lotto_excel():
    for f in os.listdir(BASE_DIR):
        if f.startswith("로또") and f.endswith(".xlsx"):
            return os.path.join(BASE_DIR, f)
    return os.path.join(BASE_DIR, "로또 회차별 당첨번호_20260206104150.xlsx")
EXCEL_PATH = find_lotto_excel()

# 전역: 번호별 출현 횟수, (i,j) 동시 출현 횟수
FREQ = None
COOCCUR = None
DRAWS = None


def get_number_columns(df):
    """당첨번호 6개 컬럼 찾기 (다양한 엑셀 형식 대응)"""
    # 번호1~번호6 형태
    num_cols = [c for c in df.columns if isinstance(c, str) and "번호" in c and c.replace("번호", "").strip().isdigit()]
    if num_cols:
        num_cols.sort(key=lambda x: int(x.replace("번호", "").strip()))
        return num_cols[:6] if len(num_cols) >= 6 else None
    # 1,2,3,4,5,6 컬럼명 (숫자 또는 문자열)
    for start in [1, "1", 0]:
        cand = [start + i if isinstance(start, int) else str(int(start) + i) for i in range(6)]
        if all(c in df.columns for c in cand):
            return cand
    # 처음 6개 숫자형 컬럼
    numeric = [c for c in df.columns if str(getattr(df[c].dtype, "name", df[c].dtype)) in ("int64", "float64")]
    if len(numeric) >= 6:
        return list(numeric)[:6]
    # 컬럼 인덱스 2~7 (0:회차, 1:날짜 가정)
    if len(df.columns) >= 8:
        return list(df.columns[2:8])
    # 맨 앞 6개 컬럼
    if len(df.columns) >= 6:
        return list(df.columns[:6])
    return None


def _parse_draws_from_df(df, num_cols):
    """DataFrame과 번호 컬럼 리스트로 당첨 번호 6개씩 추출"""
    draws = []
    for _, row in df.iterrows():
        nums = []
        for c in num_cols:
            try:
                v = row[c] if c in row.index else row.get(c)
            except Exception:
                v = None
            if pd.isna(v):
                continue
            try:
                n = int(float(v))
                if 1 <= n <= 45:
                    nums.append(n)
            except (ValueError, TypeError):
                continue
        if len(nums) >= 6:
            draws.append([int(x) for x in nums[:6]])
    return draws


def _parse_draws_by_column_index(df, start_col=2, end_col=8):
    """컬럼 인덱스로 번호 추출 (0:회차, 1:날짜, 2~7:번호1~6 가정)"""
    draws = []
    cols = list(df.columns)[start_col:end_col]
    if len(cols) < 6:
        return draws
    for _, row in df.iterrows():
        nums = []
        for c in cols:
            try:
                v = row[c]
                if pd.notna(v):
                    n = int(float(v))
                    if 1 <= n <= 45:
                        nums.append(n)
            except (ValueError, TypeError):
                pass
        if len(nums) >= 6:
            draws.append([int(x) for x in nums[:6]])
    return draws


def load_lotto_data():
    global FREQ, COOCCUR, DRAWS
    draws = []

    # 1) 기본: 첫 행을 헤더로 읽기
    df = pd.read_excel(EXCEL_PATH)
    num_cols = get_number_columns(df)
    if num_cols is not None:
        draws = _parse_draws_from_df(df, num_cols)
    if not draws and len(df.columns) >= 6:
        draws = _parse_draws_from_df(df, list(df.columns[:6]))

    # 2) 0건이면 헤더를 1행으로 다시 시도 (첫 행이 제목인 경우)
    if not draws:
        df2 = pd.read_excel(EXCEL_PATH, header=1)
        num_cols2 = get_number_columns(df2)
        if num_cols2 is not None:
            draws = _parse_draws_from_df(df2, num_cols2)
        if not draws and len(df2.columns) >= 6:
            draws = _parse_draws_from_df(df2, list(df2.columns[:6]))

    # 3) 0건이면 헤더 없이 읽고 컬럼 인덱스 2~7을 번호로 사용
    if not draws:
        df3 = pd.read_excel(EXCEL_PATH, header=None)
        if len(df3.columns) >= 8:
            for _, row in df3.iterrows():
                nums = []
                for idx in range(2, 8):
                    try:
                        v = row.iloc[idx]
                        if pd.notna(v):
                            n = int(float(v))
                            if 1 <= n <= 45:
                                nums.append(n)
                    except (ValueError, TypeError, IndexError):
                        pass
                if len(nums) >= 6:
                    draws.append([int(x) for x in nums[:6]])
        if not draws and len(df3.columns) >= 6:
            for _, row in df3.iterrows():
                nums = []
                for idx in range(6):
                    try:
                        v = row.iloc[idx]
                        if pd.notna(v):
                            n = int(float(v))
                            if 1 <= n <= 45:
                                nums.append(n)
                    except (ValueError, TypeError, IndexError):
                        pass
                if len(nums) >= 6:
                    draws.append([int(x) for x in nums[:6]])

    # 4) 0건이면 첫 행 제외하고 header=None 데이터에서 2~7열 사용
    if not draws:
        df4 = pd.read_excel(EXCEL_PATH, header=None)
        if len(df4.columns) >= 8:
            for i in range(1, len(df4)):  # 첫 행(헤더) 스킵
                row = df4.iloc[i]
                nums = []
                for idx in range(2, 8):
                    try:
                        v = row.iloc[idx]
                        if pd.notna(v):
                            n = int(float(v))
                            if 1 <= n <= 45:
                                nums.append(n)
                    except (ValueError, TypeError, IndexError):
                        pass
                if len(nums) >= 6:
                    draws.append([int(x) for x in nums[:6]])
        # 4-2) 번호가 1~6열(0-based)에 있는 경우
        if not draws and len(df4.columns) >= 6:
            for i in range(1, len(df4)):
                row = df4.iloc[i]
                nums = []
                for idx in range(6):
                    try:
                        v = row.iloc[idx]
                        if pd.notna(v):
                            n = int(float(v))
                            if 1 <= n <= 45:
                                nums.append(n)
                    except (ValueError, TypeError, IndexError):
                        pass
                if len(nums) >= 6:
                    draws.append([int(x) for x in nums[:6]])

    # 디버그: 여전히 0건이면 엑셀 구조를 파일로 저장
    if not draws:
        try:
            df_debug = pd.read_excel(EXCEL_PATH, header=None)
            debug_path = os.path.join(BASE_DIR, "lotto_excel_debug.txt")
            with open(debug_path, "w", encoding="utf-8") as f:
                f.write("columns: " + str(list(df_debug.columns)) + "\n")
                f.write("shape: " + str(df_debug.shape) + "\n")
                f.write("first 5 rows:\n")
                f.write(df_debug.head().to_string() + "\n")
            print(f"데이터 0건. 엑셀 구조 확인용: {debug_path}")
        except Exception as e:
            print(f"디버그 파일 저장 실패: {e}")

    DRAWS = draws
    # 출현 빈도
    FREQ = defaultdict(int)
    for draw in DRAWS:
        for n in draw:
            FREQ[n] += 1
    # 동시 출현 (같은 회차에 함께 나온 횟수)
    COOCCUR = defaultdict(lambda: defaultdict(int))
    for draw in DRAWS:
        for i in range(len(draw)):
            for j in range(len(draw)):
                if i != j:
                    a, b = draw[i], draw[j]
                    COOCCUR[a][b] += 1
    return len(DRAWS)


def get_range(n):
    if 1 <= n <= 10:
        return "1~10"
    if 11 <= n <= 20:
        return "11~20"
    if 21 <= n <= 30:
        return "21~30"
    if 31 <= n <= 40:
        return "31~40"
    if 41 <= n <= 45:
        return "41~45"
    return None


# 앱 시작 시 데이터 로드
def init_data():
    global DRAWS, FREQ, COOCCUR
    try:
        n_draws = load_lotto_data()
        print(f"로또 데이터 로드 완료: {n_draws}회차")
        return n_draws
    except Exception as e:
        import traceback
        print(f"데이터 로드 실패: {e}")
        traceback.print_exc()
        DRAWS = []
        FREQ = defaultdict(int)
        COOCCUR = defaultdict(lambda: defaultdict(int))
        return 0

with app.app_context():
    init_data()


@app.route("/")
def index():
    return send_from_directory("static", "index.html")


@app.route("/api/stats")
def api_stats():
    """전체 번호별 출현 횟수 (구간별 정렬용)"""
    if FREQ is None:
        return jsonify({"error": "데이터 없음"}), 500
    by_range = {
        "1~10": [{"num": n, "count": FREQ[n]} for n in range(1, 11)],
        "11~20": [{"num": n, "count": FREQ[n]} for n in range(11, 21)],
        "21~30": [{"num": n, "count": FREQ[n]} for n in range(21, 31)],
        "31~40": [{"num": n, "count": FREQ[n]} for n in range(31, 41)],
        "41~45": [{"num": n, "count": FREQ[n]} for n in range(41, 46)],
    }
    return jsonify({"by_range": by_range, "total_draws": len(DRAWS) if DRAWS else 0})


@app.route("/api/recommend")
def api_recommend():
    """선택한 번호와 함께 가장 많이 나온 번호 추천"""
    try:
        num = request.args.get("number", type=int)
        if num is None or num < 1 or num > 45:
            return jsonify({"error": "1~45 사이 번호를 넣어주세요."}), 400
        if COOCCUR is None or DRAWS is None:
            return jsonify({"error": "데이터 없음"}), 500
        # 해당 번호와 함께 가장 많이 나온 번호 (해당 번호 제외)
        pairs = [(other, int(COOCCUR[num][other])) for other in range(1, 46) if other != num and COOCCUR[num][other] > 0]
        pairs.sort(key=lambda x: -x[1])
        top = pairs[:5]
        # 같은 구간 내 1순위
        r = get_range(num)
        in_range = [(n, c) for n, c in top if get_range(n) == r]
        return jsonify({
            "selected": int(num),
            "range": r,
            "recommendations": [{"number": int(n), "count": int(c)} for n, c in top],
            "top_in_same_range": [{"number": int(n), "count": int(c)} for n, c in in_range[:3]] if in_range else [],
            "total_draws": len(DRAWS) if DRAWS else 0,
        })
    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({"error": f"서버 오류: {str(e)}"}), 500


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)
