# -*- coding: utf-8 -*-
"""
동행복권 API로 회차별 당첨번호를 자동 수집한 뒤 data.json + docs 를 갱신합니다.
수동 엑셀 없이, GitHub Actions 등에서 주기 실행해 완전 자동화할 수 있습니다.

사용법:
  pip install requests
  python update_lotto_auto.py
"""
import json
import os
import re
import time
from collections import defaultdict
from datetime import datetime

try:
    import requests
except ImportError:
    print("requests 필요: pip install requests")
    raise

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
os.chdir(BASE_DIR)

# 1회차 = 2002-12-07 (토), 매주 토요일 추첨
FIRST_DRAW_DATE = datetime(2002, 12, 7)
API_URL = "https://www.dhlottery.co.kr/common.do?method=getLottoNumber&drwNo={}"
# API 실패 시 사용할 엑셀 다운로드 URL (외부 정리 사이트)
EXCEL_FALLBACK_URL = "https://superkts.com/lotto/download_excel.php"


def get_latest_round():
    """오늘 기준 추정 최신 회차 (이미 추첨된 회차까지)"""
    now = datetime.now()
    days = (now - FIRST_DRAW_DATE).days
    # 토요일 추첨: 해당 주 토요일 지나야 회차 확정
    week = days // 7
    return max(1, week + 1)


def fetch_one_round(round_no):
    """한 회차 당첨번호 조회. 성공 시 [n1,n2,n3,n4,n5,n6], 실패 시 None"""
    try:
        r = requests.get(
            API_URL.format(round_no),
            headers={
                "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
                "Accept": "application/json, text/javascript, */*; q=0.01",
                "Accept-Language": "ko-KR,ko;q=0.9,en;q=0.8",
                "Referer": "https://www.dhlottery.co.kr/",
                "X-Requested-With": "XMLHttpRequest",
            },
            timeout=10,
        )
        r.raise_for_status()
        text = r.text.strip()
        if not text:
            return None
        # 1) JSON 파싱 시도
        try:
            data = r.json()
            if data.get("returnValue") == "fail":
                return None
            nums = []
            for i in range(1, 7):
                key = "drwtNo{}".format(i)
                if key not in data:
                    return None
                n = int(data[key])
                if 1 <= n <= 45:
                    nums.append(n)
            if len(nums) == 6:
                return nums
            return None
        except ValueError:
            pass
        # 2) HTML/텍스트 안에서 drwtNo1~6 숫자 추출 (따옴표 유무 모두)
        for pattern in [
            r'"drwtNo%d"\s*:\s*(\d+)',
            r"'drwtNo%d'\s*:\s*(\d+)",
            r"drwtNo%d\s*[=:]\s*(\d+)",
        ]:
            nums = []
            for i in range(1, 7):
                m = re.search(pattern % i, text)
                if m:
                    n = int(m.group(1))
                    if 1 <= n <= 45:
                        nums.append(n)
            if len(nums) == 6:
                return nums
    except Exception:
        pass
    return None


def fetch_draws_from_excel_url():
    """엑셀 다운로드 URL에서 당첨번호 수집. 성공 시 draws 리스트, 실패 시 []"""
    try:
        import io
        import pandas as pd
        import urllib3
        urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
        r = requests.get(
            EXCEL_FALLBACK_URL,
            headers={"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"},
            timeout=15,
            verify=False,
        )
        r.raise_for_status()
        if "application/vnd.openxmlformats" not in r.headers.get("Content-Type", "") and "octet-stream" not in r.headers.get("Content-Type", ""):
            if r.text.strip().startswith("{"):
                return []
        df = pd.read_excel(io.BytesIO(r.content))
        draws = []
        num_cols = None
        for c in df.columns:
            if isinstance(c, str) and "번호" in c and c.replace("번호", "").strip().isdigit():
                num_cols = sorted([x for x in df.columns if isinstance(x, str) and "번호" in x and x.replace("번호", "").strip().isdigit()], key=lambda x: int(x.replace("번호", "").strip()))[:6]
                break
        if not num_cols and len(df.columns) >= 6:
            num_cols = list(df.columns[:6])
        if not num_cols:
            return []
        for _, row in df.iterrows():
            nums = []
            for c in num_cols:
                try:
                    v = row.get(c) if c in row.index else row[c]
                    if pd.isna(v):
                        continue
                    n = int(float(v))
                    if 1 <= n <= 45:
                        nums.append(n)
                except (ValueError, TypeError, KeyError):
                    continue
            if len(nums) >= 6:
                draws.append([int(x) for x in nums[:6]])
        return draws
    except Exception as e:
        print("엑셀 폴백 실패:", e)
        return []


def builds_from_draws(draws):
    """draws 리스트로 freq, cooccur 계산 (lotto_app과 동일 구조)"""
    freq = defaultdict(int)
    cooccur = defaultdict(lambda: defaultdict(int))
    for draw in draws:
        for n in draw:
            freq[n] += 1
        for i in range(len(draw)):
            for j in range(len(draw)):
                if i != j:
                    a, b = draw[i], draw[j]
                    cooccur[a][b] += 1
    return freq, cooccur


def main():
    latest = get_latest_round()
    # 1회차 한 번 테스트 (실패 시 응답 샘플 저장)
    test = fetch_one_round(1)
    if test is None:
        try:
            r = requests.get(
                API_URL.format(1),
                headers={
                    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
                    "Accept": "application/json, text/html, */*",
                    "Referer": "https://www.dhlottery.co.kr/",
                },
                timeout=10,
            )
            debug_path = os.path.join(BASE_DIR, "lotto_api_debug.txt")
            with open(debug_path, "w", encoding="utf-8") as f:
                f.write("status=%s\n" % r.status_code)
                f.write("Content-Type=%s\n\n" % r.headers.get("Content-Type", ""))
                f.write(r.text[:3000])
            print("동행복권 API 응답 없음. 응답 샘플: %s" % debug_path)
        except Exception as e:
            print("동행복권 API 오류:", e)
        print("엑셀 다운로드 URL로 시도 중...")
        draws = fetch_draws_from_excel_url()
        if not draws:
            print("엑셀 폴백도 실패. 엑셀 파일을 직접 받아 build_static_data.py 를 사용하세요.")
            return 0
        print("엑셀에서 %d회차 수집됨. data.json 생성 중..." % len(draws))
        freq, cooccur = builds_from_draws(draws)
        total = len(draws)
        freq_dict = {str(n): freq[n] for n in range(1, 46)}
        cooccur_dict = {}
        for num in range(1, 46):
            cooccur_dict[str(num)] = {
                str(other): int(cooccur[num][other])
                for other in range(1, 46)
                if other != num and cooccur[num][other] > 0
            }
        data = {"totalDraws": total, "freq": freq_dict, "cooccur": cooccur_dict}
        static_dir = os.path.join(BASE_DIR, "static")
        docs_dir = os.path.join(BASE_DIR, "docs")
        os.makedirs(static_dir, exist_ok=True)
        os.makedirs(docs_dir, exist_ok=True)
        for folder in (static_dir, docs_dir):
            with open(os.path.join(folder, "data.json"), "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False)
        if os.path.exists(os.path.join(static_dir, "index.html")):
            with open(os.path.join(static_dir, "index.html"), "r", encoding="utf-8") as f:
                with open(os.path.join(docs_dir, "index.html"), "w", encoding="utf-8") as g:
                    g.write(f.read())
        print("저장 완료: %d회차 → static/data.json, docs/data.json, docs/index.html" % total)
        return total
    print(f"1회차 테스트 OK: {test}")
    print(f"1 ~ {latest} 회차 수집 중...")
    draws = []
    for no in range(1, latest + 1):
        row = fetch_one_round(no)
        if row:
            draws.append(row)
        if no % 100 == 0:
            print(f"  {no}회차까지 {len(draws)}건")
        time.sleep(0.05)
    if not draws:
        print("수집된 데이터가 없습니다. API 확인 또는 나중에 다시 시도하세요.")
        return 0
    freq, cooccur = builds_from_draws(draws)
    total = len(draws)
    freq_dict = {str(n): freq[n] for n in range(1, 46)}
    cooccur_dict = {}
    for num in range(1, 46):
        cooccur_dict[str(num)] = {
            str(other): int(cooccur[num][other])
            for other in range(1, 46)
            if other != num and cooccur[num][other] > 0
        }
    data = {"totalDraws": total, "freq": freq_dict, "cooccur": cooccur_dict}
    static_dir = os.path.join(BASE_DIR, "static")
    docs_dir = os.path.join(BASE_DIR, "docs")
    os.makedirs(static_dir, exist_ok=True)
    os.makedirs(docs_dir, exist_ok=True)
    for folder, name in [(static_dir, "static"), (docs_dir, "docs")]:
        path = os.path.join(folder, "data.json")
        with open(path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False)
    index_src = os.path.join(static_dir, "index.html")
    index_dst = os.path.join(docs_dir, "index.html")
    if os.path.exists(index_src):
        with open(index_src, "r", encoding="utf-8") as f:
            with open(index_dst, "w", encoding="utf-8") as g:
                g.write(f.read())
    print(f"저장 완료: {total}회차 → static/data.json, docs/data.json, docs/index.html")
    return total


if __name__ == "__main__":
    main()
