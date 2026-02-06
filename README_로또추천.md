# 로또 회차별 당첨번호 추천

당첨 이력(1~1209회차)을 바탕으로, 선택한 번호와 함께 가장 자주 나온 번호를 추천하는 웹 페이지입니다.

---

## 1) 로컬에서 실행 (노트북/PC)

- **필요**: Python, 엑셀 파일 `로또 회차별 당첨번호_*.xlsx` (같은 폴더에 두기)
- 터미널에서:
  ```bash
  pip install flask pandas openpyxl
  python lotto_app.py
  ```
- 브라우저에서 **http://127.0.0.1:5000** 접속

---

## 2) GitHub에 올려서 “사이트 들어가면 바로 사용” 하기 (GitHub Pages)

GitHub은 **코드만** 저장합니다. Python 서버는 돌리지 않으므로,  
**미리 계산한 데이터를 JSON으로 만들어 두고**, 그걸 웹에서만 쓰면 **별도 서버 없이** 사이트처럼 사용할 수 있습니다.

### 한 번만 할 작업 (데이터 만들기)

1. 프로젝트 폴더에 **로또 당첨번호 엑셀 파일**을 넣어 둡니다.
2. 터미널에서 한 번 실행:
   ```bash
   pip install pandas openpyxl
   python build_static_data.py
   ```
3. 이렇게 하면 **`docs/`** 폴더 안에 `data.json`, `index.html`이 생성됩니다.

### GitHub에 올리고 Pages 켜기

1. 이 프로젝트를 **GitHub 저장소**에 올립니다 (예: `내아이디/lotto-recommend`).
2. GitHub 저장소 페이지에서 **Settings → Pages** 로 이동합니다.
3. **Source** 를 다음처럼 설정합니다.
   - **Deploy from a branch**
   - **Branch**: `main` (또는 사용 중인 기본 브랜치)
   - **Folder**: **/docs**
   - **Save** 클릭
4. 1~2분 지나면 아래 주소로 접속할 수 있습니다.  
   `https://내아이디.github.io/저장소이름/`  
   (예: `https://myuser.github.io/lotto-recommend/`)

이 주소로 들어가면 **노트북에서 실행하지 않아도** 그대로 사용할 수 있습니다.

### 데이터/디자인 수정 후 다시 배포

- 엑셀을 바꾸거나, `static/index.html` 을 수정한 뒤에는:
  1. `python build_static_data.py` 다시 실행
  2. `docs/` 안의 변경된 파일을 커밋 후 푸시  
  → GitHub Pages가 자동으로 다시 배포합니다.

---

## 파일 역할

| 파일/폴더 | 설명 |
|-----------|------|
| `lotto_app.py` | 로컬용 Flask 서버 (엑셀 읽어서 API 제공) |
| `build_static_data.py` | 엑셀 → `data.json` 생성 + `docs/` 에 배포용 파일 생성 |
| `static/index.html` | 웹 화면 (로컬·GitHub Pages 공용) |
| `docs/` | GitHub Pages 로 “사이트” 로 쓸 때 사용 (빌드 시 자동 생성) |

요약하면, **GitHub에 올린 뒤에는 `docs` 폴더만 있으면 서버 없이 사이트처럼 사용 가능**합니다.
