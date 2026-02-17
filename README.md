# record_archive

기록(현재는 텍스트 형태)을 생성/분류/검색/편집하고, 동시에 Word 파일로도 저장하는 데스크톱 프로그램입니다.

## 주요 기능

- 기록 생성
  - 제목/본문 입력
  - 생성 시점 날짜 자동 저장
  - 1차 Attribute 필수, 2차/3차 Attribute 선택
- 계층형 Attribute
  - 1차: 독립 분류
  - 2차: 선택한 1차 소속
  - 3차: 선택한 2차 소속 (결과적으로 상위 1차에도 포함)
  - 기존 Attribute 선택 + 신규 Attribute 즉시 생성 지원
- 기록 검색
  - Attribute / 날짜 / 제목 / 본문 기준 검색
  - 검색 전 기준 선택 가능
  - 선택하지 않으면 전체 기준(모든 항목)으로 검색
- 기록 편집
  - 기존 기록 선택 후 편집 저장
- Word(.docx) 동시 저장
  - 기본 경로: `D:\기록 프로젝트`
  - 경로 변경 가능
  - `저장경로\YYYY-MM-DD\YYYY-MM-DD: 제목.docx` 형태로 저장

## 실행 방법 (개발 모드)

```bash
python -m pip install -r requirements.txt
python app.py
```

## EXE 빌드/실행 (Windows)

### 1) 빌드 + 실행 (권장)

```bat
build_exe.bat
```

- `build_exe.bat`은 **의존성 설치 + EXE 빌드 후 자동 실행**합니다.
- 창이 자동으로 안 떠도 `dist\record_archive.exe`가 생성되었는지 먼저 확인하세요.

### 2) 이미 빌드된 EXE만 실행

```bat
run_app.bat
```

### 3) 수동 빌드

```bash
python -m pip install -r requirements.txt
pyinstaller --noconfirm --onefile --windowed --name record_archive app.py
```

빌드 결과물은 `dist/record_archive.exe`에 생성됩니다.

## 저장 데이터

- 로컬 DB: `records.db` (SQLite)
- 설정 파일: `config.json` (Word 저장 경로)
