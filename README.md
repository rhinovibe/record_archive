# record_archive

기록(현재는 텍스트 형태)을 생성/분류/검색/편집하고, 동시에 Word 파일로도 저장하는 데스크톱 프로그램입니다.

## 주요 기능

- 기록 생성
  - 제목/본문 입력
  - 생성 시점 날짜 자동 저장
  - 1차 Attribute 1개 이상 필수, 2차/3차 Attribute 선택
  - 하나의 기록에 여러 개의 1차/2차/3차 Attribute 동시 부여 가능(목록 다중 선택)
- 계층형 Attribute
  - 1차: 독립 분류
  - 2차: 선택한 1차 소속
  - 3차: 선택한 2차 소속 (결과적으로 상위 1차에도 포함)
  - 기존 Attribute 선택 + 신규 Attribute 즉시 생성 지원
  - 신규 Attribute 입력란은 쉼표(,)로 여러 개를 한 번에 추가 가능
- 기록 검색
  - Attribute / 날짜 / 제목 / 본문 기준 검색
  - 검색 전 기준 선택 가능
  - 선택하지 않으면 전체 기준(모든 항목)으로 검색
  - 검색은 부분 일치 기반이며, 공백으로 여러 키워드를 입력하면 각 키워드를 모두 포함하는 기록을 찾음
- 기록 편집
  - 기존 기록 선택 후 편집 저장
- 저장 안정성
  - 저장 시 DB 트랜잭션 커밋 후 즉시 재조회로 저장 성공 여부를 검증
  - 실패 시 현재 DB 경로와 오류를 팝업/상태줄에 표시
  - 성공 시에도 DB 저장 위치를 팝업과 상태줄에 명시 (워드 저장 이전 단계에서 즉시 표시)
  - 저장 성공 이력은 `<프로젝트>/save_audit.log`에 기록
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

### 2) run_app.bat 실행 (매번 빌드 후 실행)

```bat
run_app.bat
```

- `run_app.bat`은 내부에서 `build_exe.bat`를 먼저 호출한 뒤 최신 EXE를 실행합니다.

### 3) 수동 빌드

```bash
python -m pip install -r requirements.txt
pyinstaller --noconfirm --onefile --windowed --name record_archive app.py
```

빌드 결과물은 `dist/record_archive.exe`에 생성됩니다.

## 저장 데이터

- 로컬 DB:
  - Windows: `D:\기록 프로젝트\records.db` 우선 사용
  - Windows에서 D: 접근 실패 시: `%USERPROFILE%\record_archive\records.db` 폴백
  - 비-Windows 개발 환경: `<프로젝트 폴더>/records.db`
  - 구버전 `<프로젝트 폴더>/records.db`가 있으면 Windows 실행 시 `D:\기록 프로젝트\records.db`로 자동 복사
- 설정 파일: `config.json` (Word 저장 경로)
