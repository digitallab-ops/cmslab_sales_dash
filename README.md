# CMS Lab 매출 대시보드 — 인수인계 문서

> 내부 매출 데이터 시각화 및 관리 시스템. Excel 업로드 → 자동 파싱 → 팀/채널별 대시보드 제공.

---

## 시스템 구성

| 구분 | 내용 |
|---|---|
| **서버** | FastAPI + Uvicorn (Render 배포) |
| **DB** | Supabase PostgreSQL (`sales_dashboard` 스키마) |
| **프론트** | Jinja2 서버사이드 렌더링 (별도 빌드 없음) |
| **스케줄러** | APScheduler (자정 캐시 초기화) |
| **AI** | OpenAI GPT-4o-mini (어드민 AI 탭) |
| **MCP** | FastMCP Streamable HTTP (`/mcp` 엔드포인트) |

---

## 디렉토리 구조

```
sales_dashboard_web/
├── app/
│   ├── main.py              # FastAPI 앱 진입점, 마이그레이션
│   ├── models.py            # SQLAlchemy ORM 모델
│   ├── database.py          # DB 연결, 세션
│   ├── auth.py              # JWT 인증, 비밀번호 해시
│   ├── tab_registry.py      # 탭 목록 및 권한 헬퍼
│   ├── scheduler.py         # 자정 캐시 초기화 스케줄
│   ├── data/
│   │   └── parser.py        # Excel 파싱, HTML 생성, 캐시
│   ├── routes/
│   │   ├── auth_routes.py   # 로그인/로그아웃/회원가입
│   │   ├── dashboard.py     # 대시보드/비교 페이지
│   │   ├── admin.py         # 어드민 패널 (업로드, 사용자, 설정)
│   │   ├── api.py           # REST API (/api/v1/*)
│   │   ├── chat.py          # AI 챗봇 엔드포인트
│   │   └── mcp_gateway.py   # 통합 MCP 게이트웨이
│   └── templates/
│       └── admin.html       # 어드민 패널 UI
└── mcp_unified.py           # MCP 로컬 stdio 실행용 (Claude Desktop)
```

---

## Render 배포 설정

### 환경변수 (Render → Environment)

| 변수 | 설명 | 필수 |
|---|---|---|
| `DATABASE_URL` | Supabase Session Pooler URL | ✅ |
| `SECRET_KEY` | JWT 서명 키 (랜덤 문자열) | ✅ |
| `FIRST_ADMIN_EMAIL` | 최초 관리자 이메일 | ✅ |
| `FIRST_ADMIN_PASSWORD` | 최초 관리자 비밀번호 | ✅ |
| `OPENAI_API_KEY` | OpenAI API 키 (어드민 AI 탭용) | 선택 |
| `CHANNEL_MCP_API_KEY` | 채널 인사이트 MCP Bearer 토큰 | 선택 |
| `MCP_SECRET` | `/mcp` 엔드포인트 인증 토큰 | 선택 |

### Start Command
```
uvicorn app.main:app --host 0.0.0.0 --port $PORT
```

---

## 주요 기능

### 데이터 업로드
- 어드민 패널 → 데이터 관리 탭에서 Excel(.xlsx) 업로드
- 스냅샷 단위로 저장 (주차별 이력 보존)
- 활성 스냅샷이 대시보드에 표시됨

### 사용자/권한 관리
- **소속팀**: 특정 팀 데이터만 열람 가능 (NULL = 전체)
- **접근 탭**: 접근 허용 탭 목록 (NULL = 전체)
- **소속 그룹**: 그룹의 탭 권한을 상속
- 권한 우선순위: 개인 설정 → 그룹 기본값 → 전체 허용

### REST API (`/api/v1/*`)
- `GET /api/v1/summary` — 팀별 누적 실적 요약
- `GET /api/v1/teams/{team_name}` — 팀 월별 상세
- `GET /api/v1/snapshots` — 스냅샷 이력
- 인증: `X-API-Key` 헤더 (어드민 패널 → 시스템 설정에서 발급)

### 통합 MCP 게이트웨이 (`/mcp`)
- Sales 툴 3개 (내부 DB 직접 조회)
- 올리브영 툴 8개 / 쿠팡 4개 / 네이버 4개 (채널 MCP 프록시)
- 인증: `Authorization: Bearer {MCP_SECRET}`

---

## 로컬 개발

```bash
# 의존성 설치
pip install -r requirements.txt

# .env 파일 생성
DATABASE_URL=postgresql+psycopg2://...
SECRET_KEY=dev-secret-key
FIRST_ADMIN_EMAIL=admin@example.com
FIRST_ADMIN_PASSWORD=admin1234

# 실행
uvicorn app.main:app --reload
```

---

## DB 스키마 (`sales_dashboard` 스키마)

| 테이블 | 설명 |
|---|---|
| `users` | 사용자 계정, 권한, 탭/팀 접근 설정 |
| `teams` | 그룹 (탭 권한 기본값 포함) |
| `snapshots` | 업로드 이력 (주차별) |
| `sales_records` | 실적 데이터 (팀/월/실적/계획/전년) |
| `upload_history` | 파일 업로드 로그 |
| `app_config` | 앱 설정 (공지, API 키, 앱 이름 등) |

마이그레이션은 `app/main.py` `_run_migrations()`에서 자동 실행 (멱등).

---

## 변경 이력

### v1.0 — 초기 릴리스
- Excel(.xlsx) 업로드 → pandas 파싱 → 팀/월별 매출 대시보드 자동 생성
- JWT 기반 로그인·로그아웃, bcrypt 비밀번호 해시
- 관리자 전용 어드민 패널 (업로드, 사용자 관리, 공지 설정)
- Supabase PostgreSQL 연결 (`sales_dashboard` 스키마 분리)
- Render 자동 배포 구성

### v1.1 — 이력 관리 & 비교 기능
- **스냅샷 구조 도입**: 업로드마다 독립 스냅샷으로 저장, 주차 이력 전체 보존
- **매출현황(표) 페이지**: 주차 간 실적 비교 테이블 (이전 주차 대비 증감 표시)
- 업로드 시각 UTC→KST 변환 표시
- 스냅샷 전환 드롭다운 (네비바 상단)

### v1.2 — 권한 시스템
- **탭별 접근 권한**: 사용자별 허용 탭 목록 설정 (`tab_registry.py` 중앙 관리)
  - 라우트 레벨 차단 + UI 탭 링크 자동 숨김
- **그룹(팀) 권한 상속**: 팀 단위 탭 기본값 설정, 개인 오버라이드 가능
- **소속팀 데이터 필터**: 특정 팀 데이터만 열람하도록 쿼리 레벨 제한
- 어드민 패널에서 소속팀·탭 권한 인라인 수정 UI 추가

### v1.3 — 외부 연동
- **REST API** (`/api/v1/*`): API Key 인증 기반 외부 연동 엔드포인트
  - `/summary`, `/teams/{name}`, `/snapshots`
  - API Key 어드민 패널에서 발급·재발급
- **OpenAI 챗봇** (어드민 AI 탭): 현재 실적 데이터를 컨텍스트로 GPT-4o 질의응답
  - 대시보드 플로팅 위젯은 구현 후 임시 비활성화 (추후 재오픈 예정)

### v1.4 — 통합 MCP 게이트웨이
- **`/mcp` 엔드포인트**: FastMCP Streamable HTTP 기반 MCP 서버
  - Sales 툴 3개: 내부 DB 직접 조회
  - 채널 툴 16개: 올리브영(8) · 쿠팡(4) · 네이버(4) → 채널 인사이트 MCP 프록시
- Bearer 토큰 인증 미들웨어 (`MCP_SECRET` 환경변수)
- Claude Desktop / Claude Code에서 단일 MCP로 통합 분석 가능

---

## 문의
- 개발: 임승민 (lsmlub99@cms-lab.co.kr)
