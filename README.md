# CMS Lab Sales Dashboard

> B2B 화장품 브랜드의 채널별 매출 실적을 실시간으로 추적·분석하는 내부 대시보드.  
> Excel 업로드 → 자동 파싱 → 팀/채널/브랜드별 시각화까지 풀스택으로 직접 구현.

---

## 주요 기능

- **주차별 매출 스냅샷** — Excel 업로드 시 자동 파싱, 주차 이력 보존 및 비교
- **다차원 필터링** — 팀 / 연도 / 브랜드 / 채널 / 월별 자유 조합 조회
- **사용자 권한 시스템** — 탭별·팀별 접근 권한, 그룹 상속 구조
- **REST API** — 외부 연동용 API Key 인증 엔드포인트
- **통합 MCP 게이트웨이** — 매출 데이터 + 올리브영/쿠팡/네이버 채널 인사이트를 하나의 MCP 서버로 통합
- **AI 분석** — OpenAI GPT-4o 기반 매출 데이터 질의응답 (어드민 전용)

---

## 기술 스택

| 구분 | 기술 |
|---|---|
| Backend | Python 3.12 · FastAPI · SQLAlchemy 2.x |
| Database | PostgreSQL (Supabase) |
| Frontend | Jinja2 SSR · Bootstrap 5 · Vanilla JS |
| Auth | JWT (python-jose) · bcrypt |
| Infra | Render (PaaS) · GitHub Actions 없이 자동 배포 |
| AI | OpenAI API (gpt-4o-mini / gpt-4o) |
| MCP | FastMCP · Streamable HTTP transport |
| Data | pandas · openpyxl |

---

## 아키텍처

```
[사용자 브라우저]
      │ HTTPS
      ▼
[Render — FastAPI]
  ├── /dashboard, /compare   → Excel 파싱 HTML + 권한 필터
  ├── /admin                 → 업로드·사용자·설정 관리
  ├── /api/v1/*              → REST API (API Key 인증)
  ├── /chat                  → OpenAI 챗봇
  └── /mcp                   → 통합 MCP 게이트웨이
           │ Bearer Token
           ▼
  [채널 인사이트 MCP — Vercel]
  (올리브영 · 쿠팡 · 네이버 16개 툴)
      │
      ▼
[Supabase PostgreSQL]
  sales_dashboard 스키마
```

---

## 구현 포인트

### 스냅샷 기반 이력 관리
단순 덮어쓰기가 아닌 주차별 스냅샷으로 저장해 과거 데이터 조회 및 주차 간 비교가 가능하도록 설계.

### 탭·팀 이중 권한 시스템
- **소속팀**: DB 쿼리 레벨에서 팀 데이터 필터링 (열람 범위 제한)
- **접근 탭**: 라우트 레벨 차단 + UI 링크 자동 숨김
- **그룹 상속**: 팀 단위 기본값 설정 → 개인 오버라이드 가능
- 권한 해석 로직을 `tab_registry.py` 단일 모듈로 집중

### 통합 MCP 게이트웨이
사내 매출 데이터와 외부 채널 MCP(올리브영·쿠팡·네이버)를 하나의 엔드포인트로 통합.  
Claude Desktop / Claude Code에서 "올리브영 부정 리뷰랑 우리 팀 매출 같이 분석해줘" 같은 크로스 데이터 분석이 가능.

### 서버사이드 HTML 캐싱
pandas로 생성한 대시보드 HTML을 메모리 캐시에 보관해 반복 요청 시 파싱 오버헤드 제거.  
대용량 데이터셋은 집계 테이블로 자동 전환하는 임계값 로직 구현.

---

## 스크린샷

| 매출 대시보드 | 매출현황(표) | 어드민 패널 |
|---|---|---|
| 월별 추이 차트, KPI 카드, 팀별 실적 | 주차 비교 테이블 | 사용자·권한·업로드 관리 |

---

## 로컬 실행

```bash
git clone https://github.com/lsmlub99/cmslab_sales_dash.git
cd cmslab_sales_dash

pip install -r requirements.txt

# .env
DATABASE_URL=postgresql+psycopg2://...
SECRET_KEY=your-secret
FIRST_ADMIN_EMAIL=admin@example.com
FIRST_ADMIN_PASSWORD=admin1234

uvicorn app.main:app --reload
# → http://localhost:8000
```

---

## Claude Desktop MCP 연결

```json
{
  "mcpServers": {
    "cms-unified": {
      "type": "http",
      "url": "https://your-app.onrender.com/mcp",
      "headers": { "Authorization": "Bearer your-mcp-secret" }
    }
  }
}
```

---

## 개발 배경

사내에서 Excel 파일을 매주 수동으로 공유하던 매출 보고 프로세스를 자동화하기 위해 개발.  
단순 시각화를 넘어 채널별 인사이트 MCP와 연동해 AI가 직접 데이터를 조회·분석하는 구조까지 확장.
