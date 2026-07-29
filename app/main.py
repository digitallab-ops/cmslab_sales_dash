import os
from contextlib import asynccontextmanager
from fastapi import FastAPI, Request
from fastapi.responses import RedirectResponse, JSONResponse
from dotenv import load_dotenv

load_dotenv()

from .database import Base, engine
from .routes import auth_routes, dashboard, admin, api, chat
from .routes.mcp_gateway import get_mcp_app
from .models import User
from .auth import hash_password
from .scheduler import start_scheduler


def _migrate_schema():
    """public 스키마 테이블을 sales_dashboard 스키마로 이동 (최초 1회, 이후 멱등).

    - documents 등 public에 남아있는 기존 테이블은 건드리지 않음
    - 이미 sales_dashboard에 있거나 public에 없으면 조용히 스킵
    """
    from .database import SessionLocal, SCHEMA
    from sqlalchemy import text
    db = SessionLocal()
    tables = ["users", "snapshots", "sales_records", "upload_history", "teams", "app_config"]
    try:
        db.execute(text(f"CREATE SCHEMA IF NOT EXISTS {SCHEMA}"))
        db.commit()
        for t in tables:
            try:
                db.execute(text(f"ALTER TABLE public.{t} SET SCHEMA {SCHEMA}"))
                db.commit()
                print(f"[Schema] public.{t} → {SCHEMA}.{t}")
            except Exception:
                db.rollback()   # 이미 이동됐거나 public에 없으면 무시
        print(f"[Schema] {SCHEMA} 준비 완료")
    except Exception as e:
        print(f"[Schema] 오류: {e}")
        db.rollback()
    finally:
        db.close()


def _run_migrations():
    """기존 테이블에 누락된 컬럼을 추가한다 (멱등 실행 가능)."""
    from .database import SessionLocal, SCHEMA
    from sqlalchemy import text
    db = SessionLocal()
    try:
        db.execute(text(
            f"ALTER TABLE {SCHEMA}.users ADD COLUMN IF NOT EXISTS email_verified BOOLEAN DEFAULT FALSE"
        ))
        db.execute(text(
            f"ALTER TABLE {SCHEMA}.users ADD COLUMN IF NOT EXISTS verification_token VARCHAR(100)"
        ))
        db.execute(text(
            f"ALTER TABLE {SCHEMA}.users ADD COLUMN IF NOT EXISTS allowed_tabs TEXT"
        ))
        db.execute(text(
            f"ALTER TABLE {SCHEMA}.users ADD COLUMN IF NOT EXISTS group_team_id INTEGER"
        ))
        db.execute(text(
            f"ALTER TABLE {SCHEMA}.teams ADD COLUMN IF NOT EXISTS allowed_tabs TEXT"
        ))
        # 탭별 권한(tab_perms) 컬럼
        db.execute(text(
            f"ALTER TABLE {SCHEMA}.users ADD COLUMN IF NOT EXISTS tab_perms TEXT"
        ))
        db.execute(text(
            f"ALTER TABLE {SCHEMA}.teams ADD COLUMN IF NOT EXISTS tab_perms TEXT"
        ))
        # 관리자는 이메일 인증 없이 바로 로그인 가능해야 함
        db.execute(text(
            f"UPDATE {SCHEMA}.users SET email_verified = TRUE WHERE role = 'admin'"
        ))
        # 탭별 권한 기본값 정책(v2): 명시 설정 없으면 대시보드=전 팀, 그 외=본인 팀.
        # 과거 자동 백필분을 일회성으로 비워 스마트 기본값을 따르게 함(관리자 수동 설정은 이후 보존).
        marker = db.execute(text(
            f"SELECT value FROM {SCHEMA}.app_config WHERE key='perms_default_v2'"
        )).first()
        if not marker:
            db.execute(text(f"UPDATE {SCHEMA}.users SET tab_perms = NULL"))
            db.execute(text(f"UPDATE {SCHEMA}.teams SET tab_perms = NULL"))
            db.execute(text(
                f"INSERT INTO {SCHEMA}.app_config(key, value) VALUES('perms_default_v2','done')"
            ))
        db.commit()
        print("[Migration] 완료")
    except Exception as e:
        print(f"[Migration] 오류: {e}")
        db.rollback()
    finally:
        db.close()


def _create_first_admin():
    """FIRST_ADMIN_EMAIL로 관리자 계정 생성 또는 업그레이드.
    DELETE_USER_EMAIL이 설정된 경우 해당 계정을 먼저 삭제한다.
    """
    from .database import SessionLocal
    db = SessionLocal()
    try:
        # DELETE_USER_EMAIL: 지정한 계정 삭제 (테스트용 계정 정리 등)
        delete_email = os.getenv("DELETE_USER_EMAIL", "").strip()
        if delete_email:
            target = db.query(User).filter(User.email == delete_email).first()
            if target:
                db.delete(target)
                db.commit()
                print(f"[Init] 계정 삭제: {delete_email}")

        email = os.getenv("FIRST_ADMIN_EMAIL", "admin@cms-lab.co.kr").strip()
        password = os.getenv("FIRST_ADMIN_PASSWORD", "").strip()
        if not password:
            return

        existing = db.query(User).filter(User.email == email).first()
        if existing:
            # 이미 있으면 관리자 권한 확인 후 업그레이드
            if existing.role != "admin":
                existing.role = "admin"
                existing.is_active = True
                existing.email_verified = True
                existing.hashed_password = hash_password(password)
                db.commit()
                print(f"[Init] 기존 계정 관리자로 업그레이드: {email}")
        else:
            db.add(User(
                email=email,
                hashed_password=hash_password(password),
                name="관리자",
                role="admin",
                allowed_teams=None,
                is_active=True,
                email_verified=True,
            ))
            db.commit()
            print(f"[Init] 관리자 계정 생성: {email}")
    finally:
        db.close()


@asynccontextmanager
async def lifespan(app: FastAPI):
    _migrate_schema()                      # public → sales_dashboard 이동 (최초 1회)
    Base.metadata.create_all(bind=engine)  # sales_dashboard 스키마에 누락 테이블 생성
    _run_migrations()
    _create_first_admin()
    scheduler = start_scheduler()
    yield
    scheduler.shutdown(wait=False)


app = FastAPI(title="CMS Lab 매출 대시보드", lifespan=lifespan)

app.include_router(auth_routes.router)
app.include_router(dashboard.router)
app.include_router(admin.router)
app.include_router(api.router)
app.include_router(chat.router)

@app.middleware("http")
async def mcp_auth_middleware(request: Request, call_next):
    if request.url.path.startswith("/mcp"):
        secret = os.getenv("MCP_SECRET", "").strip()
        if secret:
            auth = request.headers.get("Authorization", "")
            if auth != f"Bearer {secret}":
                return JSONResponse({"error": "Unauthorized"}, status_code=401)
    return await call_next(request)

app.mount("/mcp", get_mcp_app())


@app.get("/")
async def root():
    return RedirectResponse("/dashboard")


@app.api_route("/health", methods=["GET", "HEAD"])
async def health():
    return {"status": "ok"}