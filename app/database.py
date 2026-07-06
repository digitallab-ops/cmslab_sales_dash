import os
from sqlalchemy import create_engine, MetaData
from sqlalchemy.orm import sessionmaker, DeclarativeBase

DATABASE_URL = os.getenv(
    "DATABASE_URL",
    "postgresql://postgres:postgres@localhost:5432/sales_db"
)

# Render 등에서 postgres:// 로 시작하는 URL을 postgresql:// 로 변환
if DATABASE_URL.startswith("postgres://"):
    DATABASE_URL = DATABASE_URL.replace("postgres://", "postgresql://", 1)

engine = create_engine(
    DATABASE_URL,
    pool_pre_ping=True,
    pool_size=3,        # 기본 커넥션 수 (Supabase Session Pooler 한도 15)
    max_overflow=2,     # 초과 허용 커넥션 (최대 5개)
    pool_recycle=300,   # 5분마다 커넥션 재생성 (유휴 커넥션 방지)
    pool_timeout=10,    # 커넥션 대기 타임아웃
)
SessionLocal = sessionmaker(autocommit=False, autoflush=False, bind=engine)

# 앱 전용 스키마 — public.documents 등 기존 테이블과 분리
SCHEMA = "sales_dashboard"


class Base(DeclarativeBase):
    # MetaData에 schema 지정 → 모든 ORM DDL/DML이 sales_dashboard.tablename 사용
    metadata = MetaData(schema=SCHEMA)


def get_db():
    db = SessionLocal()
    try:
        yield db
    finally:
        db.close()
