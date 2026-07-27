import os, json, datetime
from sqlalchemy import Column, Integer, String, Numeric, Boolean, DateTime, ForeignKey, Text, TypeDecorator, Index, UniqueConstraint
from sqlalchemy.orm import relationship
from .database import Base

# PostgreSQL ARRAY 대체: SQLite 호환 JSON 직렬화 타입
class TextList(TypeDecorator):
    impl = Text
    cache_ok = True

    def process_bind_param(self, value, dialect):
        if value is None:
            return None
        return json.dumps(value, ensure_ascii=False)

    def process_result_value(self, value, dialect):
        if value is None:
            return None
        if isinstance(value, list):
            return value
        if isinstance(value, str):
            # JSON 배열 형식 ["팀A","팀B"]
            if value.startswith('['):
                return json.loads(value)
            # psycopg2가 TEXT 컬럼에서 그대로 반환한 PostgreSQL 배열 리터럴 {팀A,팀B}
            if value.startswith('{') and value.endswith('}'):
                inner = value[1:-1]
                return [s.strip().strip('"') for s in inner.split(',')] if inner else []
        return value


class User(Base):
    __tablename__ = "users"

    id = Column(Integer, primary_key=True, index=True)
    email = Column(String(100), unique=True, nullable=False, index=True)
    hashed_password = Column(String(255), nullable=False)
    name = Column(String(100), default="")
    role = Column(String(20), default="viewer")         # 'admin' | 'viewer'
    allowed_teams = Column(TextList, nullable=True)     # NULL = 전체 팀 열람
    allowed_tabs  = Column(TextList, nullable=True)     # NULL = 그룹 기본값 상속 또는 전체
    group_team_id = Column(Integer, ForeignKey("teams.id"), nullable=True)  # 탭 권한 상속 그룹
    is_active = Column(Boolean, default=True)
    email_verified = Column(Boolean, default=False)
    verification_token = Column(String(100), nullable=True)
    created_at = Column(DateTime, default=datetime.datetime.utcnow)

    snapshots  = relationship("Snapshot", back_populates="uploader")
    group_team = relationship("Team", foreign_keys="[User.group_team_id]")


class Snapshot(Base):
    __tablename__ = "snapshots"

    id = Column(Integer, primary_key=True, index=True)
    week_label = Column(String(100), default="")
    base_date = Column(String(50), default="")
    uploaded_by = Column(Integer, ForeignKey("users.id"), nullable=True)
    uploaded_at = Column(DateTime, default=datetime.datetime.utcnow)
    is_active = Column(Boolean, default=True)

    uploader = relationship("User", back_populates="snapshots")
    records = relationship(
        "SalesRecord", back_populates="snapshot",
        cascade="all, delete-orphan"
    )


class SalesRecord(Base):
    __tablename__ = "sales_records"

    id = Column(Integer, primary_key=True, index=True)
    snapshot_id = Column(
        Integer, ForeignKey("snapshots.id", ondelete="CASCADE"), nullable=False, index=True
    )
    team = Column(String(50))
    channel = Column(String(100))
    brand = Column(String(10))
    code = Column(String(20))
    month = Column(Integer)           # 1~12

    y2024 = Column(Numeric(15, 3), default=0)
    y2025b = Column(Numeric(15, 3), default=0)
    y2025 = Column(Numeric(15, 3), default=0)
    plan = Column(Numeric(15, 3), default=0)
    actual = Column(Numeric(15, 3), default=0)

    fw1 = Column(Numeric(15, 3), nullable=True)
    fw2 = Column(Numeric(15, 3), nullable=True)
    fw3 = Column(Numeric(15, 3), nullable=True)
    fw4 = Column(Numeric(15, 3), nullable=True)
    fw5 = Column(Numeric(15, 3), nullable=True)

    snapshot = relationship("Snapshot", back_populates="records")

    __table_args__ = (
        UniqueConstraint("snapshot_id", "team", "channel", "brand", "code", "month", name="uq_sr_key"),
        Index("ix_sr_snapshot_team", "snapshot_id", "team"),
        Index("ix_sr_snapshot_month", "snapshot_id", "month"),
    )


class UploadHistory(Base):
    __tablename__ = "upload_history"

    id = Column(Integer, primary_key=True, index=True)
    week_label = Column(String(100), default="")
    base_date = Column(String(50), default="")
    uploaded_by = Column(Integer, ForeignKey("users.id"), nullable=True)
    uploaded_at = Column(DateTime, default=datetime.datetime.utcnow)
    upserted_count = Column(Integer, default=0)


class Team(Base):
    __tablename__ = "teams"

    id = Column(Integer, primary_key=True, index=True)
    name = Column(String(100), unique=True, nullable=False)
    display_order = Column(Integer, default=0)
    is_active = Column(Boolean, default=True)
    allowed_tabs = Column(TextList, nullable=True)   # NULL = 이 그룹 소속 사용자는 전체 탭


class AppConfig(Base):
    __tablename__ = "app_config"

    key = Column(String(100), primary_key=True)
    value = Column(Text, nullable=True)


# ─── 품목별 매출 대시보드 (item dashboard) ───────────────────────────────────

class ItemSnapshot(Base):
    """품목 데이터 업로드 배치 단위. 원본(ItemRaw)과 집계(ItemRecord)를 묶는다."""
    __tablename__ = "item_snapshots"

    id = Column(Integer, primary_key=True, index=True)
    base_date = Column(String(50), default="")
    channel_order = Column(Text, nullable=True)   # JSON 문자열: 채널 노출 순서
    raw_count = Column(Integer, default=0)
    agg_count = Column(Integer, default=0)
    uploaded_by = Column(Integer, ForeignKey("users.id"), nullable=True)
    uploaded_at = Column(DateTime, default=datetime.datetime.utcnow)
    is_active = Column(Boolean, default=True)


class ItemRaw(Base):
    """통합매출 원본 트랜잭션 (~150만 행). 언제든 SQL 조회용."""
    __tablename__ = "item_raw"

    id = Column(Integer, primary_key=True)
    snapshot_id = Column(
        Integer, ForeignKey("item_snapshots.id", ondelete="CASCADE"), nullable=False, index=True
    )
    yr = Column(Integer)
    date_int = Column(Integer)          # YYYYMMDD
    cust_code = Column(String(30))
    cust_name = Column(String(200))
    item_code = Column(String(30))
    item_name = Column(String(200))
    qty = Column(Numeric(18, 3), default=0)
    rev = Column(Numeric(18, 3), default=0)
    cost = Column(Numeric(18, 3), default=0)

    __table_args__ = (
        Index("ix_iraw_snap_item", "snapshot_id", "item_code"),
        Index("ix_iraw_snap_cust", "snapshot_id", "cust_code"),
    )


class ItemRecord(Base):
    """15차원 집계 결과 (~68,890건). 품목 대시보드 렌더용."""
    __tablename__ = "item_records"

    id = Column(Integer, primary_key=True)
    snapshot_id = Column(
        Integer, ForeignKey("item_snapshots.id", ondelete="CASCADE"), nullable=False, index=True
    )
    yr = Column(Integer)
    month = Column(Integer)
    quarter = Column(String(4))
    team = Column(String(50))
    channel = Column(String(100))
    country = Column(String(60))
    customer = Column(String(200))
    brand = Column(String(20))
    theme = Column(String(100))
    dl_cat = Column(String(100))
    item_group = Column(String(100))
    item_cat = Column(String(100))
    sku = Column(String(30))
    sku_name = Column(String(200))
    sale_type = Column(String(20))
    qty = Column(Numeric(18, 3), default=0)
    rev = Column(Numeric(18, 3), default=0)
    cost = Column(Numeric(18, 3), default=0)

    __table_args__ = (
        Index("ix_irec_snap_team", "snapshot_id", "team"),
    )
