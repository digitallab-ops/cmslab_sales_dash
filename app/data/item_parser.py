"""
품목별 매출 대시보드 — 서버측 파이프라인.

흐름:
  1) ingest_raw()      : 통합매출 xlsx 스트리밍 → item_raw 청크 insert (메모리 안전)
  2) aggregate_from_db(): item_raw를 DB 커서로 스트리밍 → 15차원 dict 누적 → item_records
  3) get_cached_item_html(): item_records → pilot make_html 재활용 → 캐시

무거운 계산(150만 행)을 한 번에 메모리에 올리지 않고, 스트리밍/누적으로 처리한다.
pilot(품목_대시보드_pilot.py)의 로더·resolve_team·make_html·load_chartjs를 importlib로 재활용.
"""
import os, sys, json, importlib.util, datetime
from pathlib import Path
from typing import Dict, List, Optional

from sqlalchemy.orm import Session

from . import parser as _p   # 공유 캐시(_html_cache)·태스크(_set_task)·_CHUNK_SIZE 재사용

# pilot 스크립트 경로 (parser 의 scripts 폴더와 동일)
_ITEM_SRC_NAME = "품목_대시보드_pilot.py"
_item_mod = None
_chartjs_cache: Optional[str] = None


def _load_item_mod(force: bool = False):
    """품목_대시보드_pilot.py 를 importlib로 로드."""
    global _item_mod
    if _item_mod is not None and not force:
        return _item_mod
    path = os.path.join(os.path.abspath(_p.DASHBOARD_SRC_PATH), _ITEM_SRC_NAME)
    if not os.path.exists(path):
        raise FileNotFoundError(f"Item dashboard source not found: {path}")
    spec = importlib.util.spec_from_file_location("item_dashboard_main", path)
    mod = importlib.util.module_from_spec(spec)
    sys.modules["item_dashboard_main"] = mod
    spec.loader.exec_module(mod)
    _item_mod = mod
    return mod


# ─── 마스터 로드 (업로드된 파일을 pilot 로더로 재사용) ────────────────────────

# pilot find_latest 글롭 패턴에 맞는 저장 파일명
MASTER_NAMES = {
    "factory":  "공장품목조회.xls",
    "acct":     "1. 관리회계 기준정보.xlsx",
    "team_ref": "팀참고.xlsx",
}
SALES_NAMES = {2025: "통합매출_2025.xlsx", 2026: "통합매출_2026.xlsx"}


def _load_masters(mod, tmp_dir: str):
    """업로드 임시폴더를 pilot INPUT_DIR로 지정하고 마스터 3종을 dict로 로드."""
    mod.INPUT_DIR = Path(tmp_dir)
    factory = mod.load_factory()
    cmap, country_map, customer_map = mod.load_channel_map()
    channel_order = mod.load_channel_order()
    return factory, cmap, country_map, customer_map, channel_order


# ─── 1) 원본 적재 (스트리밍 → 청크 insert) ────────────────────────────────────

def ingest_raw(db: Session, snapshot_id: int, xlsx_path: str, yr: int,
               task_id: Optional[str] = None) -> int:
    """통합매출 xlsx를 스트리밍하며 item_raw에 청크 insert. 적재 건수 반환."""
    import openpyxl
    from ..models import ItemRaw

    wb = openpyxl.load_workbook(xlsx_path, read_only=True, data_only=True)
    ws = wb.active
    ins = ItemRaw.__table__.insert()

    buf: List[dict] = []
    inserted = 0

    def _flush():
        nonlocal inserted, buf
        if buf:
            db.execute(ins, buf)
            db.commit()
            inserted += len(buf)
            buf = []
            if task_id:
                _p._set_task(task_id, message=f"원본 적재 중 ({yr}년, {inserted:,}건)")

    # pilot load_sales 와 동일한 컬럼 인덱스/필터 (min_row=3)
    for row in ws.iter_rows(min_row=3, values_only=True):
        date_val = row[1]
        cust_raw = row[2]
        cust_nm  = row[3] if len(row) > 3 else None
        sku_raw  = row[4]
        item_nm  = row[5] if len(row) > 5 else None
        qty  = float(row[6] or 0)
        rev  = float(row[7] or 0)
        cost = float(row[8] or 0)
        if not (date_val and sku_raw and (rev or qty)):
            continue
        try:
            d = int(float(str(date_val)))
            month = (d % 10000) // 100
            if not 1 <= month <= 12:
                continue
        except (ValueError, TypeError):
            continue
        sku = str(sku_raw).strip()
        try:
            cust = str(int(float(str(cust_raw))))
        except (ValueError, TypeError):
            cust = str(cust_raw).strip()

        buf.append({
            "snapshot_id": snapshot_id,
            "yr": yr,
            "date_int": d,
            "cust_code": cust,
            "cust_name": (str(cust_nm).strip() if cust_nm else "")[:200],
            "item_code": sku,
            "item_name": (str(item_nm).strip() if item_nm else "")[:200],
            "qty": qty,
            "rev": rev,
            "cost": cost,
        })
        if len(buf) >= _p._CHUNK_SIZE:
            _flush()

    _flush()
    wb.close()
    return inserted


# ─── 2) DB 원본 → 15차원 집계 → item_records ──────────────────────────────────

_DIMS = ["yr", "month", "quarter", "team", "channel", "country", "customer",
         "brand", "theme", "dl_cat", "item_group", "item_cat", "sku", "sku_name", "sale_type"]


def aggregate_from_db(db: Session, snapshot_id: int, mod, masters,
                      task_id: Optional[str] = None) -> int:
    """item_raw를 스트리밍으로 읽어 15차원 누적 후 item_records에 청크 insert. 집계 건수 반환."""
    from ..models import ItemRaw, ItemRecord

    factory, cmap, country_map, customer_map, _ = masters
    acc: Dict[tuple, list] = {}   # key(15차원) -> [qty, rev, cost]

    if task_id:
        _p._set_task(task_id, message="집계 중 (원본 스트리밍)...")

    # pass 1: 스트리밍 누적 (server-side cursor)
    q = (
        db.query(ItemRaw)
        .filter(ItemRaw.snapshot_id == snapshot_id)
        .yield_per(10000)
    )
    seen = 0
    for r in q:
        sku = r.item_code
        cust = r.cust_code
        month = (r.date_int % 10000) // 100
        quarter = f"Q{(month - 1) // 3 + 1}"
        item = factory.get(sku, {})
        channel = cmap.get(cust, "")
        country = country_map.get(cust, "")
        customer = customer_map.get(cust, "") or cust
        brand = item.get("brand", "")
        team = mod.resolve_team(channel, brand)
        key = (
            r.yr, month, quarter, team, channel, country, customer, brand,
            item.get("theme", ""), item.get("dl_cat", ""),
            item.get("item_group", ""), item.get("item_cat", ""),
            sku, item.get("sku_name", r.item_name or ""),
            item.get("sale_type", "증정품"),
        )
        cell = acc.get(key)
        qv = float(r.qty or 0); rv = float(r.rev or 0); cv = float(r.cost or 0)
        if cell is None:
            acc[key] = [qv, rv, cv]
        else:
            cell[0] += qv; cell[1] += rv; cell[2] += cv
        seen += 1
        if task_id and seen % 200000 == 0:
            _p._set_task(task_id, message=f"집계 중 ({seen:,}건 처리, {len(acc):,} 그룹)")

    # pass 2: 청크 insert
    ins = ItemRecord.__table__.insert()
    rows: List[dict] = []
    agg = 0
    for key, (qv, rv, cv) in acc.items():
        rec = dict(zip(_DIMS, key))
        rec["snapshot_id"] = snapshot_id
        rec["qty"] = round(qv)
        rec["rev"] = round(rv)
        rec["cost"] = round(cv)
        rows.append(rec)
        if len(rows) >= _p._CHUNK_SIZE:
            db.execute(ins, rows); db.commit()
            agg += len(rows); rows = []
            if task_id:
                _p._set_task(task_id, message=f"집계 저장 중 ({agg:,}건)")
    if rows:
        db.execute(ins, rows); db.commit()
        agg += len(rows)
    return agg


# ─── 3) 오케스트레이션 ────────────────────────────────────────────────────────

def save_item_upload(db: Session, paths: Dict[str, str], uploaded_by: Optional[int] = None,
                     task_id: Optional[str] = None) -> dict:
    """엑셀 5종 경로(dict)를 받아 원본 적재 + 집계 저장. paths keys: factory, acct, team_ref, sales_2025, sales_2026."""
    import tempfile, shutil
    from ..models import ItemSnapshot

    mod = _load_item_mod()

    # pilot 로더가 글롭으로 찾도록, 임시폴더에 규칙명으로 복사
    tmp_dir = tempfile.mkdtemp(prefix="item_up_")
    try:
        shutil.copy(paths["factory"],  os.path.join(tmp_dir, MASTER_NAMES["factory"]))
        shutil.copy(paths["acct"],     os.path.join(tmp_dir, MASTER_NAMES["acct"]))
        if paths.get("team_ref"):
            shutil.copy(paths["team_ref"], os.path.join(tmp_dir, MASTER_NAMES["team_ref"]))

        if task_id:
            _p._set_task(task_id, status="parsing", message="마스터 데이터 로드 중...")
        masters = _load_masters(mod, tmp_dir)
        channel_order = masters[4]

        base_date = datetime.date.today().strftime("%Y년 %m월 %d일")
        snap = ItemSnapshot(
            base_date=base_date,
            channel_order=json.dumps(channel_order, ensure_ascii=False),
            uploaded_by=uploaded_by,
            is_active=True,
            uploaded_at=datetime.datetime.utcnow(),
        )
        db.add(snap)
        db.flush()   # snap.id 확보
        snap_id = snap.id

        # 원본 적재
        if task_id:
            _p._set_task(task_id, status="inserting", message="원본 적재 시작...")
        raw_total = 0
        for yr, key in ((2025, "sales_2025"), (2026, "sales_2026")):
            if paths.get(key):
                raw_total += ingest_raw(db, snap_id, paths[key], yr, task_id)

        # 집계
        if task_id:
            _p._set_task(task_id, status="aggregating", message="집계 시작...")
        agg_total = aggregate_from_db(db, snap_id, mod, masters, task_id)

        snap.raw_count = raw_total
        snap.agg_count = agg_total

        # 이전 스냅샷 inactive + 오래된 것 정리 (active 포함 2개만 유지)
        others = (
            db.query(ItemSnapshot)
            .filter(ItemSnapshot.id != snap_id)
            .order_by(ItemSnapshot.uploaded_at.desc())
            .all()
        )
        for i, o in enumerate(others):
            o.is_active = False
            if i >= 1:   # 최신(현재) + 직전 1개만 남기고 삭제
                db.delete(o)
        db.commit()

        _p.clear_html_cache()
        prewarm_item_cache(db)
        return {"ok": True, "raw": raw_total, "agg": agg_total, "base_date": base_date}
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)


# ─── 조회 · 렌더 ──────────────────────────────────────────────────────────────

def get_active_item_snapshot_info(db: Session) -> Optional[dict]:
    from ..models import ItemSnapshot
    _KST = datetime.timedelta(hours=9)
    s = db.query(ItemSnapshot).filter(ItemSnapshot.is_active == True).first()
    if not s:
        return None
    return {
        "id": s.id,
        "base_date": s.base_date,
        "channel_order": json.loads(s.channel_order) if s.channel_order else [],
        "raw_count": s.raw_count,
        "agg_count": s.agg_count,
        "uploaded_at": (s.uploaded_at + _KST).strftime("%Y-%m-%d %H:%M") if s.uploaded_at else "",
    }


def get_active_item_records(db: Session, allowed_teams: Optional[List[str]] = None) -> List[Dict]:
    """active 스냅샷의 item_records를 pilot make_html이 기대하는 dict 리스트로 반환."""
    from ..models import ItemSnapshot, ItemRecord

    snap = db.query(ItemSnapshot).filter(ItemSnapshot.is_active == True).first()
    if not snap:
        return []
    q = db.query(ItemRecord).filter(ItemRecord.snapshot_id == snap.id)
    if allowed_teams:
        q = q.filter(ItemRecord.team.in_(allowed_teams))

    out = []
    for r in q.yield_per(10000):
        out.append({
            "yr": r.yr, "month": r.month, "quarter": r.quarter,
            "team": r.team, "channel": r.channel, "country": r.country,
            "customer": r.customer, "brand": r.brand, "theme": r.theme,
            "dl_cat": r.dl_cat, "item_group": r.item_group, "item_cat": r.item_cat,
            "sku": r.sku, "sku_name": r.sku_name, "sale_type": r.sale_type,
            "qty": int(r.qty or 0), "rev": int(r.rev or 0), "cost": int(r.cost or 0),
        })
    return out


def make_item_html(records: List[Dict], base_date: str, channel_order: list) -> str:
    global _chartjs_cache
    mod = _load_item_mod()
    if _chartjs_cache is None:
        _chartjs_cache = mod.load_chartjs()
    return mod.make_html(records, _chartjs_cache, base_date, channel_order=channel_order)


def get_cached_item_html(snapshot_id: int, allowed_teams: Optional[List[str]],
                         records: List[Dict], base_date: str, channel_order: list) -> str:
    key = (snapshot_id, _p._teams_key(allowed_teams), "items")
    if key in _p._html_cache:
        return _p._html_cache[key]
    html = make_item_html(records, base_date, channel_order)
    _p._html_cache[key] = html
    return html


def prewarm_item_cache(db: Session):
    """업로드 완료 후 전체(all-teams) HTML만 미리 생성.
    품목 HTML은 ~25MB로 크므로 팀별 조합은 캐시 폭증을 막기 위해 최초 접근 시 지연 생성한다.
    """
    info = get_active_item_snapshot_info(db)
    if not info:
        return
    records = get_active_item_records(db, allowed_teams=None)
    if records:
        get_cached_item_html(info["id"], None, records, info["base_date"], info["channel_order"])
