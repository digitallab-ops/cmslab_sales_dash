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


# ─── 전용 DB 커넥션 (풀 recycle/timeout 간섭 없이 장시간 대량 작업용) ──────────

def _pg_connect():
    """SQLAlchemy 풀과 무관한 전용 psycopg2 커넥션 (대량 적재/집계 전용)."""
    import psycopg2
    from ..database import DATABASE_URL
    return psycopg2.connect(DATABASE_URL)


# ─── 스트리밍 xlsx 리더 (openpyxl read_only 메모리 누수 회피, 메모리 일정) ─────
# openpyxl read_only는 행을 읽을수록 메모리를 계속 점유(150만 행 → 수백MB → OOM).
# 통합매출 대용량 파일은 zip+XML iterparse로 직접 스트리밍해 메모리를 평탄하게 유지한다.

_XL_NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"


def _xl_col_index(ref: str) -> int:
    s = ''.join(ch for ch in ref if ch.isalpha())
    n = 0
    for ch in s:
        n = n * 26 + (ord(ch.upper()) - 64)
    return n - 1


def _xl_shared_strings(z):
    from xml.etree import ElementTree as ET
    import io
    ss = []
    try:
        data = z.read('xl/sharedStrings.xml')
    except KeyError:
        return ss
    for _, el in ET.iterparse(io.BytesIO(data), events=('end',)):
        if el.tag == _XL_NS + 'si':
            ss.append(''.join(t.text or '' for t in el.iter(_XL_NS + 't')))
            el.clear()
    return ss


def _xl_active_sheet(z) -> str:
    import re
    wb = z.read('xl/workbook.xml').decode('utf-8', 'ignore')
    m = re.search(r'<sheet[^>]*r:id="([^"]+)"', wb)
    target = None
    if m:
        rid = m.group(1)
        rels = z.read('xl/_rels/workbook.xml.rels').decode('utf-8', 'ignore')
        m2 = (re.search(r'Id="' + re.escape(rid) + r'"[^>]*Target="([^"]+)"', rels)
              or re.search(r'Target="([^"]+)"[^>]*Id="' + re.escape(rid) + r'"', rels))
        if m2:
            target = m2.group(1)
    target = (target or 'worksheets/sheet1.xml').lstrip('/')
    if not target.startswith('xl/'):
        target = 'xl/' + target
    return target


def _iter_xlsx_rows(path: str, min_row: int = 1):
    """xlsx 시트를 (열 인덱스순) 값 리스트로 스트리밍. 메모리 일정."""
    import zipfile
    from xml.etree import ElementTree as ET
    z = zipfile.ZipFile(path)
    try:
        ss = _xl_shared_strings(z)
        sheet = _xl_active_sheet(z)
        with z.open(sheet) as f:
            context = ET.iterparse(f, events=('start', 'end'))
            _, root = next(context)
            for ev, el in context:
                if ev == 'end' and el.tag == _XL_NS + 'row':
                    if int(el.get('r', '0')) >= min_row:
                        vals, maxc = {}, -1
                        for c in el:
                            if c.tag != _XL_NS + 'c':
                                continue
                            ci = _xl_col_index(c.get('r', ''))
                            t = c.get('t')
                            v = c.find(_XL_NS + 'v')
                            if v is not None:
                                text = v.text
                            else:
                                ise = c.find(_XL_NS + 'is')
                                text = ''.join(tt.text or '' for tt in ise.iter(_XL_NS + 't')) if ise is not None else None
                            if t == 's' and text is not None:
                                try:
                                    val = ss[int(text)]
                                except (ValueError, IndexError):
                                    val = None
                            else:
                                val = text
                            vals[ci] = val
                            if ci > maxc:
                                maxc = ci
                        yield [vals.get(i) for i in range(maxc + 1)]
                    el.clear()
                    root.clear()
    finally:
        z.close()


# ─── 1) 원본 적재 (스트리밍 → COPY, 전용 커넥션) ──────────────────────────────

_COPY_COLS = ["snapshot_id", "yr", "date_int", "cust_code", "cust_name",
              "item_code", "item_name", "qty", "rev", "cost"]
_COPY_CHUNK = 50_000

_DIMS = ["yr", "month", "quarter", "team", "channel", "country", "customer",
         "brand", "theme", "dl_cat", "item_group", "item_cat", "sku", "sku_name", "sale_type"]
_REC_COLS = ["snapshot_id"] + _DIMS + ["qty", "rev", "cost"]


def ingest_raw(conn, snapshot_id: int, xlsx_path: str, yr: int,
               base_count: int = 0, task_id: Optional[str] = None) -> int:
    """통합매출 xlsx를 스트리밍하며 item_raw에 Postgres COPY로 적재 (전용 커넥션).
    청크마다 커밋 + item_snapshots.raw_count를 갱신해 진행상황을 DB에 남긴다. 적재 건수 반환.
    xlsx는 openpyxl(메모리 누수) 대신 스트리밍 리더로 읽어 메모리를 평탄하게 유지."""
    import io, csv
    from ..database import SCHEMA

    # NULL 마커를 특수 문자열로 지정 → 빈 문자열('')이 NULL로 저장되는 CSV 기본동작 회피
    copy_sql = f"COPY {SCHEMA}.item_raw ({','.join(_COPY_COLS)}) FROM STDIN WITH (FORMAT csv, NULL '\\N')"

    sio = io.StringIO()
    writer = csv.writer(sio, lineterminator="\n")
    buf_rows = 0
    inserted = 0

    def _flush():
        nonlocal buf_rows, inserted, sio, writer
        if buf_rows == 0:
            return
        sio.seek(0)
        cur = conn.cursor()
        cur.copy_expert(copy_sql, sio)
        inserted += buf_rows
        # 진행상황 DB 영속화 (UI 폴링이 끊겨도 어디까지 됐는지 확인 가능)
        cur.execute(
            f"UPDATE {SCHEMA}.item_snapshots SET raw_count=%s WHERE id=%s",
            (base_count + inserted, snapshot_id),
        )
        cur.close()
        conn.commit()
        sio = io.StringIO(); writer = csv.writer(sio, lineterminator="\n"); buf_rows = 0
        if task_id:
            _p._set_task(task_id, message=f"원본 적재 중 ({yr}년, {base_count + inserted:,}건)")

    def _cell(row, i):
        return row[i] if i < len(row) else None

    # pilot load_sales 와 동일한 컬럼 인덱스/필터 (min_row=3)
    for row in _iter_xlsx_rows(xlsx_path, min_row=3):
        date_val = _cell(row, 1); cust_raw = _cell(row, 2)
        cust_nm  = _cell(row, 3)
        sku_raw  = _cell(row, 4)
        item_nm  = _cell(row, 5)
        qty  = float(_cell(row, 6) or 0); rev = float(_cell(row, 7) or 0); cost = float(_cell(row, 8) or 0)
        if not (date_val and sku_raw and (rev or qty)):
            continue
        try:
            d = int(float(str(date_val))); month = (d % 10000) // 100
            if not 1 <= month <= 12:
                continue
        except (ValueError, TypeError):
            continue
        sku = str(sku_raw).strip()
        try:
            cust = str(int(float(str(cust_raw))))
        except (ValueError, TypeError):
            cust = str(cust_raw).strip()
        cust_name = (str(cust_nm).strip() if cust_nm else "")[:200]
        item_name = (str(item_nm).strip() if item_nm else "")[:200]
        writer.writerow([snapshot_id, yr, d, cust, cust_name, sku, item_name, qty, rev, cost])
        buf_rows += 1
        if buf_rows >= _COPY_CHUNK:
            _flush()

    _flush()
    return inserted


# ─── 2) DB 원본 → 15차원 집계 (전용 커넥션 + 서버사이드 커서) ──────────────────

class _Row:
    """psycopg2 튜플을 속성 접근으로 감싸는 경량 래퍼 (_accumulate 재사용용)."""
    __slots__ = ("yr", "date_int", "cust_code", "item_code", "item_name", "qty", "rev", "cost")
    def __init__(self, t):
        (self.yr, self.date_int, self.cust_code, self.item_code,
         self.item_name, self.qty, self.rev, self.cost) = t


_SELECT_RAW = "SELECT yr,date_int,cust_code,item_code,item_name,qty,rev,cost FROM {schema}.item_raw WHERE snapshot_id=%s"


def _accumulate(row_iter, mod, masters, task_id=None, progress_label="") -> Dict[tuple, list]:
    """item_raw 행 iterator(_Row)를 pilot과 동일한 per-row 변환으로 15차원 누적."""
    factory, cmap, country_map, customer_map = masters[0], masters[1], masters[2], masters[3]
    acc: Dict[tuple, list] = {}
    seen = 0
    for r in row_iter:
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
            _p._set_task(task_id, message=f"{progress_label} ({seen:,}건 처리, {len(acc):,} 그룹)")
    return acc


def _acc_to_records(acc: Dict[tuple, list], allowed_teams: Optional[List[str]] = None) -> List[Dict]:
    """누적 dict → make_html용 record dict 리스트 (정수 반올림, 팀 필터)."""
    team_set = set(allowed_teams) if allowed_teams else None
    out = []
    for key, (qv, rv, cv) in acc.items():
        if team_set is not None and key[3] not in team_set:   # key[3] == team
            continue
        rec = dict(zip(_DIMS, key))
        rec["qty"] = round(qv); rec["rev"] = round(rv); rec["cost"] = round(cv)
        out.append(rec)
    return out


def aggregate_from_db(conn, snapshot_id: int, mod, masters,
                      task_id: Optional[str] = None) -> int:
    """item_raw 전체를 서버사이드 커서로 스트리밍 집계 후 item_records에 COPY. 집계 건수 반환."""
    import io, csv
    from ..database import SCHEMA

    if task_id:
        _p._set_task(task_id, message="집계 중 (원본 스트리밍)...")

    rcur = conn.cursor(name="item_agg_stream")   # 서버사이드 커서 → 스트리밍
    rcur.itersize = 20000
    rcur.execute(_SELECT_RAW.format(schema=SCHEMA), (snapshot_id,))
    acc = _accumulate((_Row(t) for t in rcur), mod, masters, task_id, "집계 중")
    rcur.close()

    # item_records COPY 저장
    copy_sql = f"COPY {SCHEMA}.item_records ({','.join(_REC_COLS)}) FROM STDIN WITH (FORMAT csv, NULL '\\N')"
    sio = io.StringIO(); writer = csv.writer(sio, lineterminator="\n")
    buf = 0; agg = 0
    wcur = conn.cursor()

    def _flush():
        nonlocal sio, writer, buf, agg
        if buf == 0:
            return
        sio.seek(0)
        wcur.copy_expert(copy_sql, sio)
        conn.commit()
        agg += buf
        sio = io.StringIO(); writer = csv.writer(sio, lineterminator="\n"); buf = 0
        if task_id:
            _p._set_task(task_id, message=f"집계 저장 중 ({agg:,}건)")

    for key, (qv, rv, cv) in acc.items():
        writer.writerow([snapshot_id] + list(key) + [round(qv), round(rv), round(cv)])
        buf += 1
        if buf >= _COPY_CHUNK:
            _flush()
    _flush()
    wcur.close()
    return agg


def aggregate_range_records(db: Session, snapshot_id: int,
                            date_from: int, date_to: int,
                            allowed_teams: Optional[List[str]] = None) -> List[Dict]:
    """지정 일자 구간만 item_raw에서 즉석 재집계해 records 반환 (저장 안 함)."""
    from ..database import SCHEMA

    mod = _load_item_mod()
    masters = _load_masters_from_db(db, snapshot_id)
    conn = _pg_connect()
    try:
        cur = conn.cursor(name="item_range_stream")
        cur.itersize = 20000
        cur.execute(
            _SELECT_RAW.format(schema=SCHEMA) + " AND date_int BETWEEN %s AND %s",
            (snapshot_id, date_from, date_to),
        )
        acc = _accumulate((_Row(t) for t in cur), mod, masters)
        cur.close()
    finally:
        conn.close()
    return _acc_to_records(acc, allowed_teams)


# ─── 마스터 영속화 / 복원 ─────────────────────────────────────────────────────

def _persist_masters(db: Session, snapshot_id: int, masters):
    """업로드 시 로드한 마스터 dict를 DB에 저장 (날짜범위 재집계 시 재사용)."""
    from ..models import ItemMasterSku, ItemMasterCustomer
    factory, cmap, country_map, customer_map, _ = masters

    sku_rows = [{
        "snapshot_id": snapshot_id, "item_code": code,
        "sku_name": (v.get("sku_name") or "")[:200], "brand": (v.get("brand") or "")[:20],
        "dl_cat": (v.get("dl_cat") or "")[:100], "item_cat": (v.get("item_cat") or "")[:100],
        "item_group": (v.get("item_group") or "")[:100], "sale_type": (v.get("sale_type") or "")[:20],
        "theme": (v.get("theme") or "")[:100],
    } for code, v in factory.items()]
    ins_s = ItemMasterSku.__table__.insert()
    for i in range(0, len(sku_rows), _p._CHUNK_SIZE):
        db.execute(ins_s, sku_rows[i:i + _p._CHUNK_SIZE])
    db.commit()

    codes = set(cmap) | set(country_map) | set(customer_map)
    cust_rows = [{
        "snapshot_id": snapshot_id, "cust_code": c,
        "cust_name": (customer_map.get(c) or "")[:200],
        "channel": (cmap.get(c) or "")[:100],
        "country": (country_map.get(c) or "")[:60],
    } for c in codes]
    ins_c = ItemMasterCustomer.__table__.insert()
    for i in range(0, len(cust_rows), _p._CHUNK_SIZE):
        db.execute(ins_c, cust_rows[i:i + _p._CHUNK_SIZE])
    db.commit()


def _load_masters_from_db(db: Session, snapshot_id: int):
    """DB에 저장된 마스터를 dict로 복원 → (factory, cmap, country_map, customer_map, [])."""
    from ..models import ItemMasterSku, ItemMasterCustomer
    factory = {}
    for r in db.query(ItemMasterSku).filter(ItemMasterSku.snapshot_id == snapshot_id).yield_per(5000):
        factory[r.item_code] = {
            "sku_name": r.sku_name, "brand": r.brand, "dl_cat": r.dl_cat,
            "item_cat": r.item_cat, "item_group": r.item_group,
            "sale_type": r.sale_type, "theme": r.theme,
        }
    cmap, country_map, customer_map = {}, {}, {}
    for r in db.query(ItemMasterCustomer).filter(ItemMasterCustomer.snapshot_id == snapshot_id).yield_per(5000):
        if r.channel:
            cmap[r.cust_code] = r.channel
        if r.country:
            country_map[r.cust_code] = r.country
        if r.cust_name:
            customer_map[r.cust_code] = r.cust_name
    return factory, cmap, country_map, customer_map, []


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
        # 완료 전까지 is_active=False → 중간에 죽어도 반쪽 데이터가 화면에 노출되지 않음
        snap = ItemSnapshot(
            base_date=base_date,
            channel_order=json.dumps(channel_order, ensure_ascii=False),
            uploaded_by=uploaded_by,
            is_active=False,
            uploaded_at=datetime.datetime.utcnow(),
        )
        db.add(snap)
        db.commit()   # 커밋 → snap.id가 전용 커넥션에서도 FK로 보이게
        snap_id = snap.id

        # 마스터 영속화 (날짜범위 재집계 시 재사용)
        if task_id:
            _p._set_task(task_id, message="마스터 저장 중...")
        _persist_masters(db, snap_id, masters)

        # 무거운 적재/집계는 전용 psycopg2 커넥션으로 (풀 recycle/timeout 간섭 차단)
        conn = _pg_connect()
        try:
            if task_id:
                _p._set_task(task_id, status="inserting", message="원본 적재 시작...")
            raw_total = 0
            for yr, key in ((2025, "sales_2025"), (2026, "sales_2026")):
                if paths.get(key):
                    raw_total += ingest_raw(conn, snap_id, paths[key], yr, raw_total, task_id)

            if task_id:
                _p._set_task(task_id, status="aggregating", message="집계 시작...")
            agg_total = aggregate_from_db(conn, snap_id, mod, masters, task_id)
        finally:
            conn.close()

        # 완료 처리: 카운트 기록 + 활성화 + 이전 스냅샷 정리 (ORM 세션)
        snap = db.query(ItemSnapshot).filter(ItemSnapshot.id == snap_id).first()
        snap.raw_count = raw_total
        snap.agg_count = agg_total
        snap.is_active = True
        others = (
            db.query(ItemSnapshot)
            .filter(ItemSnapshot.id != snap_id)
            .order_by(ItemSnapshot.uploaded_at.desc())
            .all()
        )
        for i, o in enumerate(others):
            o.is_active = False
            if i >= 1:   # 현재 + 직전 1개만 남기고 삭제 (원본/마스터 CASCADE 삭제)
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


def get_item_date_bounds(db: Session, snapshot_id: int):
    """해당 스냅샷 item_raw의 (최소, 최대) date_int. 없으면 None."""
    from sqlalchemy import func
    from ..models import ItemRaw
    row = db.query(func.min(ItemRaw.date_int), func.max(ItemRaw.date_int)).filter(
        ItemRaw.snapshot_id == snapshot_id
    ).first()
    if not row or row[0] is None:
        return None
    return (int(row[0]), int(row[1]))


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
            "yr": r.yr, "month": r.month, "quarter": r.quarter or "",
            "team": r.team or "", "channel": r.channel or "", "country": r.country or "",
            "customer": r.customer or "", "brand": r.brand or "", "theme": r.theme or "",
            "dl_cat": r.dl_cat or "", "item_group": r.item_group or "", "item_cat": r.item_cat or "",
            "sku": r.sku or "", "sku_name": r.sku_name or "", "sale_type": r.sale_type or "",
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
                         records: List[Dict], base_date: str, channel_order: list,
                         cache_tag: str = "items") -> str:
    key = (snapshot_id, _p._teams_key(allowed_teams), cache_tag)
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
