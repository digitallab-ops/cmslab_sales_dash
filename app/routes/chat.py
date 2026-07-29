"""
AI 챗봇 엔드포인트 — OpenAI API 연동
대시보드 현재 실적 데이터를 컨텍스트로 제공
"""
import os, re
from typing import Optional, List
from collections import defaultdict

from fastapi import APIRouter, Depends, HTTPException
from pydantic import BaseModel
from sqlalchemy.orm import Session
from sqlalchemy import func

from ..database import get_db
from ..auth import get_current_user
from ..models import User, Snapshot, SalesRecord

router = APIRouter()

ALLOWED_MODELS = {
    "gpt-4o-mini": "GPT-4o mini (빠름)",
    "gpt-4o":      "GPT-4o (정확)",
}


class ChatRequest(BaseModel):
    message: str
    history: Optional[List[dict]] = None   # [{role, content}, ...]
    model: str = "gpt-4o-mini"


# ── 컨텍스트 빌더 ─────────────────────────────────────────────────────────────

def _build_sales_context(db: Session, allowed_teams: Optional[list]) -> str:
    snap = db.query(Snapshot).filter(Snapshot.is_active == True).first()
    if not snap:
        return "현재 업로드된 데이터가 없습니다."

    m = re.search(r'(\d{1,2})월', snap.base_date)
    base_month = int(m.group(1)) if m else 12

    q = (
        db.query(
            SalesRecord.team,
            SalesRecord.month,
            func.sum(SalesRecord.actual).label("actual"),
            func.sum(SalesRecord.plan).label("plan"),
            func.sum(SalesRecord.y2025).label("y2025"),
        )
        .filter(SalesRecord.snapshot_id == snap.id, SalesRecord.month <= base_month)
    )
    if allowed_teams:
        q = q.filter(SalesRecord.team.in_(allowed_teams))

    rows = q.group_by(SalesRecord.team, SalesRecord.month).order_by(SalesRecord.team, SalesRecord.month).all()
    if not rows:
        return "접근 가능한 데이터가 없습니다."

    by_team = defaultdict(list)
    for r in rows:
        by_team[r.team].append(r)

    lines = [
        f"기준: {snap.week_label} ({snap.base_date})",
        f"누적 범위: 1~{base_month}월",
        "",
        "=== 팀별 Y2026 누적 실적 (백만원) ===",
    ]

    total_a = total_p = total_y = 0.0
    for team in sorted(by_team):
        team_rows = by_team[team]
        a = sum(float(r.actual or 0) for r in team_rows)
        p = sum(float(r.plan or 0) for r in team_rows)
        y = sum(float(r.y2025 or 0) for r in team_rows)
        ach = f"{a/p*100:.1f}%" if p else "N/A"
        monthly = ", ".join(f"{r.month}월:{float(r.actual or 0):,.0f}" for r in sorted(team_rows, key=lambda x: x.month))
        lines.append(f"[{team}] 실적:{a:,.0f} / 계획:{p:,.0f} (달성률:{ach}) / 전년동기:{y:,.0f}")
        lines.append(f"  월별: {monthly}")
        total_a += a; total_p += p; total_y += y

    ach_total = f"{total_a/total_p*100:.1f}%" if total_p else "N/A"
    lines.append(f"\n[전사합계] 실적:{total_a:,.0f} / 계획:{total_p:,.0f} (달성률:{ach_total}) / 전년동기:{total_y:,.0f}")
    return "\n".join(lines)


def _build_item_context(db: Session, allowed_teams: Optional[list]) -> str:
    """품목별 매출(item_records) 요약: 팀별·품목군·상위 SKU (백만원 단위 rev, 수량)."""
    from ..models import ItemSnapshot, ItemRecord

    snap = db.query(ItemSnapshot).filter(ItemSnapshot.is_active == True).first()
    if not snap:
        return "품목별 매출 데이터가 아직 업로드되지 않았습니다."

    base = db.query(ItemRecord).filter(ItemRecord.snapshot_id == snap.id)
    if allowed_teams:
        base = base.filter(ItemRecord.team.in_(allowed_teams))

    def _top(col, label, limit):
        rows = (
            base.with_entities(col, func.sum(ItemRecord.rev).label("rev"), func.sum(ItemRecord.qty).label("qty"))
            .group_by(col).order_by(func.sum(ItemRecord.rev).desc()).limit(limit).all()
        )
        out = []
        for r in rows:
            name = getattr(r, col.key) or "(미분류)"
            out.append(f"  {name}: 매출 {float(r.rev or 0)/1e6:,.0f}백만 / 수량 {float(r.qty or 0):,.0f}")
        return out

    scope = ("전체 팀" if not allowed_teams else ", ".join(allowed_teams))
    lines = [
        f"=== 품목별 매출 요약 (기준일 {snap.base_date}, 범위: {scope}) ===",
        "[팀별]",
    ]
    lines += _top(ItemRecord.team, "팀", 20)
    lines += ["[품목군 상위]"]
    lines += _top(ItemRecord.item_group, "품목군", 12)
    lines += ["[브랜드]"]
    lines += _top(ItemRecord.brand, "브랜드", 10)
    lines += ["[상위 SKU]"]
    lines += _top(ItemRecord.sku_name, "SKU", 15)
    return "\n".join(lines)


# ── 엔드포인트 ────────────────────────────────────────────────────────────────

@router.get("/chat/status")
async def chat_status():
    configured = bool(os.getenv("OPENAI_API_KEY", "").strip())
    return {"configured": configured}


@router.post("/chat")
async def chat(
    req: ChatRequest,
    db: Session = Depends(get_db),
    current_user: Optional[User] = Depends(get_current_user),
):
    if not current_user:
        raise HTTPException(401, "로그인이 필요합니다.")

    api_key = os.getenv("OPENAI_API_KEY", "").strip()
    if not api_key:
        raise HTTPException(503, "OPENAI_API_KEY가 설정되지 않았습니다. Render 환경변수를 확인하세요.")

    model = req.model if req.model in ALLOWED_MODELS else "gpt-4o-mini"

    from ..tab_registry import tab_teams
    from ..models import Team
    group_team = db.query(Team).filter(Team.id == current_user.group_team_id).first() if current_user.group_team_id else None
    sales_ctx = _build_sales_context(db, tab_teams(current_user, "dashboard", group_team))
    item_ctx = _build_item_context(db, tab_teams(current_user, "items", group_team))

    system_prompt = f"""당신은 CMS Lab 매출 대시보드 AI 분석 어시스턴트입니다.
현재 사용자: {current_user.name or current_user.email} ({'관리자' if current_user.role == 'admin' else '일반 사용자'})

{sales_ctx}

{item_ctx}

답변 규칙:
- 한국어로 간결하게 답변 (3~5문장 이내 원칙, 복잡한 분석은 예외)
- 수치는 백만원 단위 기준, 큰 수는 억원으로도 표기 (예: 50,510백만원 ≈ 약 505억원)
- 비율/등락 포함해서 답변
- 데이터에 없는 정보는 "데이터에 없음"으로 명확히"""

    messages = [{"role": "system", "content": system_prompt}]
    messages += list(req.history or [])[-20:]
    messages.append({"role": "user", "content": req.message})

    try:
        from openai import AsyncOpenAI
        client = AsyncOpenAI(api_key=api_key)
        response = await client.chat.completions.create(
            model=model,
            max_tokens=1024,
            messages=messages,
        )
        reply = response.choices[0].message.content
        return {"reply": reply, "model": model}
    except Exception as e:
        err = str(e)
        if "authentication" in err.lower() or "api key" in err.lower():
            raise HTTPException(503, "OPENAI_API_KEY가 유효하지 않습니다.")
        raise HTTPException(500, f"AI 오류: {err}")
