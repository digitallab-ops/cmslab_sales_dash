import gzip
from typing import Optional
from fastapi import APIRouter, Depends, Request
from fastapi.responses import HTMLResponse, RedirectResponse, Response
from sqlalchemy.orm import Session

from ..database import get_db
from ..auth import get_current_user
from ..models import User, Team
from ..tab_registry import can_access_tab, TABS
from ..data.parser import (
    get_active_records,
    get_active_snapshot_info,
    get_all_snapshots,
    get_cached_html,
    get_record_count,
    _AGGREGATE_THRESHOLD,
)
from ..data.item_parser import (
    get_active_item_records,
    get_active_item_snapshot_info,
    get_cached_item_html,
    aggregate_range_records,
    get_item_date_bounds,
)

router = APIRouter()

_NO_DATA_HTML = """
<html><body style="font-family:sans-serif;padding:40px">
<h2>데이터가 없습니다</h2>
<p>관리자 패널에서 Excel 파일을 먼저 업로드해주세요.</p>
<p><a href="/admin">관리자 패널로 이동</a></p>
</body></html>
"""

_NO_ACCESS_HTML = """
<html><body style="font-family:sans-serif;padding:40px">
<h2>접근 권한이 없습니다</h2>
<p>이 페이지에 대한 접근 권한이 없습니다. 관리자에게 문의하세요.</p>
<p><a href="/dashboard">대시보드로 이동</a></p>
</body></html>
"""

_NAV_HEIGHT = 36   # 네비바 높이(px) — 대시보드 sticky 오프셋 계산에 사용

_CHAT_WIDGET = """
<style>
#__chat-btn{position:fixed;bottom:24px;right:24px;z-index:99997;background:#1a56a0;color:#fff;
  border-radius:50%;width:52px;height:52px;display:flex;align-items:center;justify-content:center;
  cursor:pointer;box-shadow:0 4px 14px rgba(0,0,0,.28);font-size:22px;user-select:none;
  transition:transform .15s}
#__chat-btn:hover{transform:scale(1.08)}
#__chat-panel{position:fixed;bottom:88px;right:24px;z-index:99997;width:360px;height:520px;
  background:#fff;border-radius:14px;box-shadow:0 8px 32px rgba(0,0,0,.22);
  display:none;flex-direction:column;font-family:'Malgun Gothic',sans-serif;overflow:hidden}
#__chat-msgs{flex:1;overflow-y:auto;padding:12px;display:flex;flex-direction:column;gap:8px}
#__chat-msgs .msg-ai{background:#f3f4f6;color:#374151;border-radius:10px 10px 10px 2px;
  padding:10px 13px;font-size:13px;max-width:88%;align-self:flex-start;white-space:pre-wrap;line-height:1.55}
#__chat-msgs .msg-user{background:#1a56a0;color:#fff;border-radius:10px 10px 2px 10px;
  padding:10px 13px;font-size:13px;max-width:88%;align-self:flex-end}
#__chat-msgs .msg-thinking{color:#9ca3af;font-size:12px;align-self:flex-start;padding:4px 0}
#__chat-input-row{padding:8px 10px;border-top:1px solid #e5e7eb;display:flex;gap:6px}
#__chat-input{flex:1;padding:8px 12px;border:1px solid #d1d5db;border-radius:8px;
  font-size:13px;outline:none;font-family:inherit}
#__chat-input:focus{border-color:#1a56a0}
#__chat-send{background:#1a56a0;color:#fff;border:none;border-radius:8px;padding:8px 14px;
  cursor:pointer;font-size:13px;white-space:nowrap}
#__chat-send:disabled{opacity:.5;cursor:not-allowed}
</style>

<div id="__chat-btn" onclick="__toggleChat()" title="AI 매출 분석">💬</div>
<div id="__chat-panel">
  <div style="background:#1a56a0;color:#fff;padding:11px 16px;display:flex;justify-content:space-between;align-items:center">
    <span style="font-weight:700;font-size:14px">AI 매출 분석</span>
    <div style="display:flex;gap:10px;align-items:center">
      <button onclick="__clearChat()" style="background:rgba(255,255,255,.15);border:none;color:#fff;
        border-radius:5px;padding:2px 8px;font-size:11px;cursor:pointer">초기화</button>
      <span onclick="__toggleChat()" style="cursor:pointer;font-size:20px;line-height:1">×</span>
    </div>
  </div>
  <div id="__chat-msgs">
    <div class="msg-ai">안녕하세요! 매출 데이터에 대해 질문해주세요.<br>
예) "이번 달 전사 실적은?", "RBD1팀 성과 어때?", "계획 대비 달성률 분석해줘"</div>
  </div>
  <div id="__chat-input-row">
    <input id="__chat-input" type="text" placeholder="질문 입력 후 Enter"
      onkeydown="if(event.key==='Enter'&&!event.isComposing)__sendChat()">
    <button id="__chat-send" onclick="__sendChat()">전송</button>
  </div>
</div>

<script>
(function(){
  var _history = [];
  window.__toggleChat = function(){
    var p=document.getElementById('__chat-panel');
    p.style.display=p.style.display==='flex'?'none':'flex';
    if(p.style.display==='flex') document.getElementById('__chat-input').focus();
  };
  window.__clearChat = function(){
    _history=[];
    var msgs=document.getElementById('__chat-msgs');
    msgs.innerHTML='<div class="msg-ai">대화가 초기화됐습니다. 새로운 질문을 입력하세요.</div>';
  };
  window.__sendChat = async function(){
    var input=document.getElementById('__chat-input');
    var msg=input.value.trim(); if(!msg) return;
    input.value='';
    var msgs=document.getElementById('__chat-msgs');
    var send=document.getElementById('__chat-send');

    var uDiv=document.createElement('div'); uDiv.className='msg-user'; uDiv.textContent=msg;
    msgs.appendChild(uDiv);

    var thinkDiv=document.createElement('div'); thinkDiv.className='msg-thinking';
    thinkDiv.textContent='분석 중...'; msgs.appendChild(thinkDiv);
    msgs.scrollTop=msgs.scrollHeight;

    send.disabled=true; input.disabled=true;
    try{
      var res=await fetch('/chat',{method:'POST',headers:{'Content-Type':'application/json'},
        body:JSON.stringify({message:msg,history:_history})});
      thinkDiv.remove();
      var data=await res.json();
      var reply=data.reply||data.detail||'오류가 발생했습니다.';
      _history.push({role:'user',content:msg},{role:'assistant',content:reply});
      if(_history.length>20) _history=_history.slice(-20);
      var aDiv=document.createElement('div'); aDiv.className='msg-ai'; aDiv.textContent=reply;
      msgs.appendChild(aDiv);
    }catch(e){
      thinkDiv.remove();
      var eDiv=document.createElement('div'); eDiv.className='msg-ai';
      eDiv.textContent='네트워크 오류: '+e.message; msgs.appendChild(eDiv);
    }
    msgs.scrollTop=msgs.scrollHeight;
    send.disabled=false; input.disabled=false; input.focus();
  };
})();
</script>
"""

_NAV_STYLE = f"""
<style>
#__cms-topnav{{position:fixed;top:0;left:0;right:0;z-index:99999;
  background:#1a56a0;padding:6px 20px;display:flex;align-items:center;
  justify-content:space-between;font-family:'Malgun Gothic',sans-serif;
  font-size:13px;box-shadow:0 2px 6px rgba(0,0,0,.2)}}
#__cms-topnav .nav-title{{color:#fff;font-weight:700;letter-spacing:-.3px}}
#__cms-topnav .nav-links{{display:flex;gap:4px;align-items:center;margin-left:18px}}
#__cms-topnav .nav-links a.nav-page{{color:rgba(255,255,255,.75);text-decoration:none;
  padding:4px 12px;border-radius:6px;font-size:13px;font-weight:500}}
#__cms-topnav .nav-links a.nav-page:hover{{background:rgba(255,255,255,.12);color:#fff}}
#__cms-topnav .nav-links a.nav-page.active{{background:rgba(255,255,255,.2);color:#fff;font-weight:700}}
#__cms-topnav .nav-left{{display:flex;align-items:center}}
#__cms-topnav .nav-right{{display:flex;gap:14px;align-items:center}}
#__cms-topnav .nav-datepick{{display:flex;gap:6px;align-items:center}}
#__cms-topnav .nav-datepick input[type=date]{{background:rgba(255,255,255,.15);color:#fff;
  border:1px solid rgba(255,255,255,.3);border-radius:4px;padding:2px 6px;font-size:12px;
  outline:none;color-scheme:dark}}
#__cms-topnav .nav-datepick button{{background:rgba(255,255,255,.9);color:#1a56a0;border:none;
  border-radius:4px;padding:3px 10px;font-size:12px;font-weight:700;cursor:pointer}}
#__cms-topnav .nav-datepick button:hover{{background:#fff}}
#__cms-topnav .nav-datepick .dp-reset{{color:rgba(255,255,255,.85);font-size:12px;
  text-decoration:none;padding:2px 4px}}
#__cms-topnav .nav-datepick .dp-reset:hover{{color:#fff;text-decoration:underline}}
#__cms-topnav .nav-user{{color:rgba(255,255,255,.75)}}
#__cms-topnav a{{color:rgba(255,255,255,.85);text-decoration:none}}
#__cms-topnav a:hover{{color:#fff}}
#__cms-topnav .nav-logout{{background:rgba(255,255,255,.15);padding:3px 12px;
  border-radius:5px;color:#fff !important}}
#__cms-topnav .nav-logout:hover{{background:rgba(255,255,255,.28)}}
body{{padding-top:{_NAV_HEIGHT}px !important}}
/* 대시보드 sticky/fixed 요소: 네비바 높이만큼 top 오프셋 */
.sticky-header   {{top:{_NAV_HEIGHT}px !important}}
.sticky-filterbar{{top:{44  + _NAV_HEIGHT}px !important}}
.sticky-kpi      {{top:{114 + _NAV_HEIGHT}px !important}}
#pane-print      {{top:{114 + _NAV_HEIGHT}px !important}}
</style>
"""

# #pane-table top은 JS가 (hdrH + filterH + 8)로 계산하는데 nav 높이를 빠뜨림.
# MutationObserver로 style 속성이 바뀔 때마다 nav 높이를 더해서 교정.
_NAV_SCRIPT = f"""<script>
(function(){{
  var H = {_NAV_HEIGHT}, _lock = false;
  var p = document.getElementById('pane-table');
  if (!p) return;
  new MutationObserver(function() {{
    if (_lock) return;
    var hdrH    = document.querySelector('.sticky-header')?.offsetHeight    ?? 44;
    var filterH = document.querySelector('.sticky-filterbar')?.offsetHeight ?? 70;
    var want = H + hdrH + filterH + 8;
    if (parseInt(p.style.top, 10) !== want) {{
      _lock = true; p.style.top = want + 'px'; _lock = false;
    }}
  }}).observe(p, {{attributes:true, attributeFilter:['style']}});
}})();
</script>"""


_NOCACHE = {"Cache-Control": "no-cache, no-store, must-revalidate"}


def _html_response(html: str, request: Request, headers: dict = None) -> Response:
    """대용량 HTML은 gzip으로 압축해 응답(브라우저가 gzip 허용 시). 25MB → ~3MB 전송."""
    h = dict(headers or {})
    accept = request.headers.get("accept-encoding", "").lower()
    if "gzip" in accept and len(html) > 4096:
        body = gzip.compress(html.encode("utf-8"), compresslevel=5)
        h["Content-Encoding"] = "gzip"
        h["Vary"] = "Accept-Encoding"
        return Response(content=body, media_type="text/html; charset=utf-8", headers=h)
    return HTMLResponse(html, headers=h)


def _inject_nav(html: str, user: User, db=None, base_date: str = "", snapshots=None, current_snap_id: int = 0, hidden_routes: list = None, current_route: str = "", date_bounds=None, date_from=None, date_to=None) -> str:
    from ..models import AppConfig

    app_title = "CMS Lab 매출 대시보드"
    notice_html = ""
    try:
        if db:
            title_row = db.query(AppConfig).filter(AppConfig.key == "app_title").first()
            if title_row and title_row.value:
                app_title = title_row.value
            notice_on = db.query(AppConfig).filter(AppConfig.key == "notice_enabled").first()
            notice_text = db.query(AppConfig).filter(AppConfig.key == "notice_text").first()
            if notice_on and notice_on.value == "true" and notice_text and notice_text.value:
                notice_html = f"""<div style="background:#fef3c7;color:#92400e;padding:6px 20px;
                    font-family:sans-serif;font-size:13px;text-align:center;border-bottom:1px solid #fde68a">
                    📢 {notice_text.value}</div>"""
    except Exception:
        pass

    date_badge = (
        f'<span style="color:rgba(255,255,255,.6);font-size:12px">{base_date} 기준</span>'
        if base_date else ""
    )

    # 주차 선택 드롭다운 (스냅샷 2개 이상일 때만 표시)
    week_selector = ""
    if snapshots and len(snapshots) > 1:
        options = "".join(
            f'<option value="{s["id"]}"{"  selected" if s["id"] == current_snap_id else ""}>{s["week_label"]}</option>'
            for s in snapshots
        )
        week_selector = f"""<select id="__week-sel" onchange="(function(v){{var p=location.pathname;location.href=p+'?snap='+v;}})(this.value)"
  style="background:rgba(255,255,255,.15);color:#fff;border:1px solid rgba(255,255,255,.3);
         border-radius:4px;padding:2px 6px;font-size:12px;cursor:pointer;outline:none">
  {options}
</select>"""

    admin_link = '<a href="/admin">관리자</a>' if user.role == "admin" else ""

    # 페이지 간 이동 링크 (권한 없는 페이지는 hidden_routes에 있으므로 제외)
    _hidden = set(hidden_routes or [])
    _nav_pages = [("/dashboard", "매출 대시보드"), ("/items", "품목별 매출")]
    _link_parts = []
    for route, label in _nav_pages:
        if route in _hidden:
            continue
        cls = "nav-page active" if route == current_route else "nav-page"
        _link_parts.append(f'<a href="{route}" class="{cls}">{label}</a>')
    nav_links = f'<div class="nav-links">{"".join(_link_parts)}</div>' if _link_parts else ""

    # 품목별 매출: 날짜 범위 선택기
    date_picker = ""
    if current_route == "/items":
        def _iso(n):
            return f"{n//10000}-{(n//100)%100:02d}-{n%100:02d}" if n else ""
        _min = _iso(date_bounds[0]) if date_bounds else ""
        _max = _iso(date_bounds[1]) if date_bounds else ""
        _vf = _iso(date_from) if date_from else _min
        _vt = _iso(date_to) if date_to else _max
        _reset = '<a href="/items" class="dp-reset" title="전체 기간">전체</a>' if (date_from or date_to) else ""
        date_picker = f"""<div class="nav-datepick">
      <input type="date" id="__dp-from" value="{_vf}" min="{_min}" max="{_max}">
      <span style="color:rgba(255,255,255,.6)">~</span>
      <input type="date" id="__dp-to" value="{_vt}" min="{_min}" max="{_max}">
      <button onclick="(function(){{var f=document.getElementById('__dp-from').value,t=document.getElementById('__dp-to').value;if(!f||!t){{alert('시작일과 종료일을 선택하세요');return;}}location.href='/items?from='+f+'&to='+t;}})()">적용</button>
      {_reset}
    </div>"""

    nav = f"""{_NAV_STYLE}
<div id="__cms-topnav">
  <div class="nav-left">
    <span class="nav-title">{app_title}</span>
    {nav_links}
  </div>
  <div class="nav-right">
    {date_picker}
    {week_selector}
    {date_badge}
    <span class="nav-user">{user.name or user.email}</span>
    {admin_link}
    <a href="/logout" class="nav-logout">로그아웃</a>
  </div>
</div>{notice_html}"""
    if "<body" in html:
        html = html.replace("<body", nav + "<body", 1)
    else:
        html = nav + html
    # </body> 직전에 pane-table top 교정 스크립트 + 챗 위젯 삽입
    # 접근 불가 탭 링크 숨기기
    hide_script = ""
    if hidden_routes:
        routes_js = ", ".join(f'"{r}"' for r in hidden_routes)
        hide_script = f"""<script>
(function(){{
  var hidden = [{routes_js}];
  function hideTabLinks() {{
    hidden.forEach(function(route) {{
      document.querySelectorAll('a[href="' + route + '"]').forEach(function(a) {{
        var el = a.closest('li,td,th,div.tab,span') || a;
        el.style.display = 'none';
      }});
    }});
  }}
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', hideTabLinks);
  else hideTabLinks();
}})();
</script>"""

    if "</body>" in html:
        html = html.replace("</body>", _NAV_SCRIPT + hide_script + "</body>", 1)
    else:
        html = html + _NAV_SCRIPT + hide_script
    return html


@router.get("/dashboard", response_class=HTMLResponse)
async def dashboard(
    snap: Optional[int] = None,
    db: Session = Depends(get_db),
    current_user: Optional[User] = Depends(get_current_user),
):
    if not current_user:
        return RedirectResponse("/login", status_code=302)
    group_team = db.query(Team).filter(Team.id == current_user.group_team_id).first() if current_user.group_team_id else None
    if not can_access_tab(current_user, "dashboard", group_team):
        return HTMLResponse(_NO_ACCESS_HTML, status_code=403)

    info = get_active_snapshot_info(db, snapshot_id=snap)
    if not info:
        return HTMLResponse(_NO_DATA_HTML)

    use_agg = get_record_count(db, snapshot_id=info["id"]) > _AGGREGATE_THRESHOLD
    records = get_active_records(db, current_user.allowed_teams, aggregated=use_agg, snapshot_id=info["id"])
    if not records:
        return HTMLResponse(_NO_DATA_HTML)

    all_snaps = get_all_snapshots(db)
    hidden = [t["route"] for t in TABS if not can_access_tab(current_user, t["id"], group_team)]
    html = get_cached_html("dashboard", info["id"], current_user.allowed_teams, records, info["base_date"])
    return HTMLResponse(
        _inject_nav(html, current_user, db, base_date=info["base_date"], snapshots=all_snaps, current_snap_id=info["id"], hidden_routes=hidden, current_route="/dashboard"),
        headers={"Cache-Control": "no-cache, no-store, must-revalidate"},
    )


@router.get("/compare", response_class=HTMLResponse)
async def compare(
    snap: Optional[int] = None,
    db: Session = Depends(get_db),
    current_user: Optional[User] = Depends(get_current_user),
):
    if not current_user:
        return RedirectResponse("/login", status_code=302)
    group_team = db.query(Team).filter(Team.id == current_user.group_team_id).first() if current_user.group_team_id else None
    if not can_access_tab(current_user, "compare", group_team):
        return HTMLResponse(_NO_ACCESS_HTML, status_code=403)

    info = get_active_snapshot_info(db, snapshot_id=snap)
    if not info:
        return HTMLResponse(_NO_DATA_HTML)

    use_agg = get_record_count(db, snapshot_id=info["id"]) > _AGGREGATE_THRESHOLD
    records = get_active_records(db, current_user.allowed_teams, aggregated=use_agg, snapshot_id=info["id"])
    if not records:
        return HTMLResponse(_NO_DATA_HTML)

    all_snaps = get_all_snapshots(db)
    hidden = [t["route"] for t in TABS if not can_access_tab(current_user, t["id"], group_team)]
    html = get_cached_html("compare", info["id"], current_user.allowed_teams, records, info["base_date"])
    return HTMLResponse(
        _inject_nav(html, current_user, db, base_date=info["base_date"], snapshots=all_snaps, current_snap_id=info["id"], hidden_routes=hidden, current_route="/compare"),
        headers={"Cache-Control": "no-cache, no-store, must-revalidate"},
    )


def _parse_date_param(s: str) -> Optional[int]:
    """'2026-06-15' 또는 '20260615' → 20260615(int). 실패 시 None."""
    if not s:
        return None
    digits = s.replace("-", "").strip()
    if len(digits) == 8 and digits.isdigit():
        return int(digits)
    return None


@router.get("/items", response_class=HTMLResponse)
async def items(
    request: Request,
    db: Session = Depends(get_db),
    current_user: Optional[User] = Depends(get_current_user),
):
    if not current_user:
        return RedirectResponse("/login", status_code=302)
    group_team = db.query(Team).filter(Team.id == current_user.group_team_id).first() if current_user.group_team_id else None
    if not can_access_tab(current_user, "items", group_team):
        return HTMLResponse(_NO_ACCESS_HTML, status_code=403)

    info = get_active_item_snapshot_info(db)
    if not info:
        return HTMLResponse(_NO_DATA_HTML)

    # 날짜 범위 파라미터 (?from=YYYY-MM-DD&to=YYYY-MM-DD)
    d_from = _parse_date_param(request.query_params.get("from", ""))
    d_to = _parse_date_param(request.query_params.get("to", ""))
    bounds = get_item_date_bounds(db, info["id"])   # (min, max) or None

    hidden = [t["route"] for t in TABS if not can_access_tab(current_user, t["id"], group_team)]

    if d_from and d_to and d_from <= d_to:
        # 지정 구간만 즉석 재집계
        records = aggregate_range_records(db, info["id"], d_from, d_to, current_user.allowed_teams)
        if not records:
            return HTMLResponse(_NO_DATA_HTML)
        def _fmt(n): return f"{n//10000}-{(n//100)%100:02d}-{n%100:02d}"
        label = f"{_fmt(d_from)} ~ {_fmt(d_to)}"
        html = get_cached_item_html(
            info["id"], current_user.allowed_teams, records, label, info["channel_order"],
            cache_tag=f"items:{d_from}:{d_to}",
        )
    else:
        records = get_active_item_records(db, current_user.allowed_teams)
        if not records:
            return HTMLResponse(_NO_DATA_HTML)
        html = get_cached_item_html(
            info["id"], current_user.allowed_teams, records, info["base_date"], info["channel_order"]
        )

    final = _inject_nav(html, current_user, db, base_date=info["base_date"], hidden_routes=hidden,
                        current_route="/items", date_bounds=bounds, date_from=d_from, date_to=d_to)
    return _html_response(final, request, _NOCACHE)
