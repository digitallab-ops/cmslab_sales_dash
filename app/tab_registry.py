"""
탭 레지스트리 — 앱에 존재하는 모든 탭을 여기서 관리.
새 탭 추가 시 이 목록에만 추가하면 어드민 권한 설정 UI에 자동 반영됨.
"""

# scope_default: 명시적 권한(tab_perms)이 없을 때 그 탭에서 보이는 팀 범위 기본값
#   "all" = 전체 팀,  "own" = 사용자 본인 소속 팀(group_team)만
# 매출 대시보드만 전체 팀, 그 외 모든 탭(신규 포함)은 기본 본인 팀.
TABS = [
    {"id": "dashboard", "label": "매출 대시보드", "route": "/dashboard", "scope_default": "all"},
    {"id": "compare",   "label": "매출현황(표)",  "route": "/compare",   "scope_default": "own"},
    {"id": "items",     "label": "품목별 매출",   "route": "/items",     "scope_default": "own"},
]

# 미지정 탭(신규 등)의 기본은 "own" — 본인 팀만 보이도록 안전하게.
_SCOPE_DEFAULT = {t["id"]: t.get("scope_default", "own") for t in TABS}


def resolve_perms(user, group_team=None):
    """유효 tab_perms 반환. 개인 설정 → 그룹 기본값 → None(전체 탭·전체 팀).
    tab_perms = {tab_id: "ALL"|[teams]}. 키 없으면 그 탭 접근 불가.
    """
    if getattr(user, "tab_perms", None) is not None:
        return user.tab_perms
    if group_team is not None and getattr(group_team, "tab_perms", None) is not None:
        return group_team.tab_perms
    return None


def can_access_tab(user, tab_id: str, group_team=None) -> bool:
    """탭 접근 권한 확인. tab_perms가 없으면(NULL 상속) 전체 허용."""
    perms = resolve_perms(user, group_team)
    if perms is None:
        return True
    return tab_id in perms


def tab_teams(user, tab_id: str, group_team=None):
    """해당 탭에서 볼 수 있는 팀 목록. None = 전체 팀(제한 없음).

    명시적 tab_perms가 없으면 탭 기본값(scope_default) 적용:
      - "all"  → 전체 팀
      - "own"  → 사용자 본인 소속 팀(group_team)만. 소속 없으면 전체(관리자가 그룹 지정 필요).
    """
    perms = resolve_perms(user, group_team)
    if perms is None:
        if _SCOPE_DEFAULT.get(tab_id, "own") == "own":
            name = getattr(group_team, "name", None) if group_team is not None else None
            return [name] if name else None
        return None
    scope = perms.get(tab_id)
    if scope is None or scope == "ALL":
        return None
    return scope if scope else None
