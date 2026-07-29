"""
탭 레지스트리 — 앱에 존재하는 모든 탭을 여기서 관리.
새 탭 추가 시 이 목록에만 추가하면 어드민 권한 설정 UI에 자동 반영됨.
"""

TABS = [
    {"id": "dashboard", "label": "매출 대시보드", "route": "/dashboard"},
    {"id": "compare",   "label": "매출현황(표)",  "route": "/compare"},
    {"id": "items",     "label": "품목별 매출",   "route": "/items"},
]


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
    """해당 탭에서 볼 수 있는 팀 목록. None = 전체 팀(제한 없음)."""
    perms = resolve_perms(user, group_team)
    if perms is None:
        return None
    scope = perms.get(tab_id)
    if scope is None or scope == "ALL":
        return None
    return scope if scope else None
