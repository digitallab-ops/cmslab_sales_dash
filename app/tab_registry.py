"""
탭 레지스트리 — 앱에 존재하는 모든 탭을 여기서 관리.
새 탭 추가 시 이 목록에만 추가하면 어드민 권한 설정 UI에 자동 반영됨.
"""

TABS = [
    {"id": "dashboard", "label": "매출 대시보드", "route": "/dashboard"},
    {"id": "compare",   "label": "매출현황(표)",  "route": "/compare"},
    {"id": "items",     "label": "품목별 매출",   "route": "/items"},
]


def can_access_tab(user, tab_id: str, group_team=None) -> bool:
    """탭 접근 권한 확인.

    우선순위: 개인 설정 → 그룹 기본값 → 전체 허용
    - user.allowed_tabs 가 있으면 개인 설정 사용
    - 없으면 group_team.allowed_tabs 확인
    - 둘 다 없으면 전체 허용
    """
    if user.allowed_tabs is not None:
        return tab_id in user.allowed_tabs
    if group_team is not None and group_team.allowed_tabs is not None:
        return tab_id in group_team.allowed_tabs
    return True
