"""
AI 에이전트별 맞춤형 가이드 생성기
OpenAI Codex의 "Less is More" 원칙을 적용하여 각 AI 특성에 맞는 가이드 제공
"""

import json
from enum import Enum
from typing import Any, Dict, Optional


class AIAssistant(str, Enum):
    """지원되는 AI 어시스턴트 타입"""

    default = "default"
    codex = "codex"
    claude = "claude"
    gemini = "gemini"
    copilot = "copilot"

    @classmethod
    def get_description(cls, value: str) -> str:
        """각 AI 타입에 대한 설명"""
        descriptions = {
            "default": "범용 AI 어시스턴트 (기본값)",
            "codex": "OpenAI Codex CLI - Less is More 원칙",
            "claude": "Claude Code - 체계적 워크플로우",
            "gemini": "Gemini CLI - 대화형 상호작용",
            "copilot": "GitHub Copilot - IDE 통합형",
        }
        return descriptions.get(value, "")


class OutputFormat(str, Enum):
    """출력 형식 옵션"""

    json = "json"
    text = "text"
    markdown = "markdown"


class AIGuideGenerator:
    """AI 어시스턴트별 맞춤형 가이드 생성기"""

    def __init__(self):
        """가이드 생성 전략 매핑"""
        self.strategies = {
            AIAssistant.default: self._default_guide,
            AIAssistant.codex: self._codex_guide,
            AIAssistant.claude: self._claude_guide,
            AIAssistant.gemini: self._gemini_guide,
            AIAssistant.copilot: self._default_guide,  # copilot은 default 사용
        }

    def generate(self, ai_type: AIAssistant, verbose: bool = False, lang: str = "ko") -> Dict[str, Any]:
        """AI 타입에 맞는 가이드 생성"""
        strategy = self.strategies.get(ai_type, self._default_guide)
        guide = strategy(verbose, lang)

        # 메타데이터 추가
        guide["meta"] = {
            "ai_type": ai_type.value,
            "description": AIAssistant.get_description(ai_type.value),
            "verbose": verbose,
            "lang": lang,
            "version": "1.0",
        }

        return guide

    def _default_guide(self, verbose: bool, lang: str) -> Dict[str, Any]:
        """기본 범용 가이드 - 표준 워크플로우, 균형잡힌 정보"""
        guide = {
            "workflow": {
                "1_discover": {"command": "oa excel workbook-list", "purpose": "현재 열린 워크북 확인"},
                "2_analyze": {"command": "oa excel table-list", "purpose": "테이블 구조 및 데이터 파악"},
                "3_execute": {"command": "oa excel [operation]", "purpose": "실제 작업 수행"},
            },
            "connection_methods": {
                "active": {"usage": "옵션 없음", "description": "활성 워크북 자동 사용 (기본값)"},
                "file": {"usage": '--file-path "경로/파일.xlsx"', "description": "파일 경로로 직접 지정"},
                "name": {"usage": '--workbook-name "파일.xlsx"', "description": "열린 워크북 이름으로 연결"},
            },
            "output": {"format": "json", "parsing": "AI 에이전트가 파싱하기 최적화된 구조"},
        }

        if verbose:
            guide["examples"] = [
                {"task": "데이터 읽기", "command": "oa excel range-read --range A1:C10", "output": "JSON 형식의 셀 데이터"},
                {
                    "task": "차트 생성",
                    "command": "oa excel chart-add --data-range A1:B20 --chart-type column",
                    "output": "차트 생성 결과 및 위치 정보",
                },
                {"task": "테이블 분석", "command": "oa excel table-list", "output": "모든 테이블의 구조와 샘플 데이터"},
            ]
            guide["best_practices"] = [
                "작업 전 반드시 workbook-list로 상태 확인",
                "table-list로 데이터 구조 파악 후 작업",
                "JSON 출력을 파싱하여 사용자 친화적으로 변환",
                "에러 발생 시 구조화된 에러 메시지 처리",
                "--workbook-name으로 효율적인 연속 작업",
            ]
            guide["error_handling"] = {
                "workbook_not_found": "workbook-list로 확인 후 올바른 이름 사용",
                "sheet_missing": "workbook-info로 시트 목록 확인",
                "range_invalid": "데이터 범위 경계 검증",
                "permission_denied": "Excel 애플리케이션 실행 상태 확인",
            }

        return guide

    def _codex_guide(self, verbose: bool, lang: str) -> Dict[str, Any]:
        """Codex: Less is More 최소주의 가이드 - 3-5줄 핵심만"""
        # Codex 원칙: 과도한 설명 없이 필수 정보만
        guide = {"cmd": "oa excel [operation] --format json", "flow": "workbook-list → table-list → operate", "out": "json"}

        if verbose:
            guide["ex"] = "oa excel range-read --range A1:C10"
            guide["conn"] = ["auto", "--file-path", "--workbook-name"]

        return guide

    def _claude_guide(self, verbose: bool, lang: str) -> Dict[str, Any]:
        """Claude: 체계적 워크플로우 가이드 - 안전성과 단계별 접근"""
        guide = {
            "systematic_workflow": [
                {"step": 1, "action": "discover", "command": "oa excel workbook-list", "validation": "열린 워크북 존재 확인"},
                {"step": 2, "action": "analyze", "command": "oa excel table-list", "validation": "데이터 구조 및 타입 확인"},
                {
                    "step": 3,
                    "action": "plan",
                    "command": "oa excel workbook-info",
                    "validation": "작업 대상 시트 및 범위 확정",
                },
                {"step": 4, "action": "execute", "command": "oa excel [operation]", "validation": "결과 검증 및 에러 처리"},
            ],
            "safety_principles": [
                "항상 상태 확인 후 작업 수행",
                "workbook-name 연결로 효율성 확보",
                "데이터 변경 전 백업 고려",
                "에러 발생 시 명확한 진단 제공",
            ],
            "connection_strategy": {
                "preferred": "--workbook-name (효율적 재사용)",
                "fallback": "--file-path (파일 직접 열기)",
                "automatic": "옵션 없음 (활성 워크북)",
            },
        }

        if verbose:
            guide["detailed_workflow"] = {
                "context_discovery": [
                    "workbook-list로 전체 컨텍스트 파악",
                    "workbook-info로 상세 구조 분석",
                    "table-list로 즉시 사용 가능한 데이터 확인",
                ],
                "smart_execution": [
                    "table-driven 작업으로 구조화된 데이터 활용",
                    "range-convert로 데이터 정제 후 피벗/차트 생성",
                    "auto-position으로 겹침 없는 객체 배치",
                ],
            }
            guide["error_recovery"] = {
                "file_not_found": "workbook-list 재확인 → 경로 검증",
                "sheet_missing": "workbook-info 조회 → 정확한 시트명 사용",
                "range_error": "데이터 경계 확인 → 안전한 범위 재설정",
                "com_error": "Excel 애플리케이션 상태 진단 → 재시작 권장",
            }

        return guide

    def _gemini_guide(self, verbose: bool, lang: str) -> Dict[str, Any]:
        """Gemini: 대화형 상호작용 가이드 - 시각화 중심, 스마트 제안"""
        guide = {
            "conversational_flow": {
                "greeting": "Excel 자동화 작업을 도와드리겠습니다",
                "discovery": "현재 상황을 파악해보겠습니다: workbook-list",
                "analysis": "데이터를 분석해보겠습니다: table-list",
                "suggestion": "발견한 데이터를 바탕으로 다음 작업을 제안합니다",
                "execution": "선택하신 작업을 수행하겠습니다",
            },
            "smart_suggestions": {
                "sales_data": {
                    "detected_patterns": ["매출", "판매량", "지역", "월별"],
                    "recommended_actions": [
                        "pivot-create로 지역별 집계",
                        "chart-add로 트렌드 시각화",
                        "data-analyze로 패턴 감지",
                    ],
                },
                "time_series": {
                    "detected_patterns": ["날짜", "월", "년도", "시계열"],
                    "recommended_actions": ["chart-add --type line으로 추세 분석", "data-transform으로 날짜 형식 통일"],
                },
                "large_dataset": {
                    "detected_patterns": ["1000+ 행", "다중 컬럼"],
                    "recommended_actions": [
                        "table-create로 Excel Table 변환",
                        "pivot-create로 요약 분석",
                        "slicer-add로 대화형 필터",
                    ],
                },
            },
            "visualization_priority": {
                "primary": ["chart-add", "pivot-create"],
                "secondary": ["slicer-add", "data-analyze"],
                "advanced": ["chart-pivot-create", "shape-add"],
            },
        }

        if verbose:
            guide["batch_operations"] = [
                {
                    "scenario": "데이터 탐색 및 분석",
                    "commands": ["oa excel workbook-list", "oa excel table-list", "oa excel data-analyze --range A1:Z1000"],
                    "purpose": "전체 데이터 구조와 품질 파악",
                },
                {
                    "scenario": "시각화 대시보드 생성",
                    "commands": [
                        "oa excel range-convert --remove-comma",
                        "oa excel pivot-create --auto-position",
                        "oa excel chart-pivot-create --chart-type column",
                    ],
                    "purpose": "데이터 정제부터 차트까지 원스톱",
                },
            ]
            guide["interactive_features"] = {
                "context_awareness": "이전 작업 기억하고 연속 제안",
                "smart_defaults": "데이터 패턴 기반 옵션 자동 선택",
                "multi_step_planning": "복잡한 작업을 단계별로 분해",
                "visual_feedback": "차트와 테이블 결과 즉시 확인",
            }

        return guide

    def to_markdown(self, guide: Dict[str, Any]) -> str:
        """가이드를 마크다운 형식으로 변환"""
        ai_type = guide.get("meta", {}).get("ai_type", "unknown")
        description = guide.get("meta", {}).get("description", "")

        md_lines = [f"# {ai_type.upper()} AI 어시스턴트 가이드", "", f"**{description}**", ""]

        # AI별 특화 내용을 마크다운으로 변환
        if ai_type == "codex":
            md_lines.extend(
                [
                    "## 최소주의 원칙 (Less is More)",
                    "",
                    f"```bash",
                    f"{guide.get('cmd', '')}",
                    f"```",
                    "",
                    f"**워크플로우:** {guide.get('flow', '')}",
                    f"**출력:** {guide.get('out', '')}",
                ]
            )
        elif ai_type == "claude":
            md_lines.extend(["## 체계적 워크플로우", "", "### 단계별 실행"])
            for step in guide.get("systematic_workflow", []):
                md_lines.append(f"{step['step']}. **{step['action']}**: `{step['command']}`")

            md_lines.extend(["", "### 안전 원칙"])
            for principle in guide.get("safety_principles", []):
                md_lines.append(f"- {principle}")

        # 기타 공통 내용 추가...

        return "\n".join(md_lines)

    def to_text(self, guide: Dict[str, Any]) -> str:
        """가이드를 일반 텍스트 형식으로 변환"""
        ai_type = guide.get("meta", {}).get("ai_type", "unknown")
        description = guide.get("meta", {}).get("description", "")

        lines = [f"=== {ai_type.upper()} AI 어시스턴트 가이드 ===", "", description, ""]

        # JSON 구조를 읽기 쉬운 텍스트로 변환
        def format_dict(d, indent=0):
            result = []
            for key, value in d.items():
                if key == "meta":
                    continue
                prefix = "  " * indent
                if isinstance(value, dict):
                    result.append(f"{prefix}{key}:")
                    result.extend(format_dict(value, indent + 1))
                elif isinstance(value, list):
                    result.append(f"{prefix}{key}:")
                    for item in value:
                        if isinstance(item, dict):
                            result.extend(format_dict(item, indent + 1))
                        else:
                            result.append(f"{prefix}  - {item}")
                else:
                    result.append(f"{prefix}{key}: {value}")
            return result

        lines.extend(format_dict(guide))
        return "\n".join(lines)
