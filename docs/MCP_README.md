# PyHub Office Automation MCP Server

Excel 분석을 위한 MCP (Model Context Protocol) 서버 구현

## 🚀 주요 특징

- **최소 도구 세트**: 분석에 필수적인 5개 도구만 제공하여 빠른 로딩
- **Streamable HTTP 전송**: 최신 MCP 2025-03-26 표준 적용
- **FastMCP 기반**: 프로덕션 레벨 프레임워크 사용
- **기존 코드 재사용**: pyhub-office-automation CLI 기능 그대로 활용

## 🔧 제공 도구

### Resources (2개)
1. `resource://excel/workbooks` - 열린 워크북 목록
2. `resource://excel/workbook/{name}/info` - 특정 워크북 상세 정보

### Tools (5개)
1. `excel_workbook_info` - 워크북 구조 및 정보 분석
2. `excel_range_read` - 셀 범위 데이터 읽기
3. `excel_table_read` - 테이블 데이터 읽기
4. `excel_data_analyze` - 데이터 구조 분석
5. `excel_chart_list` - 차트 목록 조회

## 📦 설치 및 실행

### 1. 패키지 설치
```bash
uv pip install fastmcp fastapi uvicorn
```

### 2. MCP 서버 시작
```bash
# 기본 실행 (localhost:8765)
oa mcp start

# 포트 변경
oa mcp start --port 8080

# 외부 접근 허용
oa mcp start --host 0.0.0.0
```

### 3. 서버 정보 확인
```bash
oa mcp info
```

### 4. 기본 테스트
```bash
oa mcp test
```

## 🌐 Claude Desktop 연동

Claude Desktop 설정에서 MCP 서버 추가:
- **URL**: `http://localhost:8765/mcp`
- **전송 방식**: Streamable HTTP
- **인증**: 불필요 (개발용)

## 📋 사용 예제

### 1. 워크북 정보 조회
```
워크북 정보를 알려줘
```
→ MCP 서버가 `excel_workbook_info` 도구 사용

### 2. 데이터 읽기
```
A1:C10 범위의 데이터를 읽어줘
```
→ MCP 서버가 `excel_range_read` 도구 사용

### 3. 데이터 분석
```
현재 데이터 구조를 분석해서 피벗테이블 추천해줘
```
→ MCP 서버가 `excel_data_analyze` 도구 사용

## 🔧 개발 모드

```bash
# 자동 재로드로 실행
oa mcp start --reload

# 서버 직접 테스트
python -m pyhub_office_automation.mcp.server

# HTTP 엔드포인트 확인
curl http://localhost:8765/mcp
curl http://localhost:8765/health
```

## 📊 API 엔드포인트

- `POST /mcp` - MCP Streamable HTTP 엔드포인트
- `GET /mcp` - 서버 정보 조회
- `GET /health` - 헬스 체크
- `GET /docs` - FastAPI 자동 문서

## 🎯 성공 지표

✅ **완료된 기능들:**
- [x] MCP 서버 기본 구현
- [x] 5개 핵심 도구 구현
- [x] Streamable HTTP 전송 지원
- [x] CLI 명령어 통합
- [x] 기본 테스트 통과
- [x] 에러 처리 및 로깅
- [x] 기존 버전 시스템 통합

## 🔮 향후 계획

### Phase 2: 확장 기능 (필요시)
- 쓰기 도구 3개 추가 (`excel_range_write`, `excel_sheet_add`, `excel_chart_add`)
- 고급 기능 (`pivot_create`, `table_transform`)

### Phase 3: 프로덕션 배포 (필요시)
- Docker 컨테이너화
- 클라우드 배포 (Google Cloud Run)
- 인증 및 보안 강화

## 💡 핵심 이점

1. **빠른 로딩**: 최소 도구로 초기화 시간 단축
2. **안정성**: 검증된 읽기 기능 중심
3. **실용성**: 분석 작업에 필수 기능만 제공
4. **확장 가능**: 필요시 점진적 기능 추가
5. **표준 준수**: 최신 MCP 표준 적용

---

**버전**: 9.2539.26
**생성일**: 2025-09-24
**최종 업데이트**: 2025-09-24