# Connector와 MCP 분석 리포트

## 📋 개요

이 리포트는 OpenAI ChatGPT의 Connector 기능과 Anthropic Claude의 MCP(Model Context Protocol)를 분석하고, `pyhub-office-automation` 프로젝트와의 통합 방안을 제시합니다.

### 주요 질문들
- OpenAI Connector와 Claude MCP의 차이점은 무엇인가?
- pyhub-office-automation 프로젝트에서 어떻게 구현해야 하는가?
- 네트워크 및 보안 제약사항은 무엇인가?

---

## 🔌 Connector vs MCP 비교 분석

### OpenAI ChatGPT Connectors

#### 개념
ChatGPT Connectors는 ChatGPT가 서드파티 애플리케이션(Google Drive, GitHub, SharePoint 등)에 안전하게 연결하여 파일 검색, 실시간 데이터 가져오기, 콘텐츠 참조를 채팅 내에서 직접 수행할 수 있게 하는 기능입니다.

#### 주요 특징
- **OAuth 기반 인증**: 사용자가 직접 로그인하여 권한 부여
- **Synced Connectors**: 선택된 콘텐츠를 사전에 동기화하여 인덱싱
- **자동 연결**: GPT-5에서는 Gmail, Google Calendar, Google Contacts를 자동으로 사용
- **파일 타입 지원**: TXT, PDF, CSV, XLSX, PPTX, DOCX 등

#### 지원 플랜
- ChatGPT Plus/Pro (개인)
- ChatGPT Enterprise/Edu (조직)
- 무료 플랜에서는 사용 불가

#### 기술적 제약사항
- OpenAI의 Connectors 프로토콜을 따라야 함
- 제한된 기본 Connectors 제공
- 커스텀 Connectors는 MCP를 통해서만 구현 가능

### Anthropic Claude MCP

#### 개념
MCP(Model Context Protocol)는 AI 어시스턴트를 데이터가 있는 시스템(콘텐츠 저장소, 비즈니스 도구, 개발 환경)에 연결하기 위한 오픈 표준입니다.

#### 주요 특징
- **오픈 표준**: USB-C처럼 표준화된 연결 방식
- **JSON-RPC 기반**: 구조화된 메시지 프로토콜
- **로컬/원격 서버**: 로컬 실행 또는 인터넷 호스팅 가능
- **3가지 핵심 요소**:
  - **Resources**: 구조화된 데이터 (문서 조각, 코드 조각)
  - **Tools**: 실행 가능한 함수 (DB 쿼리, 웹 검색, 메시지 전송)
  - **Prompts**: 준비된 지침 또는 템플릿

#### 아키텍처
```
MCP Client (Claude Desktop) ↔ MCP Server (Your Service)
```

#### 기술적 우위
- **표준화**: 2025년 3월 OpenAI도 MCP 표준 채택
- **유연성**: 로컬/원격 배포 모두 지원
- **확장성**: 다양한 전송 프로토콜 지원 (stdio, HTTP+SSE, WebSocket)

---

## 🏗️ MCP 아키텍처 심화 분석

### 프로토콜 구조

#### 1. JSON-RPC 메시지 형식
```json
{
  "jsonrpc": "2.0",
  "method": "tools/call",
  "params": {
    "name": "excel_read_range",
    "arguments": {
      "file_path": "data.xlsx",
      "sheet": "Sheet1",
      "range": "A1:C10"
    }
  },
  "id": "request-1"
}
```

#### 2. 전송 메커니즘

**Local Transport (stdio)**
```bash
# 로컬 서버 실행
python mcp_server.py
```

**Remote Transport (HTTP+SSE)**
```http
POST /mcp HTTP/1.1
Host: your-server.com
Content-Type: application/json
Accept: application/json, text/event-stream

{메시지 내용}
```

**Streamable HTTP (2025-03-26 최신 스펙)**
```http
# POST로 메시지 전송 + GET으로 SSE 스트림
GET /mcp HTTP/1.1
Accept: text/event-stream
```

#### 3. 세션 관리
- **Session ID**: HTTP 헤더를 통한 세션 식별
- **State Management**: 상태 저장 세션 프로토콜
- **Connection Persistence**: SSE를 통한 지속 연결

---

## 🚀 pyhub-office-automation MCP 통합 방안

### 현재 프로젝트 분석

#### 기존 CLI 구조
```
oa (office automation)
├── excel <command>     # Excel 자동화
│   ├── workbook-open
│   ├── range-read
│   ├── chart-add
│   └── ...
└── hwp <command>       # HWP 자동화
    ├── doc-create
    ├── text-insert
    └── ...
```

#### MCP 통합 아키텍처
```
Claude Desktop ↔ pyhub-office-automation MCP Server ↔ Office Applications
                                                    ├── Excel (xlwings)
                                                    └── HWP (pyhwpx)
```

### 구현 전략

#### 1. MCP Server 기본 구조

```python
# mcp_server.py
import asyncio
from mcp import Server, types
from mcp.server.models import InitializationOptions
from mcp.server import NotificationOptions, Server
from mcp.server.stdio import stdio_server
from mcp.types import (
    Resource,
    Tool,
    TextContent,
    ImageContent,
    EmbeddedResource,
)
import json
import subprocess
from pathlib import Path

class OfficeAutomationMCPServer:
    def __init__(self):
        self.server = Server("pyhub-office-automation")
        self.setup_handlers()

    def setup_handlers(self):
        """MCP 핸들러 설정"""

        @self.server.list_resources()
        async def handle_list_resources() -> list[types.Resource]:
            """사용 가능한 리소스 목록 반환"""
            return [
                types.Resource(
                    uri="workbook://active",
                    name="Active Workbook Info",
                    description="현재 활성화된 Excel 워크북 정보",
                    mimeType="application/json",
                ),
                types.Resource(
                    uri="workbook://list",
                    name="Open Workbooks List",
                    description="현재 열린 모든 Excel 워크북 목록",
                    mimeType="application/json",
                )
            ]

        @self.server.read_resource()
        async def handle_read_resource(uri: str) -> str:
            """리소스 데이터 읽기"""
            if uri == "workbook://active":
                result = subprocess.run([
                    "oa", "excel", "workbook-info", "--include-sheets"
                ], capture_output=True, text=True)
                return result.stdout

            elif uri == "workbook://list":
                result = subprocess.run([
                    "oa", "excel", "workbook-list", "--detailed"
                ], capture_output=True, text=True)
                return result.stdout

            raise ValueError(f"Unknown resource: {uri}")

        @self.server.list_tools()
        async def handle_list_tools() -> list[types.Tool]:
            """사용 가능한 도구 목록 반환"""
            return [
                types.Tool(
                    name="excel_read_range",
                    description="Excel 시트에서 범위 데이터 읽기",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "file_path": {
                                "type": "string",
                                "description": "Excel 파일 경로"
                            },
                            "sheet": {
                                "type": "string",
                                "description": "시트명"
                            },
                            "range": {
                                "type": "string",
                                "description": "읽을 범위 (예: A1:C10)"
                            }
                        },
                        "required": ["sheet", "range"]
                    }
                ),
                types.Tool(
                    name="excel_write_range",
                    description="Excel 시트에 데이터 쓰기",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "file_path": {"type": "string"},
                            "sheet": {"type": "string"},
                            "range": {"type": "string"},
                            "data": {
                                "type": "array",
                                "description": "쓸 데이터 (2차원 배열)"
                            }
                        },
                        "required": ["sheet", "range", "data"]
                    }
                ),
                types.Tool(
                    name="excel_create_chart",
                    description="Excel 차트 생성",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "sheet": {"type": "string"},
                            "data_range": {"type": "string"},
                            "chart_type": {
                                "type": "string",
                                "enum": ["Column", "Line", "Pie", "Bar", "Area"]
                            },
                            "title": {"type": "string"}
                        },
                        "required": ["sheet", "data_range", "chart_type"]
                    }
                ),
                types.Tool(
                    name="hwp_create_document",
                    description="HWP 문서 생성",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "file_path": {"type": "string"},
                            "content": {"type": "string"}
                        },
                        "required": ["file_path", "content"]
                    }
                )
            ]

        @self.server.call_tool()
        async def handle_call_tool(name: str, arguments: dict) -> list[types.TextContent]:
            """도구 실행"""
            if name == "excel_read_range":
                return await self.excel_read_range(**arguments)
            elif name == "excel_write_range":
                return await self.excel_write_range(**arguments)
            elif name == "excel_create_chart":
                return await self.excel_create_chart(**arguments)
            elif name == "hwp_create_document":
                return await self.hwp_create_document(**arguments)
            else:
                raise ValueError(f"Unknown tool: {name}")

    async def excel_read_range(self, sheet: str, range: str, file_path: str = None) -> list[types.TextContent]:
        """Excel 범위 읽기"""
        cmd = ["oa", "excel", "range-read", "--sheet", sheet, "--range", range]
        if file_path:
            cmd.extend(["--file-path", file_path])

        result = subprocess.run(cmd, capture_output=True, text=True)

        if result.returncode != 0:
            error_msg = f"Excel 읽기 실패: {result.stderr}"
            return [types.TextContent(type="text", text=error_msg)]

        return [types.TextContent(
            type="text",
            text=f"Excel 데이터 읽기 완료:\n{result.stdout}"
        )]

    async def excel_write_range(self, sheet: str, range: str, data: list, file_path: str = None) -> list[types.TextContent]:
        """Excel 데이터 쓰기"""
        data_json = json.dumps(data)
        cmd = ["oa", "excel", "range-write", "--sheet", sheet, "--range", range, "--data", data_json]
        if file_path:
            cmd.extend(["--file-path", file_path])

        result = subprocess.run(cmd, capture_output=True, text=True)

        if result.returncode != 0:
            error_msg = f"Excel 쓰기 실패: {result.stderr}"
            return [types.TextContent(type="text", text=error_msg)]

        return [types.TextContent(
            type="text",
            text=f"Excel 데이터 쓰기 완료:\n{result.stdout}"
        )]

    async def excel_create_chart(self, sheet: str, data_range: str, chart_type: str, title: str = None) -> list[types.TextContent]:
        """Excel 차트 생성"""
        cmd = ["oa", "excel", "chart-add", "--sheet", sheet, "--data-range", data_range, "--chart-type", chart_type]
        if title:
            cmd.extend(["--title", title])

        result = subprocess.run(cmd, capture_output=True, text=True)

        if result.returncode != 0:
            error_msg = f"차트 생성 실패: {result.stderr}"
            return [types.TextContent(type="text", text=error_msg)]

        return [types.TextContent(
            type="text",
            text=f"차트 생성 완료:\n{result.stdout}"
        )]

    async def hwp_create_document(self, file_path: str, content: str) -> list[types.TextContent]:
        """HWP 문서 생성"""
        cmd = ["oa", "hwp", "doc-create", "--file-path", file_path, "--content", content]

        result = subprocess.run(cmd, capture_output=True, text=True)

        if result.returncode != 0:
            error_msg = f"HWP 문서 생성 실패: {result.stderr}"
            return [types.TextContent(type="text", text=error_msg)]

        return [types.TextContent(
            type="text",
            text=f"HWP 문서 생성 완료:\n{result.stdout}"
        )]

    async def run(self):
        """서버 실행"""
        async with stdio_server() as (read_stream, write_stream):
            await self.server.run(
                read_stream, write_stream,
                InitializationOptions(
                    server_name="pyhub-office-automation",
                    server_version="1.0.0",
                    capabilities=self.server.get_capabilities(
                        notification_options=NotificationOptions(),
                        experimental_capabilities={}
                    )
                )
            )

async def main():
    server = OfficeAutomationMCPServer()
    await server.run()

if __name__ == "__main__":
    asyncio.run(main())
```

#### 2. 패키지 구조 확장

```
pyhub_office_automation/
├── cli/                    # 기존 CLI 명령어
├── excel/                  # Excel 자동화 스크립트
├── hwp/                    # HWP 자동화 스크립트
├── mcp/                    # MCP 통합 새 디렉토리
│   ├── __init__.py
│   ├── server.py          # MCP 서버 구현
│   ├── handlers/          # 각 기능별 핸들러
│   │   ├── excel_handler.py
│   │   ├── hwp_handler.py
│   │   └── resource_handler.py
│   └── config/            # MCP 설정
│       ├── server_config.json
│       └── tools_schema.json
└── requirements.txt       # MCP 관련 의존성 추가
```

#### 3. 의존성 추가

```python
# requirements.txt에 추가
mcp>=1.2.0
fastapi>=0.104.0          # HTTP 서버용
uvicorn>=0.24.0           # ASGI 서버
sse-starlette>=1.6.5      # SSE 지원
python-multipart>=0.0.6   # 파일 업로드
```

---

## 🌐 네트워크 및 배포 요구사항

### 프로토콜 요구사항

#### 1. HTTP Endpoint 구성
```python
# mcp/http_server.py
from fastapi import FastAPI, HTTPException
from fastapi.responses import StreamingResponse
from sse_starlette.sse import EventSourceResponse
import json
import asyncio
from typing import AsyncGenerator

app = FastAPI(title="pyhub-office-automation MCP Server")

@app.post("/mcp")
async def handle_mcp_post(request: dict):
    """JSON-RPC 메시지 처리"""
    # MCP 서버로 요청 전달
    response = await process_mcp_request(request)
    return response

@app.get("/mcp")
async def handle_mcp_sse():
    """Server-Sent Events 스트림"""
    async def event_stream() -> AsyncGenerator[dict, None]:
        while True:
            # MCP 서버에서 이벤트 수신
            event = await get_mcp_event()
            if event:
                yield {"data": json.dumps(event)}
            await asyncio.sleep(0.1)

    return EventSourceResponse(event_stream())

@app.get("/health")
async def health_check():
    """서버 상태 체크"""
    return {"status": "healthy", "version": "1.0.0"}
```

#### 2. 인증 및 보안

```python
# mcp/auth.py
from fastapi import HTTPException, Depends, Header
from fastapi.security import HTTPBearer, HTTPAuthorizationCredentials
import jwt
from typing import Optional

security = HTTPBearer()

async def verify_origin(origin: str = Header(None)):
    """Origin 헤더 검증 (DNS rebinding 공격 방지)"""
    allowed_origins = [
        "https://claude.ai",
        "https://desktop.claude.ai",
        "http://localhost:*"  # 개발용
    ]

    if not origin or not any(origin.startswith(allowed) for allowed in allowed_origins):
        raise HTTPException(status_code=403, detail="Invalid origin")

    return origin

async def verify_token(credentials: HTTPAuthorizationCredentials = Depends(security)):
    """JWT 토큰 검증"""
    try:
        payload = jwt.decode(
            credentials.credentials,
            "your-secret-key",
            algorithms=["HS256"]
        )
        return payload
    except jwt.InvalidTokenError:
        raise HTTPException(status_code=401, detail="Invalid token")

# OAuth 2.0 Resource Server 구성
async def verify_oauth_token(authorization: str = Header(None)):
    """OAuth 2.0 토큰 검증"""
    if not authorization or not authorization.startswith("Bearer "):
        raise HTTPException(status_code=401, detail="Missing or invalid authorization header")

    token = authorization.replace("Bearer ", "")
    # 토큰 검증 로직 (RFC 8707 Resource Indicators 지원)
    return await validate_oauth_token(token)
```

#### 3. 배포 설정

**Docker 컨테이너**
```dockerfile
# Dockerfile
FROM python:3.13-slim

WORKDIR /app

# 의존성 설치
COPY requirements.txt .
RUN pip install -r requirements.txt

# 애플리케이션 코드
COPY pyhub_office_automation/ ./pyhub_office_automation/

# MCP 서버 포트
EXPOSE 8000

# 서버 실행
CMD ["uvicorn", "pyhub_office_automation.mcp.http_server:app", "--host", "0.0.0.0", "--port", "8000"]
```

**클라우드 배포 (Google Cloud Run)**
```yaml
# cloudbuild.yaml
steps:
- name: 'gcr.io/cloud-builders/docker'
  args: ['build', '-t', 'gcr.io/$PROJECT_ID/pyhub-office-mcp:$COMMIT_SHA', '.']
- name: 'gcr.io/cloud-builders/docker'
  args: ['push', 'gcr.io/$PROJECT_ID/pyhub-office-mcp:$COMMIT_SHA']
- name: 'gcr.io/google.com/cloudsdktool/cloud-sdk'
  entrypoint: 'gcloud'
  args: [
    'run', 'deploy', 'pyhub-office-mcp',
    '--image', 'gcr.io/$PROJECT_ID/pyhub-office-mcp:$COMMIT_SHA',
    '--platform', 'managed',
    '--region', 'us-central1',
    '--allow-unauthenticated',
    '--port', '8000',
    '--memory', '1Gi',
    '--timeout', '300'
  ]
```

### 성능 및 확장성 고려사항

#### 1. 세션 관리
```python
# mcp/session_manager.py
import asyncio
from typing import Dict, Optional
import uuid
from dataclasses import dataclass
from datetime import datetime, timedelta

@dataclass
class MCPSession:
    session_id: str
    client_info: dict
    created_at: datetime
    last_activity: datetime
    state: dict

class SessionManager:
    def __init__(self):
        self.sessions: Dict[str, MCPSession] = {}
        self.cleanup_task = asyncio.create_task(self.cleanup_expired_sessions())

    def create_session(self, client_info: dict) -> str:
        """새 세션 생성"""
        session_id = str(uuid.uuid4())
        session = MCPSession(
            session_id=session_id,
            client_info=client_info,
            created_at=datetime.now(),
            last_activity=datetime.now(),
            state={}
        )
        self.sessions[session_id] = session
        return session_id

    async def cleanup_expired_sessions(self):
        """만료된 세션 정리"""
        while True:
            current_time = datetime.now()
            expired_sessions = [
                session_id for session_id, session in self.sessions.items()
                if current_time - session.last_activity > timedelta(hours=1)
            ]

            for session_id in expired_sessions:
                del self.sessions[session_id]

            await asyncio.sleep(300)  # 5분마다 정리
```

#### 2. 부하 분산 및 스케일링
```python
# mcp/load_balancer.py
from typing import List
import aiohttp
import asyncio
from dataclasses import dataclass

@dataclass
class MCPServerInstance:
    url: str
    health_score: float
    active_connections: int

class LoadBalancer:
    def __init__(self, server_instances: List[str]):
        self.instances = [
            MCPServerInstance(url=url, health_score=1.0, active_connections=0)
            for url in server_instances
        ]
        self.health_check_task = asyncio.create_task(self.health_check_loop())

    async def get_best_instance(self) -> MCPServerInstance:
        """가장 적합한 인스턴스 선택"""
        available_instances = [
            instance for instance in self.instances
            if instance.health_score > 0.5
        ]

        if not available_instances:
            raise Exception("No healthy instances available")

        # 연결 수와 헬스 스코어를 기반으로 선택
        return min(available_instances,
                  key=lambda x: x.active_connections / x.health_score)

    async def health_check_loop(self):
        """인스턴스 상태 모니터링"""
        while True:
            for instance in self.instances:
                try:
                    async with aiohttp.ClientSession() as session:
                        async with session.get(f"{instance.url}/health", timeout=5) as response:
                            if response.status == 200:
                                instance.health_score = min(1.0, instance.health_score + 0.1)
                            else:
                                instance.health_score = max(0.0, instance.health_score - 0.2)
                except Exception:
                    instance.health_score = max(0.0, instance.health_score - 0.3)

            await asyncio.sleep(30)  # 30초마다 헬스 체크
```

---

## ⚠️ 제약사항 및 고려사항

### 1. 기술적 제약사항

#### Windows COM 의존성
- **문제**: Excel/HWP 자동화가 Windows COM API에 의존
- **제약**: 리눅스/macOS 서버에서는 Excel/HWP 직접 조작 불가
- **해결방안**:
  ```python
  # mcp/platform_handler.py
  import platform
  import subprocess
  from typing import Optional

  class PlatformHandler:
      @staticmethod
      def is_windows() -> bool:
          return platform.system() == "Windows"

      @staticmethod
      def get_excel_handler():
          if PlatformHandler.is_windows():
              return WindowsExcelHandler()
          else:
              return CloudExcelHandler()  # 클라우드 Excel API 사용

      @staticmethod
      def get_hwp_handler():
          if PlatformHandler.is_windows():
              return WindowsHWPHandler()
          else:
              raise NotImplementedError("HWP is only supported on Windows")
  ```

#### 파일 시스템 접근
- **보안 위험**: 원격 서버에서 로컬 파일 시스템 접근
- **해결방안**:
  - 샌드박스 환경에서 실행
  - 파일 업로드/다운로드 API 제공
  - 임시 파일 자동 삭제

#### 메모리 및 리소스 관리
```python
# mcp/resource_manager.py
import psutil
import asyncio
from contextlib import asynccontextmanager

class ResourceManager:
    def __init__(self, max_memory_mb: int = 512, max_cpu_percent: int = 80):
        self.max_memory_mb = max_memory_mb
        self.max_cpu_percent = max_cpu_percent

    @asynccontextmanager
    async def resource_limit(self):
        """리소스 사용량 모니터링"""
        initial_memory = psutil.virtual_memory().available

        try:
            yield
        finally:
            # 메모리 정리
            import gc
            gc.collect()

            current_memory = psutil.virtual_memory().available
            memory_used = (initial_memory - current_memory) / (1024 * 1024)  # MB

            if memory_used > self.max_memory_mb:
                print(f"Warning: High memory usage: {memory_used:.2f}MB")

    async def check_system_resources(self):
        """시스템 리소스 체크"""
        memory = psutil.virtual_memory()
        cpu = psutil.cpu_percent()

        if memory.percent > 90 or cpu > self.max_cpu_percent:
            raise Exception("System resources exceeded")
```

### 2. 네트워크 제약사항

#### 방화벽 및 포트 설정
- **MCP 서버 포트**: 기본 8000번 (설정 가능)
- **SSE 연결**: 지속적인 HTTP 연결 필요
- **WebSocket**: 실시간 양방향 통신용 (선택사항)

#### 대역폭 및 지연시간
```python
# mcp/network_monitor.py
import time
import asyncio
from dataclasses import dataclass
from typing import List

@dataclass
class NetworkMetrics:
    latency_ms: float
    bandwidth_mbps: float
    packet_loss_percent: float

class NetworkMonitor:
    def __init__(self):
        self.metrics_history: List[NetworkMetrics] = []

    async def measure_latency(self, client_ip: str) -> float:
        """클라이언트와의 지연시간 측정"""
        start_time = time.time()
        # 핑 테스트 또는 HTTP 응답 시간 측정
        await asyncio.sleep(0)  # 실제로는 네트워크 요청
        end_time = time.time()

        return (end_time - start_time) * 1000  # ms

    def is_connection_stable(self) -> bool:
        """네트워크 연결 안정성 확인"""
        if len(self.metrics_history) < 5:
            return True

        recent_metrics = self.metrics_history[-5:]
        avg_latency = sum(m.latency_ms for m in recent_metrics) / len(recent_metrics)

        return avg_latency < 1000  # 1초 이하
```

### 3. 보안 제약사항

#### 데이터 프라이버시
- **문제**: 민감한 문서 내용이 네트워크를 통해 전송
- **해결방안**:
  - 종단간 암호화 (TLS 1.3)
  - 문서 내용 요약만 전송
  - 로컬 처리 우선순위

#### 인증 및 권한 관리
```python
# mcp/security.py
from cryptography.fernet import Fernet
import hashlib
import hmac
from typing import Optional

class SecurityManager:
    def __init__(self, secret_key: str):
        self.secret_key = secret_key.encode()
        self.cipher = Fernet(Fernet.generate_key())

    def encrypt_data(self, data: str) -> str:
        """민감한 데이터 암호화"""
        return self.cipher.encrypt(data.encode()).decode()

    def decrypt_data(self, encrypted_data: str) -> str:
        """암호화된 데이터 복호화"""
        return self.cipher.decrypt(encrypted_data.encode()).decode()

    def verify_request_signature(self, payload: str, signature: str) -> bool:
        """요청 서명 검증"""
        expected_signature = hmac.new(
            self.secret_key,
            payload.encode(),
            hashlib.sha256
        ).hexdigest()

        return hmac.compare_digest(signature, expected_signature)

    def sanitize_file_path(self, file_path: str) -> Optional[str]:
        """파일 경로 검증 및 정리"""
        # 디렉토리 순회 공격 방지
        if ".." in file_path or file_path.startswith("/"):
            return None

        # 허용된 확장자만 허용
        allowed_extensions = [".xlsx", ".docx", ".hwp", ".csv"]
        if not any(file_path.lower().endswith(ext) for ext in allowed_extensions):
            return None

        return file_path
```

---

## 📋 실행 계획 및 로드맵

### Phase 1: 기본 MCP 서버 구현 (2-3주)

#### Week 1: 개발 환경 설정
- [ ] MCP Python SDK 설치 및 테스트 환경 구축
- [ ] 기본 MCP 서버 구조 구현
- [ ] stdio 전송을 통한 로컬 테스트

```bash
# 개발 환경 설정
python -m venv venv_mcp
source venv_mcp/bin/activate  # Windows: venv_mcp\Scripts\activate
pip install mcp[cli] fastapi uvicorn

# 기본 서버 테스트
python pyhub_office_automation/mcp/server.py
```

#### Week 2: Excel 통합
- [ ] Excel Resources 구현 (워크북 정보, 시트 목록)
- [ ] Excel Tools 구현 (범위 읽기/쓰기, 차트 생성)
- [ ] 에러 처리 및 로깅

#### Week 3: HWP 통합 및 테스트
- [ ] HWP Tools 구현 (문서 생성, 텍스트 삽입)
- [ ] 통합 테스트 및 버그 수정
- [ ] 문서화 작성

### Phase 2: HTTP/SSE 원격 서버 (2-3주)

#### Week 4-5: HTTP 서버 구현
- [ ] FastAPI 기반 HTTP 엔드포인트 구현
- [ ] SSE (Server-Sent Events) 지원
- [ ] 세션 관리 및 상태 저장

#### Week 6: 보안 및 인증
- [ ] OAuth 2.0 토큰 검증
- [ ] Origin 헤더 검증
- [ ] 요청 서명 검증

### Phase 3: 프로덕션 배포 (2주)

#### Week 7: 클라우드 배포
- [ ] Docker 컨테이너 구성
- [ ] Google Cloud Run 배포 설정
- [ ] 로드 밸런싱 및 Auto Scaling

#### Week 8: 모니터링 및 최적화
- [ ] 헬스 체크 엔드포인트
- [ ] 메트릭 수집 (Prometheus/Grafana)
- [ ] 성능 최적화

### Phase 4: 고도화 기능 (3-4주)

#### Week 9-10: 고급 기능
- [ ] 스트리밍 데이터 처리
- [ ] 비동기 작업 큐 (Celery/Redis)
- [ ] 캐싱 레이어 (Redis)

#### Week 11-12: AI 최적화
- [ ] Claude Desktop 연동 테스트
- [ ] ChatGPT MCP 호환성 확인
- [ ] 사용자 경험 개선

### 구현 우선순위

#### High Priority
1. **Excel 기본 기능**: 범위 읽기/쓰기, 워크북 정보
2. **HTTP 서버**: 원격 접근 가능한 서버 구현
3. **보안**: 기본적인 인증 및 권한 관리

#### Medium Priority
1. **HWP 기능**: 한국어 문서 처리
2. **차트 생성**: 시각화 기능
3. **클라우드 배포**: 확장 가능한 인프라

#### Low Priority
1. **고급 분석**: 복잡한 데이터 처리
2. **실시간 협업**: 다중 사용자 지원
3. **AI 최적화**: 지능형 자동화

### 예상 비용 및 리소스

#### 개발 리소스
- **개발자**: 1-2명 (Python/FastAPI 경험 필요)
- **시간**: 총 8-12주
- **테스트 환경**: Windows 개발 머신 (Excel/HWP 설치)

#### 인프라 비용 (월간)
- **Google Cloud Run**: $20-50
- **Redis 인스턴스**: $15-30
- **모니터링 도구**: $10-20
- **총 예상 비용**: $45-100/월

---

## 🎯 결론 및 권장사항

### 주요 발견사항

1. **MCP의 전략적 중요성**
   - OpenAI와 Anthropic 모두 MCP 표준 채택
   - AI 서비스 통합을 위한 표준 프로토콜로 자리잡음
   - pyhub-office-automation의 경쟁 우위 확보 기회

2. **기술적 실현 가능성**
   - Python MCP SDK 성숙도 높음
   - 기존 CLI 명령어 재사용 가능
   - 단계적 구현을 통한 리스크 최소화

3. **비즈니스 임팩트**
   - AI 어시스턴트와의 직접 통합
   - 사용자 경험 크게 향상
   - 한국어 오피스 자동화 선도 위치 확보

### 권장 구현 전략

#### 1. 최소 실행 가능 제품 (MVP) 우선
- Excel 기본 기능부터 시작
- stdio 전송으로 로컬 테스트
- 점진적 기능 확장

#### 2. 보안 우선 설계
- 데이터 프라이버시 최우선 고려
- 종단간 암호화 구현
- 최소 권한 원칙 적용

#### 3. 확장 가능한 아키텍처
- 마이크로서비스 패턴
- 컨테이너 기반 배포
- 수평적 확장 지원

### 다음 단계

1. **개념 검증 (POC)**
   ```bash
   # 1주차: 기본 MCP 서버 구현 및 테스트
   git checkout -b feature/mcp-integration
   mkdir pyhub_office_automation/mcp
   # 위의 예제 코드로 기본 서버 구현
   ```

2. **이해관계자 리뷰**
   - 기술팀과 아키텍처 검토
   - 보안팀과 리스크 평가
   - 비즈니스팀과 우선순위 조정

3. **프로토타입 개발 시작**
   - Phase 1 실행계획 따라 구현
   - 주간 진척사항 리뷰
   - 사용자 피드백 수집

이 MCP 통합을 통해 pyhub-office-automation은 차세대 AI 워크플로우의 핵심 구성요소가 될 수 있습니다. 한국어 오피스 자동화 분야에서의 선도적 위치를 확보하고, 글로벌 AI 생태계와의 표준 호환성을 달성할 수 있는 중요한 기회입니다.

---

*문서 생성일: 2025-09-24*
*작성자: Claude Code Assistant*
*버전: 1.0*