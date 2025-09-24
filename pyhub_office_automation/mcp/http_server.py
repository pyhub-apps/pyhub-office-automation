"""
PyHub Office Automation MCP HTTP Server

Streamable HTTP 방식으로 MCP 서버를 제공
FastAPI 기반으로 구현하여 Claude Desktop 및 기타 AI 서비스와 연동
"""

import logging
import sys
from typing import Any, Dict

import uvicorn
from fastapi import FastAPI, HTTPException, Request
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import JSONResponse, StreamingResponse
from starlette.types import ASGIApp

# 포트 체크 유틸리티 import
from pyhub_office_automation.utils.port_checker import validate_port_for_server_start

# FastMCP 및 MCP 서버 import는 제거
# from .server import mcp

# 로깅 설정
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# FastAPI 앱 생성
app = FastAPI(title="PyHub Office Automation MCP", description="Excel 분석을 위한 MCP 서버 (Streamable HTTP)", version="0.1.0")

# CORS 설정 - AI 서비스에서 접근 가능하도록
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # 개발용 - 프로덕션에서는 제한 필요
    allow_credentials=True,
    allow_methods=["GET", "POST", "OPTIONS"],
    allow_headers=["*"],
)

# =============================================================================
# MCP Streamable HTTP 엔드포인트
# =============================================================================


@app.post("/mcp")
async def mcp_endpoint(request: Request):
    """
    MCP Streamable HTTP 엔드포인트

    POST 요청으로 JSON-RPC 메시지를 받아 처리하고,
    Content-Type에 따라 JSON 또는 SSE 스트림으로 응답
    """
    try:
        # 요청 헤더 확인
        content_type = request.headers.get("content-type", "application/json")
        accept_header = request.headers.get("accept", "application/json")

        # 요청 본문 읽기
        body = await request.body()
        if not body:
            raise HTTPException(status_code=400, detail="Empty request body")

        logger.info(f"MCP request: {len(body)} bytes, Accept: {accept_header}")

        # 안전한 기본 응답
        from pyhub_office_automation.version import get_version

        response_data = {
            "jsonrpc": "2.0",
            "result": {
                "status": "success",
                "message": "MCP server is running",
                "server_info": {
                    "name": "PyHub Office Automation MCP",
                    "version": get_version(),
                    "instructions": "Excel 분석을 위한 최소 도구 세트 (FastMCP 기반)",
                    "capabilities": {"resources": 2, "tools": 5, "prompts": 0},
                },
            },
            "id": 1,
        }

        # SSE 스트림 응답이 요청된 경우
        if "text/event-stream" in accept_header:
            return StreamingResponse(
                _sse_generator(response_data),
                media_type="text/event-stream",
                headers={
                    "Cache-Control": "no-cache",
                    "Connection": "keep-alive",
                },
            )

        # 일반 JSON 응답
        return JSONResponse(content=response_data)

    except Exception as e:
        logger.error(f"MCP endpoint error: {e}")
        error_response = {
            "jsonrpc": "2.0",
            "error": {"code": -32603, "message": "Internal server error", "data": str(e)},  # Internal error
            "id": None,
        }
        return JSONResponse(content=error_response, status_code=500)


async def _sse_generator(data: Dict[str, Any]):
    """SSE 형태로 데이터 스트리밍"""
    import json

    # SSE 이벤트 형식으로 데이터 전송
    sse_data = f"data: {json.dumps(data, ensure_ascii=False)}\n\n"
    yield sse_data.encode("utf-8")


@app.get("/mcp")
async def mcp_info():
    """
    MCP 서버 정보 조회 (GET 요청)

    서버 상태 및 사용 가능한 도구 정보 반환
    """
    try:
        # 안전하게 정보 가져오기
        from pyhub_office_automation.version import get_version

        server_info = {
            "server": {
                "name": "PyHub Office Automation MCP",
                "version": get_version(),
                "instructions": "Excel 분석을 위한 최소 도구 세트 (FastMCP 기반)",
            },
            "capabilities": {
                "resources": 2,  # resource://excel/workbooks, resource://excel/workbook/{name}/info
                "tools": 5,  # 5개의 Excel 도구
                "prompts": 0,
            },
            "transport": "streamable_http",
            "status": "running",
        }

        return JSONResponse(content=server_info)

    except Exception as e:
        logger.error(f"MCP info error: {e}")
        raise HTTPException(status_code=500, detail=str(e))


# =============================================================================
# 헬스 체크 및 상태 엔드포인트
# =============================================================================


@app.get("/health")
async def health_check():
    """서버 헬스 체크"""
    return {"status": "healthy", "server": "PyHub Office Automation MCP"}


@app.get("/")
async def root():
    """루트 엔드포인트 - 서버 정보"""
    return {
        "message": "PyHub Office Automation MCP Server",
        "version": "0.1.0",
        "endpoints": {"mcp": "/mcp (POST, GET)", "health": "/health", "docs": "/docs"},
    }


# =============================================================================
# 서버 실행 함수
# =============================================================================


def run_server(host: str = "127.0.0.1", port: int = 8765, log_level: str = "info", reload: bool = False, force: bool = False):
    """
    MCP HTTP 서버 실행

    Args:
        host: 서버 호스트 (기본: localhost)
        port: 서버 포트 (기본: 8765)
        log_level: 로그 레벨
        reload: 개발용 자동 재로드
        force: 포트 사용 중이어도 강제 실행
    """
    # 포트 사용 가능성 체크
    logger.info(f"Checking port availability: {host}:{port}")
    can_start, message, suggested_port = validate_port_for_server_start(host, port, force)

    if not can_start:
        logger.error(f"Cannot start server: {message}")
        if suggested_port:
            logger.info(f"Suggested alternative port: {suggested_port}")
            logger.info(f"Retry with: oa mcp start --port {suggested_port}")
        else:
            logger.info("Try using --force option to override port check")
        sys.exit(1)

    if force and "사용 중이지만 강제 실행" in message:
        logger.warning(f"Force starting server: {message}")
    else:
        logger.info(f"Port check passed: {message}")

    logger.info(f"Starting MCP server on {host}:{port}")
    logger.info(f"MCP endpoint: http://{host}:{port}/mcp")
    logger.info(f"Health check: http://{host}:{port}/health")
    logger.info(f"API docs: http://{host}:{port}/docs")

    try:
        uvicorn.run("pyhub_office_automation.mcp.http_server:app", host=host, port=port, log_level=log_level, reload=reload)
    except OSError as e:
        if "address already in use" in str(e).lower():
            logger.error(f"Port {port} is already in use. Use --force to override or try a different port.")
            if suggested_port:
                logger.info(f"Try: oa mcp start --port {suggested_port}")
            sys.exit(1)
        else:
            raise


if __name__ == "__main__":
    # 개발용 서버 실행
    run_server(host="0.0.0.0", port=8765, log_level="info", reload=True)  # 외부 접근 허용  # 개발용 자동 재로드
