"""
간단한 MCP HTTP 서버 - 디버깅용
"""

import uvicorn
from fastapi import FastAPI
from fastapi.responses import JSONResponse

app = FastAPI(title="Simple MCP Server")


@app.get("/")
async def root():
    return {"message": "Simple MCP Server", "status": "running"}


@app.get("/mcp")
async def mcp_info():
    """MCP 서버 정보"""
    from pyhub_office_automation.version import get_version

    return {
        "server": {
            "name": "PyHub Office Automation MCP",
            "version": get_version(),
            "instructions": "Excel 분석을 위한 최소 도구 세트",
        },
        "capabilities": {"resources": 2, "tools": 5, "prompts": 0},
        "transport": "streamable_http",
        "status": "running",
    }


@app.get("/health")
async def health():
    return {"status": "healthy"}


if __name__ == "__main__":
    uvicorn.run(app, host="0.0.0.0", port=8765)
