"""
포트 사용 체크 유틸리티

크로스 플랫폼으로 포트 사용 상태를 확인하고 충돌을 방지합니다.
"""

import logging
import socket
import sys
from typing import Any, Dict, Optional, Tuple

logger = logging.getLogger(__name__)


def is_port_in_use(host: str, port: int) -> bool:
    """
    지정된 호스트와 포트가 사용 중인지 확인

    Args:
        host: 확인할 호스트 주소
        port: 확인할 포트 번호

    Returns:
        True if port is in use, False otherwise
    """
    try:
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
            sock.settimeout(1)  # 1초 타임아웃
            result = sock.connect_ex((host, port))
            return result == 0  # 0이면 연결 성공 (포트 사용 중)
    except Exception as e:
        logger.error(f"포트 체크 중 오류: {e}")
        return False


def find_available_port(host: str, start_port: int, max_attempts: int = 100) -> Optional[int]:
    """
    사용 가능한 포트를 찾아 반환

    Args:
        host: 확인할 호스트 주소
        start_port: 시작 포트 번호
        max_attempts: 최대 시도 횟수

    Returns:
        사용 가능한 포트 번호 또는 None
    """
    for port in range(start_port, start_port + max_attempts):
        if not is_port_in_use(host, port):
            return port
    return None


def check_port_with_recommendation(host: str, port: int) -> Dict[str, Any]:
    """
    포트 사용 상태를 확인하고 대안 제시

    Args:
        host: 확인할 호스트 주소
        port: 확인할 포트 번호

    Returns:
        포트 상태 및 권장사항 정보
    """
    result = {
        "host": host,
        "requested_port": port,
        "is_available": False,
        "alternative_port": None,
        "message": "",
        "can_proceed": False,
    }

    # 요청된 포트 확인
    if not is_port_in_use(host, port):
        result.update({"is_available": True, "message": f"포트 {port}를 사용할 수 있습니다.", "can_proceed": True})
        return result

    # 포트가 사용 중인 경우 대안 찾기
    result["message"] = f"포트 {port}가 이미 사용 중입니다."

    # 대안 포트 찾기
    alternative = find_available_port(host, port + 1, 50)
    if alternative:
        result.update({"alternative_port": alternative, "message": f"포트 {port}가 이미 사용 중입니다. 대안: {alternative}"})
    else:
        result["message"] = f"포트 {port}가 이미 사용 중이며, 인근 포트도 모두 사용 중입니다."

    return result


def get_port_info(host: str, port: int) -> Dict[str, Any]:
    """
    포트에 대한 상세 정보 반환

    Args:
        host: 확인할 호스트 주소
        port: 확인할 포트 번호

    Returns:
        포트 상세 정보
    """
    info = {
        "host": host,
        "port": port,
        "is_available": not is_port_in_use(host, port),
        "platform": sys.platform,
        "check_method": "socket",
    }

    # OS별 추가 정보 수집 시도
    try:
        if sys.platform == "win32":
            # Windows: netstat으로 프로세스 정보 확인
            import subprocess

            result = subprocess.run(["netstat", "-ano", f"findstr", f":{port}"], capture_output=True, text=True, timeout=5)
            if result.stdout.strip():
                info["process_info"] = result.stdout.strip().split("\n")
        else:
            # Unix/Linux: lsof로 프로세스 정보 확인
            import subprocess

            result = subprocess.run(["lsof", "-i", f":{port}"], capture_output=True, text=True, timeout=5)
            if result.stdout.strip():
                info["process_info"] = result.stdout.strip().split("\n")
    except Exception as e:
        logger.debug(f"프로세스 정보 수집 실패: {e}")
        info["process_info_error"] = str(e)

    return info


def validate_port_for_server_start(host: str, port: int, force: bool = False) -> Tuple[bool, str, Optional[int]]:
    """
    서버 시작 전 포트 유효성 검사

    Args:
        host: 서버 호스트
        port: 서버 포트
        force: 강제 실행 여부

    Returns:
        (can_start: bool, message: str, suggested_port: Optional[int])
    """
    if not (1 <= port <= 65535):
        return False, f"잘못된 포트 번호: {port} (1-65535 범위여야 함)", None

    check_result = check_port_with_recommendation(host, port)

    if check_result["is_available"]:
        return True, f"포트 {port} 사용 가능", None

    if force:
        return True, f"포트 {port} 사용 중이지만 강제 실행", None

    return False, check_result["message"], check_result["alternative_port"]


# CLI 지원
if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="포트 사용 상태 확인")
    parser.add_argument("port", type=int, help="확인할 포트 번호")
    parser.add_argument("--host", default="127.0.0.1", help="확인할 호스트 (기본: 127.0.0.1)")
    parser.add_argument("--find-alternative", action="store_true", help="대안 포트 찾기")
    parser.add_argument("--detailed", action="store_true", help="상세 정보 표시")

    args = parser.parse_args()

    if args.detailed:
        info = get_port_info(args.host, args.port)
        print(f"Host: {info['host']}")
        print(f"Port: {info['port']}")
        print(f"Available: {'Yes' if info['is_available'] else 'No'}")
        print(f"Platform: {info['platform']}")

        if "process_info" in info:
            print("Process Information:")
            for line in info["process_info"]:
                print(f"  {line}")

    elif args.find_alternative:
        result = check_port_with_recommendation(args.host, args.port)
        print(result["message"])
        if result["alternative_port"]:
            print(f"권장 포트: {result['alternative_port']}")

    else:
        in_use = is_port_in_use(args.host, args.port)
        print(f"Port {args.port} on {args.host}: {'IN USE' if in_use else 'AVAILABLE'}")
        sys.exit(1 if in_use else 0)
