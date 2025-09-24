"""
HWP HTML Export 명령어
HWP 문서를 HTML 형식으로 변환하는 기능
"""

import json
import os
import platform
import tempfile
from pathlib import Path
from typing import Optional

import typer

from pyhub_office_automation.version import get_version

from .utils import (
    ExecutionTimer,
    check_hwp_installed,
    clean_html_content,
    cleanup_temp_file,
    create_error_response,
    create_success_response,
    create_temp_html_file,
    format_output,
    get_file_size,
    normalize_path,
    validate_file_path,
)


def hwp_export(
    file_path: str = typer.Option(..., "--file-path", help="변환할 HWP 파일의 절대 경로"),
    format_type: str = typer.Option("html", "--format", help="출력 형식 (현재 html만 지원)"),
    output_file: Optional[str] = typer.Option(None, "--output-file", help="HTML 저장 경로 (선택, 미지정시 표준출력)"),
    encoding: str = typer.Option("utf-8", "--encoding", help="출력 인코딩 (기본값: utf-8)"),
    include_css: bool = typer.Option(False, "--include-css/--no-include-css", help="CSS 스타일 포함 여부 (기본값: False, 모든 CSS 제거)"),
    include_images: bool = typer.Option(False, "--include-images/--no-include-images", help="이미지 포함 여부 (기본값: False, Base64 인코딩으로 포함)"),
    temp_cleanup: bool = typer.Option(True, "--temp-cleanup/--no-temp-cleanup", help="임시 파일 자동 정리 (기본값: True)"),
    output_format: str = typer.Option("json", "--output-format", help="응답 출력 형식 (json)"),
):
    """
    HWP 문서를 HTML 형식으로 변환합니다.

    pyhwpx 라이브러리를 사용하여 HWP 문서를 HTML로 변환하고,
    AI 에이전트가 파싱하기 쉬운 구조화된 JSON 형식으로 결과를 반환합니다.

    \b
    주요 기능:
      • HWP → HTML 변환
      • CSS 정리 및 최적화
      • UTF-8 인코딩 지원
      • 임시 파일 자동 정리

    \b
    사용 예제:
      oa hwp export --file-path "문서.hwp" --format html
      oa hwp export --file-path "문서.hwp" --format html --output-file "문서.html"
      oa hwp export --file-path "문서.hwp" --format html --include-css --output-file "스타일포함.html"

    \b
    요구사항:
      • Windows 운영체제
      • HWP(한글) 2010 이상 설치
      • pyhwpx 라이브러리

    \b
    참고사항:
      • HWP 프로그램이 잠깐(0.1-0.5초) 화면에 나타날 수 있습니다
      • 이는 Windows COM 아키텍처의 특성상 불가피한 현상입니다
    """

    try:
        # 실행 시간 측정 시작
        with ExecutionTimer() as timer:
            # 1. 기본 검증
            if platform.system() != "Windows":
                error_response = create_error_response(
                    "HWP 변환 기능은 Windows에서만 지원됩니다",
                    "PlatformError"
                )
                typer.echo(format_output(error_response, output_format))
                raise typer.Exit(1)

            # 지원 형식 검증
            if format_type.lower() != "html":
                error_response = create_error_response(
                    f"지원하지 않는 형식입니다: {format_type} (현재 html만 지원)",
                    "UnsupportedFormatError"
                )
                typer.echo(format_output(error_response, output_format))
                raise typer.Exit(1)

            # HWP 설치 확인
            if not check_hwp_installed():
                error_response = create_error_response(
                    "HWP 프로그램이 설치되어 있지 않습니다. 한글 프로그램을 설치한 후 다시 시도해 주세요.",
                    "HWPNotInstalledError"
                )
                typer.echo(format_output(error_response, output_format))
                raise typer.Exit(1)

            # 2. 파일 경로 검증 및 정규화
            validated_file_path = validate_file_path(file_path)
            original_size = get_file_size(validated_file_path)

            # 3. 출력 파일 경로 처리
            if output_file:
                output_file = normalize_path(output_file)
                # 출력 디렉토리가 없으면 생성
                output_dir = Path(output_file).parent
                output_dir.mkdir(parents=True, exist_ok=True)

            # 4. HWP → HTML 변환 실행
            html_content = _convert_hwp_to_html(
                validated_file_path,
                temp_cleanup=temp_cleanup
            )

            # 5. HTML 후처리 (charset 수정, CSS 처리, 이미지 처리)
            html_content = _process_html_content(
                html_content,
                include_css=include_css,
                include_images=include_images
            )

            # 6. 출력 처리
            output_size = len(html_content.encode(encoding))

            if output_file:
                # 파일로 저장
                try:
                    with open(output_file, 'w', encoding=encoding) as f:
                        f.write(html_content)
                except Exception as e:
                    error_response = create_error_response(
                        f"파일 저장 실패: {str(e)}",
                        "FileWriteError"
                    )
                    typer.echo(format_output(error_response, output_format))
                    raise typer.Exit(1)
            else:
                # 표준 출력에 HTML 직접 출력 (JSON 래핑 없이)
                typer.echo(html_content)
                return

            # 7. 성공 응답 생성 (파일 저장 시에만)
            processing_stats = {
                "original_file_size_bytes": original_size,
                "output_size_bytes": output_size,
                "processing_time_ms": timer.duration_ms,
                "css_included": include_css,
                "images_included": include_images,
                "temp_files_cleaned": temp_cleanup
            }

            metadata = {
                "encoding": encoding,
                "source_format": "HWP",
                "target_format": "HTML"
            }

            data = {
                "input_file": validated_file_path,
                "output_file": output_file,
                "format": format_type.lower()
            }

            success_response = create_success_response(
                data=data,
                processing_stats=processing_stats,
                metadata=metadata
            )

            typer.echo(format_output(success_response, output_format))

    except typer.Exit:
        # typer.Exit은 그대로 전파
        raise
    except Exception as e:
        # 예상치 못한 에러 처리
        error_response = create_error_response(
            f"변환 중 오류가 발생했습니다: {str(e)}",
            "UnexpectedError"
        )
        typer.echo(format_output(error_response, output_format))
        raise typer.Exit(1)


def _convert_hwp_to_html(
    file_path: str,
    temp_cleanup: bool = True
) -> str:
    """
    HWP 파일을 HTML로 변환하는 내부 함수

    Args:
        file_path: HWP 파일 경로
        temp_cleanup: 임시 파일 정리 여부

    Returns:
        변환된 HTML 내용

    Raises:
        Exception: 변환 실패 시
    """
    hwp = None
    temp_html_path = None

    try:
        # pyhwpx import with warning suppression (COM 캐시 재구축 경고 방지)
        import warnings
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            import pyhwpx

        # HWP 객체 생성 (Windows COM 특성상 잠깐 화면에 나타남)
        hwp = pyhwpx.Hwp(visible=False)

        # 파일 열기
        if not hwp.open(file_path):
            raise Exception(f"HWP 파일을 열 수 없습니다: {file_path}")

        # 임시 HTML 파일 경로 생성
        temp_html_path = create_temp_html_file()

        # HTML로 변환하여 저장
        if not hwp.save_as(temp_html_path, format="HTML"):
            raise Exception("HTML 변환 실패")

        # HTML 파일 내용 읽기 (다양한 인코딩 시도)
        html_content = None
        for encoding_try in ['utf-8', 'cp949', 'euc-kr', 'latin1']:
            try:
                with open(temp_html_path, 'r', encoding=encoding_try) as f:
                    html_content = f.read()
                break
            except UnicodeDecodeError:
                continue

        if html_content is None:
            raise Exception("HTML 파일을 읽을 수 없습니다. 인코딩 문제가 발생했습니다.")

        return html_content

    except ImportError:
        raise Exception("pyhwpx 라이브러리가 설치되어 있지 않습니다. pip install pyhwpx로 설치해 주세요.")

    except Exception as e:
        raise Exception(f"HWP 변환 처리 중 오류: {str(e)}")

    finally:
        # 리소스 정리
        if hwp:
            try:
                hwp.quit()
            except Exception:
                pass  # 정리 실패해도 무시

        # 임시 파일 정리
        if temp_cleanup and temp_html_path:
            cleanup_temp_file(temp_html_path)


def _process_html_content(html_content: str, include_css: bool = False, include_images: bool = False) -> str:
    """
    HTML 내용 후처리 (charset 수정, CSS 처리, 이미지 처리)

    Args:
        html_content: 원본 HTML 내용
        include_css: CSS 포함 여부 (False시 모든 CSS 제거)
        include_images: 이미지 포함 여부 (True시 Base64 인코딩, False시 제거)

    Returns:
        후처리된 HTML 내용
    """
    import re

    # 1. charset을 UTF-8로 변경
    html_content = re.sub(
        r'<meta[^>]*charset\s*=\s*["\']?[^"\'>\s]*["\']?[^>]*>',
        '<meta http-equiv="Content-Type" content="text/html; charset=utf-8">',
        html_content,
        flags=re.IGNORECASE
    )

    # 2. CSS 처리
    if not include_css:
        # 모든 CSS 제거 (더 강력한 정리)
        html_content = _remove_all_css(html_content)
    else:
        # CSS 포함 시에는 기존의 clean_html_content 함수 사용 (불필요한 것만 제거)
        html_content = clean_html_content(html_content)

    # 3. 이미지 처리
    if include_images:
        # Base64로 이미지 인코딩
        html_content = _encode_images_to_base64(html_content)
    else:
        # 모든 img 태그 제거
        html_content = _remove_all_images(html_content)

    return html_content


def _remove_all_css(html_content: str) -> str:
    """
    HTML에서 모든 CSS 관련 요소 제거 및 손상된 태그 복구

    Args:
        html_content: 원본 HTML 내용

    Returns:
        CSS가 모두 제거되고 태그가 복구된 HTML 내용
    """
    import re

    # 1. <style> 태그 및 내용 모두 제거
    html_content = re.sub(r'<style[^>]*>.*?</style>', '', html_content, flags=re.DOTALL | re.IGNORECASE)

    # 2. CSS 관련 스크립트 제거
    html_content = re.sub(r'<script[^>]*>.*?</script>', '', html_content, flags=re.DOTALL | re.IGNORECASE)

    # 3. 손상된 span 태그들을 우선적으로 복구 (HWP 특수 케이스)
    # HWP에서 font-family가 CSS와 함께 태그명에 붙어나오는 경우 처리
    html_content = re.sub(
        r'<span([^>]*?)([^"\s>]+)"([^>]*?)>',
        r'<span\1\3">',
        html_content,
        flags=re.IGNORECASE
    )

    # 4. 완전히 깨진 태그들 복구 (태그명에 스타일이 붙은 경우)
    def fix_broken_span_tags(match):
        full_match = match.group(0)
        # span 태그명에 CSS가 붙어있는 경우를 찾아서 수정
        if 'span' in full_match.lower():
            # 간단히 span 태그로 변환
            return '<span>'
        return full_match

    html_content = re.sub(r'<span[^>]*[^"\s>]+["\'][^>]*>', fix_broken_span_tags, html_content, flags=re.IGNORECASE)

    # 5. 남은 style과 class 속성 정리
    html_content = re.sub(r'\s*style\s*=\s*["\'][^"\']*["\']', '', html_content, flags=re.IGNORECASE)
    html_content = re.sub(r'\s*class\s*=\s*["\'][^"\']*["\']', '', html_content, flags=re.IGNORECASE)

    # 6. 시각적 레이아웃 속성 제거 (구조적 속성은 보존)
    html_content = _remove_layout_attributes(html_content)

    # 7. 빈 속성이나 잘못된 속성 정리
    html_content = re.sub(r'<span\s*>', '<span>', html_content, flags=re.IGNORECASE)

    # 8. 불필요한 공백 정리
    html_content = re.sub(r'\n\s*\n', '\n', html_content)
    html_content = re.sub(r'>\s+<', '><', html_content)

    return html_content.strip()


def _remove_layout_attributes(html_content: str) -> str:
    """
    시각적 레이아웃 속성만 제거 (구조적 속성은 보존)

    Args:
        html_content: 원본 HTML 내용

    Returns:
        시각적 레이아웃 속성이 제거된 HTML 내용

    Note:
        rowspan, colspan은 테이블 구조의 핵심이므로 절대 제거하지 않음
    """
    import re

    # 시각적 테이블 속성만 제거 (구조적 속성 보존)
    html_content = re.sub(r'\s*cellspacing\s*=\s*["\'][^"\']*["\']', '', html_content, flags=re.IGNORECASE)
    html_content = re.sub(r'\s*cellpadding\s*=\s*["\'][^"\']*["\']', '', html_content, flags=re.IGNORECASE)
    html_content = re.sub(r'\s*valign\s*=\s*["\'][^"\']*["\']', '', html_content, flags=re.IGNORECASE)

    # rowspan, colspan은 절대 제거하지 않음 - 테이블 데이터 구조의 핵심

    # 레이아웃용 공백 span 제거
    html_content = re.sub(r'<span[^>]*>&nbsp;</span>', ' ', html_content, flags=re.IGNORECASE)
    html_content = re.sub(r'<span[^>]*>\s*</span>', '', html_content, flags=re.IGNORECASE)

    # 연속된 공백을 하나로 정리
    html_content = re.sub(r'\s+', ' ', html_content)

    return html_content


def _remove_all_images(html_content: str) -> str:
    """
    HTML에서 모든 이미지 태그 제거

    Args:
        html_content: 원본 HTML 내용

    Returns:
        이미지가 모두 제거된 HTML 내용
    """
    import re

    # 모든 img 태그 제거
    html_content = re.sub(r'<img[^>]*>', '', html_content, flags=re.IGNORECASE)

    # 불필요한 공백 정리
    html_content = re.sub(r'\n\s*\n', '\n', html_content)

    return html_content.strip()


def _encode_images_to_base64(html_content: str) -> str:
    """
    HTML의 이미지를 Base64로 인코딩

    Args:
        html_content: 원본 HTML 내용

    Returns:
        이미지가 Base64로 인코딩된 HTML 내용
    """
    import re
    import base64
    import os
    from urllib.parse import unquote

    def encode_image(match):
        img_tag = match.group(0)
        src_match = re.search(r'src\s*=\s*["\']([^"\']+)["\']', img_tag, re.IGNORECASE)

        if not src_match:
            return img_tag

        image_path = src_match.group(1)

        # file:// 프로토콜 제거
        if image_path.startswith('file:///'):
            image_path = image_path[8:]  # file:/// 제거
        elif image_path.startswith('file://'):
            image_path = image_path[7:]  # file:// 제거

        # URL 디코딩
        image_path = unquote(image_path)

        # 파일이 존재하는지 확인
        if not os.path.exists(image_path):
            # 파일이 없으면 img 태그 제거
            return ''

        try:
            # 파일 확장자로 MIME 타입 결정
            _, ext = os.path.splitext(image_path.lower())
            mime_types = {
                '.png': 'image/png',
                '.jpg': 'image/jpeg',
                '.jpeg': 'image/jpeg',
                '.gif': 'image/gif',
                '.bmp': 'image/bmp',
                '.webp': 'image/webp'
            }
            mime_type = mime_types.get(ext, 'image/png')

            # 파일을 Base64로 인코딩
            with open(image_path, 'rb') as f:
                image_data = f.read()
                base64_data = base64.b64encode(image_data).decode('utf-8')

            # src를 data URI로 교체
            new_src = f'data:{mime_type};base64,{base64_data}'
            new_img_tag = re.sub(
                r'src\s*=\s*["\'][^"\']+["\']',
                f'src="{new_src}"',
                img_tag,
                flags=re.IGNORECASE
            )

            return new_img_tag

        except Exception:
            # 인코딩 실패 시 img 태그 제거
            return ''

    # 모든 img 태그에 대해 Base64 인코딩 시도
    html_content = re.sub(r'<img[^>]*>', encode_image, html_content, flags=re.IGNORECASE)

    return html_content


# CLI 명령어로 사용될 때 직접 실행
if __name__ == "__main__":
    typer.run(hwp_export)