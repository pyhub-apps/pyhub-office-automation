# Release Notes - v9.2539.0

## 🎉 주요 개선사항

### COM 에러 메시지 사용자 친화적 개선 (Issue #41)

이제 Windows COM 에러가 발생할 때 기술적인 에러 코드 대신 이해하기 쉬운 메시지를 제공합니다.

#### 개선 전
```json
{
  "error": "(-2146827864, 'OLE error 0x800a01a8', None, None)"
}
```

#### 개선 후
```json
{
  "error": "Excel 객체에 접근할 수 없습니다. Excel이 실행 중이고 워크북이 열려있는지 확인하세요.",
  "suggestions": [
    "Excel 프로그램이 실행 중인지 확인",
    "워크북이 열려있는지 확인",
    "--visible 옵션을 사용하여 Excel 창 표시"
  ]
}
```

## 🐛 버그 수정

### Excel 명령어 버그 수정
- **workbook_open**: VersionNumber JSON 직렬화 오류 해결
- **pivot_configure**: 명령어 실행 문제 수정 (typer.run 적용)
- **Excel 명령어 전반**: JSON serialization 오류 해결

## ✨ 개선사항

### 에러 처리 시스템
- COM 에러 매핑 테이블 추가 (4개 주요 에러 코드)
  - `0x800A01A8`: Object Required
  - `0x800401A8`: Object Disconnected
  - `0x80010105`: RPC_E_SERVERFAULT
  - `0x800A03EC`: NAME_NOT_FOUND/INVALID_OPERATION
- 에러별 구체적인 해결 방법 제시
- 에러 원인 파악을 위한 상세 정보 제공

### 코드 품질
- 중앙화된 에러 처리 시스템 구축
- 일관된 에러 응답 구조 확립
- 확장 가능한 에러 매핑 시스템 구현

## 📋 변경된 파일

### 핵심 변경
- `utils.py`: COM 에러 매핑 시스템 추가
- `pivot_create.py`: COM 에러 전파 처리
- `chart_add.py`: COM 에러 전파 처리
- `chart_pivot_create.py`: COM 에러 전파 처리
- `workbook_open.py`: VersionNumber 직렬화 버그 수정
- `pivot_configure.py`: 실행 문제 수정

## 🔄 업그레이드 가이드

```bash
# 최신 버전 설치
pip install --upgrade pyhub-office-automation

# 또는 GitHub에서 직접 설치
pip install git+https://github.com/pyhub-apps/pyhub-office-automation.git@v9.2539.0
```

## ⚠️ 주의사항

- Windows 전용 기능들은 계속해서 Windows에서만 작동합니다
- COM 에러 개선은 Windows 환경에서만 적용됩니다
- macOS 사용자는 기존과 동일하게 작동합니다

## 🙏 감사의 말

이 릴리즈는 사용자분들의 피드백과 Issue 제보 덕분에 가능했습니다.
특히 COM 에러 관련 불편을 제보해주신 분들께 감사드립니다.

## 📚 문서

- [COM 에러 개선 제안서](docs/issues/com-error-improvement.md)
- [구현 보고서](docs/issues/com-error-implementation-report.md)

---

**Full Changelog**: https://github.com/pyhub-apps/pyhub-office-automation/compare/v8.2539.75...v9.2539.0