# COM 에러 메시지를 사용자 친화적으로 개선

## 문제 설명
Excel 자동화 명령어 실행 시 발생하는 COM 에러가 사용자에게 직관적이지 않은 메시지를 표시합니다.

### 현재 상황
```json
{
  "success": false,
  "error_type": "com_error",
  "error": "(-2146827864, 'OLE error 0x800a01a8', None, None)",
  "command": "pivot-create",
  "version": "8.2539.75"
}
```

### 기대하는 개선
```json
{
  "success": false,
  "error_type": "com_error",
  "error": "Excel 객체에 접근할 수 없습니다. Excel이 실행 중이고 워크북이 열려있는지 확인하세요.",
  "error_details": {
    "code": "0x800a01a8",
    "meaning": "Object Required",
    "possible_causes": [
      "Excel이 실행되지 않음",
      "워크북이 닫혀있음",
      "Excel 객체가 해제됨"
    ]
  },
  "suggestions": [
    "Excel 프로그램이 실행 중인지 확인",
    "워크북이 열려있는지 확인",
    "--visible 옵션을 사용하여 Excel 창 표시"
  ],
  "command": "pivot-create",
  "version": "8.2539.75"
}
```

## 개선 제안

### 1. COM 에러 매핑 테이블 추가 (utils.py) ✅
```python
COM_ERROR_MESSAGES = {
    0x800A01A8: {
        "message": "Excel 객체에 접근할 수 없습니다",
        "meaning": "Object Required",
        "causes": [
            "Excel이 실행되지 않음",
            "워크북이 닫혀있음",
            "Excel 객체가 해제됨"
        ],
        "suggestions": [
            "Excel 프로그램이 실행 중인지 확인",
            "워크북이 열려있는지 확인",
            "--visible 옵션을 사용하여 Excel 창 표시"
        ]
    },
    0x800401A8: {
        "message": "Excel COM 객체 연결이 끊어졌습니다",
        "meaning": "Object is disconnected from clients",
        "causes": ["Excel 프로세스가 종료됨", "COM 객체 수명 주기 문제"],
        "suggestions": ["Excel을 다시 시작", "명령을 다시 실행"]
    }
}
```

### 2. create_error_response 함수 개선 ✅
```python
def create_error_response(error: Exception, command: str):
    error_type = type(error).__name__

    # COM 에러 특별 처리
    if error_type == "com_error":
        error_code = extract_com_error_code(error)
        if error_code in COM_ERROR_MESSAGES:
            com_info = COM_ERROR_MESSAGES[error_code]
            return {
                "success": False,
                "error_type": error_type,
                "error": com_info["message"],
                "error_details": {
                    "code": hex(error_code),
                    "meaning": com_info["meaning"],
                    "possible_causes": com_info["causes"]
                },
                "suggestions": com_info["suggestions"],
                "command": command,
                "version": get_version()
            }

    # 기존 처리 로직...
```

### 3. 각 명령어의 예외 처리 개선 ✅
```python
except Exception as e:
    # COM 에러를 먼저 체크
    if "com_error" in str(type(e).__name__).lower():
        raise  # create_error_response에서 처리하도록 전달
    else:
        # 기존 RuntimeError 처리
        raise RuntimeError(f"피벗테이블 생성 실패: {str(e)}")
```

## 구현된 개선 사항

### 추가된 COM 에러 코드
- **0x800A01A8**: Object Required - Excel 객체 접근 불가
- **0x800401A8**: Object Disconnected - COM 연결 끊김
- **0x80010105**: RPC_E_SERVERFAULT - Excel 서버 종료
- **0x800A03EC**: NAME_NOT_FOUND/INVALID_OPERATION - 잘못된 작업

## 영향 범위
- ✅ `pyhub_office_automation/excel/utils.py` - COM 에러 매핑 및 처리 함수 추가
- ✅ `pyhub_office_automation/excel/pivot_create.py` - COM 에러 전파 처리
- 📋 추후 모든 Excel 명령어 파일에 동일한 패턴 적용 필요

## 테스트 시나리오
1. Excel이 실행되지 않은 상태에서 명령 실행
2. 워크북이 닫힌 상태에서 명령 실행
3. Excel 프로세스 강제 종료 후 명령 실행
4. 잘못된 시트명이나 범위로 명령 실행

## 테스트 결과 예시
```bash
# Excel이 실행되지 않은 상태
oa excel pivot-create --source-range "A1:D100"

# 개선 전
{
  "error": "(-2146827864, 'OLE error 0x800a01a8', None, None)"
}

# 개선 후
{
  "error": "Excel 객체에 접근할 수 없습니다. Excel이 실행 중이고 워크북이 열려있는지 확인하세요.",
  "suggestions": [
    "Excel 프로그램이 실행 중인지 확인",
    "워크북이 열려있는지 확인",
    "--visible 옵션을 사용하여 Excel 창 표시"
  ]
}
```

## 우선순위
높음 - 사용자 경험에 직접적인 영향

## 라벨
- enhancement
- user-experience
- error-handling

## 관련 코드 변경
- [x] utils.py에 COM_ERROR_MESSAGES 딕셔너리 추가
- [x] extract_com_error_code() 함수 추가
- [x] create_error_response() 함수 개선
- [x] pivot_create.py의 예외 처리 개선
- [ ] 다른 Excel 명령어 파일들에 동일 패턴 적용 (향후 작업)