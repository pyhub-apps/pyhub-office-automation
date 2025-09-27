# Email 자동화 사용 매뉴얼

## 개요

pyhub-office-automation의 email 기능은 AI 기반 이메일 생성 및 발송, 다중 계정 관리, Windows Credential Manager를 통한 안전한 자격증명 관리를 제공합니다.

## 주요 기능

- ✅ **AI 기반 이메일 생성**: 프롬프트를 통한 자동 이메일 작성
- ✅ **다중 계정 관리**: Gmail, Outlook, Naver 등 여러 계정 동시 사용
- ✅ **안전한 자격증명 저장**: Windows Credential Manager 연동
- ✅ **앱 비밀번호 지원**: OAuth2 없이 간편한 인증
- ✅ **크로스플랫폼 백엔드**: Outlook COM (Windows) + SMTP (범용)

## 명령어 개요

```bash
oa email <command> [options]
```

### 사용 가능한 명령어

| 명령어 | 설명 | 예시 |
|--------|------|------|
| `config` | 이메일 계정 설정 | `oa email config --provider gmail` |
| `accounts` | 계정 목록 조회 | `oa email accounts` |
| `delete` | 계정 삭제 | `oa email delete work` |
| `send` | 이메일 발송 | `oa email send --account work --to user@example.com` |
| `list` | 명령어 목록 출력 | `oa email list` |

## 계정 설정 (config)

### 기본 사용법

```bash
# 대화식 설정 (권장)
oa email config

# 명령행 옵션으로 설정
oa email config --provider gmail --username user@gmail.com --account-name work
```

### 지원 제공자

#### 1. Gmail
```bash
oa email config --provider gmail --username your@gmail.com
```
- **SMTP 서버**: smtp.gmail.com:587
- **앱 비밀번호 필요**: [설정 가이드](https://support.google.com/accounts/answer/185833)
- **2단계 인증 활성화 후 앱 비밀번호 생성**

#### 2. Outlook.com
```bash
oa email config --provider outlook --username your@outlook.com
```
- **SMTP 서버**: smtp-mail.outlook.com:587
- **앱 비밀번호 필요**: [설정 가이드](https://support.microsoft.com/account-billing)

#### 3. Naver Mail
```bash
oa email config --provider naver --username your@naver.com
```
- **SMTP 서버**: smtp.naver.com:587
- **앱 비밀번호 필요**: [설정 가이드](https://help.naver.com/alias/mail/mail_26.naver)

#### 4. Custom SMTP
```bash
oa email config --provider custom --username your@company.com --server smtp.company.com --port 587
```

### 설정 옵션

| 옵션 | 설명 | 기본값 |
|------|------|--------|
| `--provider` | 이메일 제공자 (gmail, outlook, naver, custom) | None (대화식 선택) |
| `--username` | 이메일 주소 | None (프롬프트) |
| `--account-name` | 계정 별칭 | default |
| `--server` | SMTP 서버 (custom용) | 제공자별 기본값 |
| `--port` | SMTP 포트 | 587 |
| `--no-tls` | TLS 사용 안함 | False (TLS 사용) |

### 설정 예시

```bash
# 회사용 Gmail 계정
oa email config --provider gmail --username work@gmail.com --account-name work

# 개인용 Outlook 계정
oa email config --provider outlook --username personal@outlook.com --account-name personal

# 사내 메일서버
oa email config --provider custom --username user@company.com --server mail.company.com --account-name company
```

## 계정 관리 (accounts, delete)

### 계정 목록 조회

```bash
# 테이블 형태로 출력 (기본값)
oa email accounts

# JSON 형태로 출력
oa email accounts --format json

# 상세 정보 포함
oa email accounts --verbose
```

### 출력 예시

```
📧 등록된 이메일 계정
┌─────────────┬──────────────────────┬──────────┬─────────────────────┬──────┬─────┬────────┐
│ 계정명      │ 이메일 주소           │ 제공자   │ 서버                │ 포트 │ TLS │ 상태   │
├─────────────┼──────────────────────┼──────────┼─────────────────────┼──────┼─────┼────────┤
│ work        │ work@gmail.com       │ GMAIL    │ smtp.gmail.com      │ 587  │ ✅  │ configured │
│ personal    │ personal@outlook.com │ OUTLOOK  │ smtp-mail.outlook.com│ 587  │ ✅  │ configured │
└─────────────┴──────────────────────┴──────────┴─────────────────────┴──────┴─────┴────────┘

📊 총 2개 계정이 등록되어 있습니다.
```

### 계정 삭제

```bash
# 확인 프로세스와 함께 삭제
oa email delete work

# 확인 없이 즉시 삭제
oa email delete work --confirm
```

## 이메일 발송 (send)

### 기본 사용법

```bash
# AI 기반 이메일 생성 및 발송
oa email send --account work --to recipient@example.com --prompt "프로젝트 진행 상황 보고"

# 수동으로 제목과 본문 지정
oa email send --account work --to recipient@example.com --subject "안녕하세요" --body "테스트 메일입니다"
```

### 주요 옵션

#### 필수 옵션
- `--to`: 받는 사람 이메일 주소

#### 계정 선택
- `--account`: 사용할 계정명 (미지정 시 기본 계정 또는 환경변수 사용)

#### AI 생성 관련
- `--prompt`: AI 이메일 생성 프롬프트
- `--prompt-file`: 프롬프트 파일 경로
- `--ai-provider`: AI 제공자 (auto, claude, gemini, openai 등)
- `--language`: 언어 설정 (ko, en)
- `--tone`: 어조 설정 (formal, casual, business)

#### 수동 작성
- `--subject`: 이메일 제목
- `--body`: 이메일 본문
- `--body-file`: 본문 파일 경로

#### 추가 옵션
- `--from`: 보내는 사람 주소 (기본값: 계정 이메일)
- `--cc`: 참조 (쉼표로 구분)
- `--bcc`: 숨은 참조 (쉼표로 구분)
- `--attachments`: 첨부 파일 (쉼표로 구분)
- `--body-type`: 본문 형식 (text, html)
- `--backend`: 이메일 백엔드 (auto, outlook, smtp)

### 발송 예시

#### 1. AI 기반 이메일 생성

```bash
# 한국어 비즈니스 이메일
oa email send \
  --account work \
  --to client@company.com \
  --prompt "다음 주 화요일 오후 2시 회의실 A에서 프로젝트 킥오프 미팅 안내" \
  --language ko \
  --tone business

# 영어 공식 이메일
oa email send \
  --account work \
  --to partner@international.com \
  --prompt "Request for proposal submission deadline extension" \
  --language en \
  --tone formal
```

#### 2. 수동 이메일 작성

```bash
# 간단한 알림 메일
oa email send \
  --account personal \
  --to friend@example.com \
  --subject "저녁 약속 변경" \
  --body "안녕하세요. 저녁 약속을 7시로 변경하고 싶습니다."

# 첨부파일이 있는 메일
oa email send \
  --account work \
  --to team@company.com \
  --subject "월간 보고서" \
  --body-file report_message.txt \
  --attachments "report.pdf,chart.xlsx"
```

#### 3. 참조/숨은참조 포함

```bash
oa email send \
  --account work \
  --to primary@example.com \
  --cc "manager@company.com,colleague@company.com" \
  --bcc "archive@company.com" \
  --prompt "프로젝트 완료 보고"
```

### 백엔드 선택

#### Outlook COM (Windows 전용)
```bash
oa email send --backend outlook --to user@example.com --prompt "테스트 메일"
```
- **장점**: Outlook 앱을 통한 발송, 별도 설정 불필요
- **단점**: Windows 전용, Outlook 설치 필요

#### SMTP (범용)
```bash
oa email send --backend smtp --account work --to user@example.com --prompt "테스트 메일"
```
- **장점**: 크로스플랫폼, 모든 SMTP 서버 지원
- **단점**: 계정 설정 필요

#### Auto (자동 선택)
```bash
oa email send --backend auto --to user@example.com --prompt "테스트 메일"
```
- Windows에서 Outlook 사용 가능하면 Outlook COM, 아니면 SMTP

## 보안 및 자격증명 관리

### Windows Credential Manager 저장 구조

각 계정은 다음과 같이 저장됩니다:

```
서비스명: oa-email-{account_name}
- username: 이메일 주소
- password: 앱 비밀번호
- server: SMTP 서버 주소
- port: SMTP 포트
- use_tls: TLS 사용 여부
```

### 보안 특징

1. **암호화된 저장**: Windows Credential Manager의 기본 암호화
2. **사용자별 격리**: Windows 사용자 계정별로 분리된 접근
3. **앱 비밀번호**: OAuth2 복잡성 없이 안전한 인증
4. **숨김 입력**: 비밀번호 입력 시 화면에 표시되지 않음

### 크로스플랫폼 호환성

- **Windows**: Windows Credential Manager 사용
- **macOS/Linux**: keyring 라이브러리의 기본 백엔드 사용
- **Docker**: 환경변수 폴백 지원

## 문제 해결

### 계정 설정 관련

#### Q: 앱 비밀번호가 뭔가요?
A: 2단계 인증이 활성화된 계정에서 일반 비밀번호 대신 사용하는 별도의 비밀번호입니다.

**Gmail 앱 비밀번호 생성:**
1. Google 계정 관리 → 보안
2. 2단계 인증 활성화
3. 앱 비밀번호 생성
4. "메일" 선택 후 16자리 비밀번호 복사

#### Q: 계정 설정이 저장되지 않아요
A: Windows Credential Manager 접근 권한을 확인하세요:
```bash
# keyring 테스트
python -c "import keyring; keyring.set_password('test', 'user', 'pass'); print('OK')"
```

#### Q: 계정 목록이 비어있어요
A: 다음을 확인하세요:
1. 계정이 실제로 설정되었는지: `oa email config`
2. Windows Credential Manager에서 `oa-email-*` 항목 확인
3. 계정명 오타 여부

### 이메일 발송 관련

#### Q: SMTP 연결 오류가 발생해요
A: 다음을 확인하세요:
1. 앱 비밀번호 정확성
2. SMTP 서버 주소 및 포트
3. 방화벽/보안 소프트웨어 설정
4. 2단계 인증 활성화 여부

```bash
# 계정 설정 재확인
oa email accounts --verbose

# 수동 SMTP 테스트
oa email send --backend smtp --smtp-server smtp.gmail.com --smtp-port 587 \
  --smtp-user your@gmail.com --smtp-password your-app-password \
  --to test@example.com --subject "테스트" --body "연결 테스트"
```

#### Q: Outlook COM 백엔드를 사용할 수 없어요
A: 다음을 확인하세요:
1. Windows 운영체제 여부
2. Microsoft Outlook 설치 여부
3. pywin32 라이브러리 설치: `pip install pywin32`

### AI 생성 관련

#### Q: AI 이메일 생성이 안 돼요
A: AI 제공자 설정을 확인하세요:
```bash
# 사용 가능한 AI 제공자 확인
oa email send --help

# 특정 AI 제공자 사용
oa email send --ai-provider claude --api-key your-api-key \
  --to user@example.com --prompt "테스트 메일"
```

## 고급 사용법

### 배치 처리

여러 수신자에게 동일한 이메일 발송:
```bash
#!/bin/bash
recipients="user1@example.com user2@example.com user3@example.com"

for recipient in $recipients; do
  oa email send --account work --to "$recipient" \
    --prompt "월간 뉴스레터 발송" --confirm false
done
```

### 템플릿 활용

이메일 템플릿 파일 활용:
```bash
# template.txt 파일 작성
echo "안녕하세요. 이번 달 실적 보고서를 첨부해 드립니다." > template.txt

# 템플릿 사용
oa email send --account work --to manager@company.com \
  --subject "월간 실적 보고" --body-file template.txt \
  --attachments "report.pdf"
```

### JSON 출력 활용

스크립트에서 결과 처리:
```bash
#!/bin/bash
result=$(oa email send --account work --to user@example.com \
  --prompt "테스트" --format json --no-confirm)

status=$(echo $result | jq -r '.status')
if [ "$status" = "sent" ]; then
  echo "이메일 발송 성공"
else
  echo "이메일 발송 실패: $(echo $result | jq -r '.error')"
fi
```

## 환경변수 설정

계정 설정 없이 환경변수로 SMTP 사용:

```bash
# Windows
set SMTP_SERVER=smtp.gmail.com
set SMTP_PORT=587
set SMTP_USERNAME=your@gmail.com
set SMTP_PASSWORD=your-app-password
set SMTP_USE_TLS=true

# Linux/macOS
export SMTP_SERVER=smtp.gmail.com
export SMTP_PORT=587
export SMTP_USERNAME=your@gmail.com
export SMTP_PASSWORD=your-app-password
export SMTP_USE_TLS=true

# 환경변수로 발송
oa email send --backend smtp --to user@example.com --prompt "테스트"
```

## API 참조

### JSON 출력 형식

모든 명령어는 `--format json` 옵션으로 구조화된 출력을 제공합니다.

#### 계정 목록 (accounts)
```json
{
  "status": "success",
  "version": "1.0.0",
  "accounts": [
    {
      "account_name": "work",
      "username": "work@gmail.com",
      "provider": "gmail",
      "server": "smtp.gmail.com",
      "port": 587,
      "use_tls": true,
      "status": "configured"
    }
  ],
  "total_count": 1
}
```

#### 이메일 발송 (send)
```json
{
  "status": "sent",
  "version": "1.0.0",
  "result": {
    "backend": "smtp",
    "message_id": "...",
    "to": ["user@example.com"],
    "cc": [],
    "bcc": [],
    "smtp_server": "smtp.gmail.com"
  },
  "execution_time": 2.45,
  "ai_provider": "claude",
  "generated_content": {
    "subject": "프로젝트 진행 상황 보고",
    "body": "안녕하세요. 현재 프로젝트 진행 상황을 보고드립니다..."
  }
}
```

## 버전 정보

- **버전**: 10.2539.33+
- **지원 플랫폼**: Windows 10/11 (주), macOS/Linux (제한적)
- **Python 요구사항**: Python 3.13+
- **주요 의존성**: keyring, typer, rich, xlwings

## 관련 문서

- [설치 가이드](../README.md#설치)
- [AI 제공자 설정](./ai-providers.md)
- [보안 정책](./security.md)
- [GitHub 이슈 #67](https://github.com/pyhub-apps/pyhub-office-automation/issues/67)