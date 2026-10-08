# 병력분석기 v1.6 배포용

이 버전은 팀원 배포용입니다.

## 팀원이 보는 화면
1. 팀 비밀번호 입력
2. 청약/분석 기준일 선택
3. PDF 비밀번호 입력
4. PDF 첨부
5. `병력 분석하기` 클릭
6. PDF / Excel / TXT 저장

팀원은 OpenAI API 키나 모델을 입력하지 않습니다.

## 서버에만 저장하는 값
- `OPENAI_API_KEY`: OpenAI Platform에서 만든 API 키
- `APP_PASSWORD`: 팀원들이 접속할 공용 비밀번호
- `OPENAI_MODEL`: 기본값 `gpt-6-sol`

API 키를 app.py, GitHub, 카카오톡 등에 직접 적지 마세요.

## Render 배포 요약
1. GitHub 계정 생성 또는 로그인
2. 새 Private repository 생성
3. 이 폴더 안의 파일을 repository에 업로드
4. Render 가입/로그인 후 GitHub 연결
5. New > Blueprint 또는 Web Service 선택
6. 저장소 선택
7. 환경변수에 `OPENAI_API_KEY`, `APP_PASSWORD` 입력
8. Deploy
9. 생성된 `https://...onrender.com` 주소를 팀원에게 공유

`render.yaml`이 포함되어 있어 Blueprint로 배포하면 Docker 설정을 자동으로 읽습니다.

## 무료 Render 주의
무료 Web Service는 일정 시간 사용이 없으면 잠들 수 있어 다음 첫 접속 때 로딩이 길 수 있습니다. 팀에서 상시 사용하려면 유료 인스턴스를 고려하세요.

## 보안 주의
이 앱의 공용 비밀번호는 간단한 1차 접근통제입니다. 고객 건강정보는 민감정보이므로 회사 정책 및 외부 AI 사용 허용 여부를 확인하세요. 더 강한 보안이 필요하면 회사 계정 로그인/SSO 또는 사내 서버 배포가 필요합니다.
