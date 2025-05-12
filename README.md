# 카드 지출 결의 및 개인차량 이용 내력 관리 시스템

ASP 기반의 카드 지출 결의 및 개인차량 이용 내력 관리 웹 애플리케이션입니다.

## 주요 기능

- 사용자 관리 (로그인, 회원가입)
- 카드 지출 결의 관리
  - 카드 사용 내역 등록/조회/수정
  - 카드 지출 결의서 출력
- 개인차량 이용 관리
  - 개인차량 이용 신청서 작성/조회/수정
  - 개인차량 이용 내역 출력
- 관리자 기능
  - 사용자 계정 관리
  - 카드 사용 내역 관리
  - 개인차량 이용 내역 관리
  - 보고서 생성 및 통계

## 기술 스택

- **언어**: ASP (Active Server Pages)
- **데이터베이스**: MSSQL
- **프론트엔드**: HTML, CSS, JavaScript
- **UI 라이브러리**: ShadCN UI (CSS 기반으로 구현)

## 설치 및 설정 방법

1. IIS(인터넷 정보 서비스)가 설치된 Windows 서버에 프로젝트 파일을 배포합니다.
2. `db.asp` 파일의 데이터베이스 연결 문자열을 실제 사용 환경에 맞게 수정합니다:
   ```
   strConnection = "Provider=SQLOLEDB;Data Source=YOUR_SERVER;Initial Catalog=YOUR_DATABASE;User ID=YOUR_USERNAME;Password=YOUR_PASSWORD;"
   ```
3. MSSQL 서버에 데이터베이스 스키마를 적용합니다.
   - `database_schema.md` 파일의 테이블 정의를 참고하여 데이터베이스 테이블을 생성합니다.

## 파일 구조

```
/
|- index.asp                 # 메인 페이지(로그인)
|- includes/                 # 공통 포함 파일
|  |- connection.asp         # 데이터베이스 연결
|  |- functions.asp          # 공통 함수
|  |- header.asp             # 헤더 템플릿
|  |- footer.asp             # 푸터 템플릿
|- pages/                    # 일반 사용자 페이지
|  |- dashboard.asp          # 대시보드
|  |- register.asp           # 회원가입
|  |- logout.asp             # 로그아웃
|  |- card_usage.asp         # 카드 사용 내역 목록
|  |- card_usage_add.asp     # 카드 사용 내역 등록
|  |- card_usage_edit.asp    # 카드 사용 내역 수정
|  |- vehicle_request.asp    # 개인차량 이용 신청 목록
|  |- vehicle_request_add.asp # 개인차량 이용 신청서 작성
|- admin/                    # 관리자 페이지
|  |- index.asp              # 관리자 대시보드
|  |- users.asp              # 사용자 관리
|  |- card_manage.asp        # 카드 관리
|  |- fuel_rate.asp          # 유류비 단가 관리
|- css/                      # CSS 파일
|  |- style.css              # 기본 스타일
|  |- shadcn.css             # ShadCN UI 스타일
|- js/                       # JavaScript 파일
|  |- common.js              # 공통 스크립트
|- images/                   # 이미지 파일
```

## 사용 방법

1. 웹 브라우저에서 시스템에 접속합니다.
2. 로그인 화면에서 계정이 없을 경우 '회원가입'을 클릭하여 계정을 생성합니다.
3. 로그인 후 대시보드에서 각 기능을 이용할 수 있습니다:
   - 카드 사용 내역 등록/조회
   - 개인차량 이용 신청서 작성/조회
   - 관리자 권한이 있는 경우 관리자 기능 사용

## 주의사항

- 이 프로젝트는 Windows IIS 환경에서 동작하도록 설계되었습니다.
- ASP와 MSSQL이 설치된 환경이 필요합니다.
- `db.asp` 파일은 Git에 포함되지 않으므로, 서버에 직접 배포해야 합니다.

## GitHub 저장소

- [https://github.com/qwer123toy/card_car_used](https://github.com/qwer123toy/card_car_used)

## 라이선스

이 프로젝트는 사내 사용 목적으로 개발되었습니다. 