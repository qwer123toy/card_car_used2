
<div align="center">
  <img src="/images/img4.png" alt="card car system banner">
</div>

<br>

<h1 align="center">
카드 지출 결의 및 차량 이용 내역 관리 시스템
</h1>
<p align="center">기업 내 카드 사용과 개인 차량 이용을 통합 관리하는 ASP 기반 웹 애플리케이션</p>
<p align="center">사용자와 관리자가 효율적으로 결재 내역과 신청서를 관리할 수 있는 업무 특화 시스템</p>

---

<br>

## 📌 Contents

<p align="left">목차</p>
<p align="left">
  <a href="#what-is">What is?</a>  <br>
  <a href="#key-features">Key Features</a> <br>
  <a href="#development-setup">Development Setup</a> <br>
  <a href="#repository-structure">Repository Structure</a> <br>
  <a href="#authors">Authors</a>
</p>

<br>

## What is

### 1. 개요 및 목적

 - 본 시스템은 기업 내 법인카드 사용 결의서 작성 및 개인 차량 이용 신청 과정을 전산화하여 문서화, 결재, 출력, 통계 기능을 통합한 웹 기반 업무 플랫폼
 - 기존 수기 결재나 개별 파일 관리 방식에서 벗어나 중복 업무를 줄이고 결재 프로세스를 간소화함으로써 관리 효율성을 극대화
 - 카드 지출과 차량 이용 데이터를 통합 관리함으로써 부서별, 사용 목적별 집계가 가능하며, 결재 흐름 및 신청 내역의 투명성 확보에 기여

### 2. 적용 환경 및 기술

- Windows 서버 + IIS 기반으로 운용
- ASP (Classic ASP), MSSQL, JavaScript, HTML, CSS 사용
- UI 구성에 ShadCN 스타일을 일부 반영하여 깔끔한 관리자 인터페이스 제공

---

<br>

## Key Features

| 기능 영역 | 주요 기능 설명 |
|-----------|----------------|
| **사용자 기능** | - 로그인 및 회원가입<br> - 카드 사용 내역 등록 / 조회 / 수정<br> - 개인차량 이용 신청 / 수정 / 출력 |
| **관리자 기능** | - 사용자 계정 관리 (활성화/비활성화)<br> - 카드/차량 이용 내역 검토 및 통계 출력<br> - 유류비 단가 관리 |
| **출력 및 결재 기능** | - 카드 지출 결의서 PDF 출력<br> - 개인 차량 신청서 출력<br> - 다단계 결재선 관리 및 상태 표시 |

---

<br>

## Development Setup

본 시스템은 ASP 및 MSSQL 기반의 기업용 인트라넷 전용 시스템으로, 다음과 같은 구성과 절차를 통해 개발

* **프로젝트 구조**
 - Classic ASP (VBScript) 기반 서버 사이드 스크립트로 구현
 - Microsoft IIS를 웹서버로 사용하여 내부망에서 서비스 제공
 - MSSQL Server를 RDBMS로 사용하여 테이블 기반 데이터 저장 및 쿼리 처리
 - JavaScript, HTML, CSS를 통해 프론트엔드 UI 구성
 - 일부 화면에 ShadCN 스타일 시스템을 적용하여 일관된 UX 제공
 - 공통 DB 연결과 로직은 includes/db.asp, includes/functions.asp 에 통합 관리
 
 * **개발 절차 및 방법론**
 - 요구사항 수집 및 업무 흐름 분석을 통해 카드 사용 및 차량 이용 프로세스 도식화
 - ERD 및 테이블 정의서를 기반으로 MSSQL에 테이블 생성 (엑셀 참고)
 - 각 기능은 단위 페이지별로 개발 후 ASP 내에서 파라미터 처리 방식으로 연동
 - 결재 흐름은 approvallogs 테이블을 기반으로 다단계 구조 설계
 - 파일 업로드, PDF 출력, 금액 자동 계산, 결재 상태 표시 등 실무 기능 통합 구현

 * **내 역할**
 - 전체 ASP 페이지 설계 및 개발 진행 : 카드 등록, 수정, 삭제, 출력, 차량 신청서 기능 일체 구현
 - UI는 Bootstrap 기반 스타일에 맞춰 커스터마이징 + ShadCN 적용
 - 사용자 로그인 및 세션 처리, 승인 권한 체크 로직 구현
 - 다단계 결재선 지정과 관련한 팝업 UI 및 approvallogs 연동 로직 개발
 - 차량 신청 내역의 유류비 자동 계산 알고리즘 설계 및 구현
 - 관리자 페이지 구성 및 통계 출력 기능 구현

---

<br>

## Repository Structure

```bash
card_car_used/
├── index.asp
├── logout.asp
├── register.asp
├── my_profile.asp
├── dashboard.asp
├── includes/
│   ├── db.asp
│   └── functions.asp
├── pages/
│   ├── approval_detail.asp
│   ├── approval_line_popup.asp
│   ├── approval_update.asp
│   ├── card_usage.asp
│   ├── card_usage_add.asp
│   ├── card_usage_delete.asp
│   ├── card_usage_edit.asp
│   ├── card_usage_update.asp
│   ├── card_usage_view.asp
│   ├── completed_approvals.asp
│   ├── pending_approvals.asp
│   ├── vehicle_request.asp
│   ├── vehicle_request_add.asp
│   ├── vehicle_request_delete.asp
│   ├── vehicle_request_edit.asp
│   ├── vehicle_request_update.asp
│   ├── vehicle_request_view.asp
│   └── admin/
│       ├── admin_dashboard.asp
│       ├── admin_approvals.asp
│       ├── admin_approval_view.asp
│       ├── admin_card_usage.asp
│       ├── admin_card_usage_view.asp
│       ├── admin_card_usage_process.asp
│       ├── admin_cardaccount.asp
│       ├── admin_cardaccount_process.asp
│       ├── admin_cardaccounttypes.asp
│       ├── admin_cardaccounttypes_process.asp
│       ├── admin_department.asp
│       ├── admin_fuelrate.asp
│       ├── admin_job_grade.asp
│       ├── admin_users.asp
│       ├── admin_users_process.asp
│       ├── admin_user_view.asp
│       ├── admin_vehicle_requests.asp
│       ├── admin_vehicle_request_view.asp
│       ├── admin_vehicle_request_process.asp
│       └── backup_admin_250526/, backup_admin_250527/
├── css/
│   ├── style.css
│   └── shadcn.css
└── js/
    └── common.js
```

<br>

## Authors
> 프로필 
>
> 이재민 [@깃허브 프로필 페이지](https://github.com/qwer123toy)
> 