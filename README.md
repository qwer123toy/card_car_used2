<div align="center">
  <img src="https://github.com/user-attachments/assets/600efd58-25bd-4fad-943a-8eac9b6316c2" alt="card car system banner">
</div>

# 카드 지출 결의 및 차량 이용 내역 관리 시스템
기업 내 카드 사용과 개인 차량 이용을 통합 관리하는 ASP 기반 웹 애플리케이션  
사용자와 관리자가 효율적으로 결재 내역과 신청서를 관리할 수 있는 업무 특화 시스템

---

## 📌 Contents

- [What is](#what-is)
- [Key Features](#key-features)
- [Development Setup](#development-setup)
- [Repository Structure](#repository-structure)
- [Authors](#authors)

---

## 🧾 What is

### 1. 개요 및 목적

- 카드 사용 결의와 개인 차량 이용 신청을 전산화하여 **업무 효율성과 투명성 향상**을 도모
- 내부 결재, 내역 출력, 유류비 계산 등 기능을 통합하여 **실무 중심의 기능성 시스템** 구축

### 2. 적용 환경 및 기술

- Windows 서버 + IIS 기반으로 운용
- ASP (Classic ASP), MSSQL, JavaScript, HTML, CSS 사용
- UI 구성에 ShadCN 스타일을 일부 반영하여 깔끔한 관리자 인터페이스 제공

---

## 🚀 Key Features

| 기능 영역 | 주요 기능 설명 |
|-----------|----------------|
| **사용자 기능** | - 로그인 및 회원가입<br> - 카드 사용 내역 등록 / 조회 / 수정<br> - 개인차량 이용 신청 / 수정 / 출력 |
| **관리자 기능** | - 사용자 계정 관리 (활성화/비활성화)<br> - 카드/차량 이용 내역 검토 및 통계 출력<br> - 유류비 단가 관리 |
| **출력 및 결재 기능** | - 카드 지출 결의서 PDF 출력<br> - 개인 차량 신청서 출력<br> - 다단계 결재선 관리 및 상태 표시 |

---

## 🛠️ Development Setup

1. Windows 서버에 IIS 설치
2. 프로젝트 폴더 업로드 후 IIS에서 가상 디렉토리 설정
3. `includes/db.asp` 파일 내 연결 문자열 수정:
   ```asp
   strConnection = "Provider=SQLOLEDB;Data Source=서버주소;Initial Catalog=DB명;User ID=계정;Password=비밀번호;"
   ```
4. MSSQL DB에 테이블 생성 (업로드된 엑셀 파일의 테이블 정의서 참고)

---

## 📁 Repository Structure

\`\`\`bash
card_car_used/
├─ index.asp                     # 메인 로그인 페이지
├─ includes/
│  ├─ db.asp                     # DB 연결
│  ├─ functions.asp              # 공통 함수
├─ pages/
│  ├─ dashboard.asp              # 사용자 대시보드
│  ├─ card_usage.asp             # 카드 사용 목록
│  ├─ card_usage_add.asp         # 카드 사용 등록
│  ├─ vehicle_request.asp        # 차량 신청 목록
│  └─ vehicle_request_add.asp    # 차량 신청서 작성
├─ admin/
│  ├─ users.asp                  # 사용자 계정 관리
│  ├─ card_manage.asp            # 카드 내역 관리
│  └─ fuel_rate.asp              # 유류비 단가 관리
├─ css/
│  └─ style.css, shadcn.css      # UI 스타일 파일
├─ js/
│  └─ common.js                  # 공통 스크립트
\`\`\`

---

## 👤 Authors

- 이재민  
  [GitHub 프로필](https://github.com/qwer123toy)