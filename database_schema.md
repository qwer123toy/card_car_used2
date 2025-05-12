# Database Schema Summary

## Table: Users

| Column Name | Data Type | Constraints | Description |
|-------------|-----------|-------------|-------------|
| user_id | VARCHAR(30) | PK, NOT NULL | 사용자 아이디 |
| password | VARCHAR(50) | NOT NULL | 암호화된 비밀번호 |
| name | NVARCHAR(30) | NOT NULL | 사용자 이름 |
| department_id | INT | FK, NOT NULL | 소속 부서 ID |
| job_grade | NVARCHAR(30) | NOT NULL | 직급 |

## Table: Department

| Column Name | Data Type | Constraints | Description |
|-------------|-----------|-------------|-------------|
| department_id | INT | PK, NOT NULL | 부서 고유 ID |
| name | NVARCHAR(100) | NOT NULL | 부서명 |
| parent_id | INT | FK | 상위 부서 ID |
| created_at | DATETIME | NOT NULL | 생성일시 |

## Table: CardAccountTypes

| Column Name | Data Type | Constraints | Description |
|-------------|-----------|-------------|-------------|
| account_type_id | INT | PK, NOT NULL | 카드 유형 고유 ID |
| type_name | NVARCHAR(30) | NOT NULL | 카드 유형명 |

## Table: CardAccount

| Column Name | Data Type | Constraints | Description |
|-------------|-----------|-------------|-------------|
| card_id | INT | PK, NOT NULL | 카드 고유 ID |
| account_name | NVARCHAR(100) | NOT NULL | 카드 이름 |
| account_type_id | INT | FK, NOT NULL | 카드 유형 ID |
| issuer | NVARCHAR(50) | NOT NULL | 카드 발급사 |

## Table: CardUsage

| Column Name | Data Type | Constraints | Description |
|-------------|-----------|-------------|-------------|
| usage_id | INT | PK, NOT NULL | 사용 내역 ID |
| user_id | VARCHAR(30) | FK, NOT NULL | 사용자 ID |
| card_id | INT | FK, NOT NULL | 카드 ID |
| usage_date | DATE | NOT NULL | 사용 일자 |
| amount | DECIMAL(10,2) | NOT NULL | 사용 금액 |
| usage_reason | NVARCHAR(100) | nan | 사용 사유 |
| linked_table | NVARCHAR(50) | nan | 연결된 테이블명 (예: VehicleRequests) |
| linked_id | INT | nan | 연결된 테이블의 ID |

## Table: VehicleRequests

| Column Name | Data Type | Constraints | Description |
|-------------|-----------|-------------|-------------|
| request_id | INT | PK, NOT NULL | 차량 신청 ID |
| user_id | VARCHAR(30) | FK, NOT NULL | 신청자 ID |
| request_date | DATE | NOT NULL | 신청일 |
| purpose | NVARCHAR(100) | NOT NULL | 용도 |
| start_location | NVARCHAR(100) | NOT NULL | 출발지 |
| destination | NVARCHAR(100) | NOT NULL | 도착지 |
| approval_status | NVARCHAR(20) | DEFAULT '작성중' | 결재 상태 |
| approval_log_id | INT | FK | 결재 로그 ID |
| is_deleted | BIT | DEFAULT 0 | 삭제 여부 |

## Table: FuelRate

| Column Name | Data Type | Constraints | Description |
|-------------|-----------|-------------|-------------|
| fuel_rate_id | INT | PK, NOT NULL | 유류비 ID |
| rate | DECIMAL(10,2) | NOT NULL | 유류비 단가 |
| date | DATE | GETDATE() | 단가 설정 날짜 |

## Table: ApprovalLogs

| Column Name | Data Type | Constraints | Description |
|-------------|-----------|-------------|-------------|
| approval_log_id | INT | PK, NOT NULL | 결재 로그 ID |
| approver_id | VARCHAR(30) | FK, NOT NULL | 결재자 ID |
| target_table_name | NVARCHAR(50) | NOT NULL | 결재 대상 테이블명 |
| target_id | INT | NOT NULL | 결재 대상 ID |
| approval_step | INT | nan | 결재 순서 |
| status | NVARCHAR(20) | DEFAULT '대기' | 결재 상태 |
| approved_at | DATETIME | nan | 결재 일시 |

