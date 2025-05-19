-- Users 테이블 생성
CREATE TABLE Users (
    user_id VARCHAR(30) PRIMARY KEY NOT NULL,
    password VARCHAR(50) NOT NULL,
    name NVARCHAR(30) NOT NULL,
    email NVARCHAR(100) NULL,
    department_id INT NULL,
    job_grade NVARCHAR(30) NULL,
    created_at DATETIME NOT NULL DEFAULT GETDATE()
);

-- Department 테이블 생성
CREATE TABLE Department (
    department_id INT PRIMARY KEY NOT NULL,
    name NVARCHAR(100) NOT NULL,
    parent_id INT NULL,
    created_at DATETIME NOT NULL DEFAULT GETDATE(),
    CONSTRAINT FK_Department_Parent FOREIGN KEY (parent_id) REFERENCES Department(department_id)
);

-- CardAccountTypes 테이블 생성
CREATE TABLE CardAccountTypes (
    account_type_id INT PRIMARY KEY NOT NULL,
    type_name NVARCHAR(30) NOT NULL
);

-- CardAccount 테이블 생성
CREATE TABLE CardAccount (
    card_id INT PRIMARY KEY NOT NULL,
    account_name NVARCHAR(100) NOT NULL,
    account_type_id INT NOT NULL,
    issuer NVARCHAR(50) NOT NULL,
    CONSTRAINT FK_CardAccount_Type FOREIGN KEY (account_type_id) REFERENCES CardAccountTypes(account_type_id)
);

-- CardUsage 테이블 생성
CREATE TABLE CardUsage (
    usage_id INT IDENTITY(1,1) PRIMARY KEY NOT NULL,
    user_id VARCHAR(30) NOT NULL,
    card_id INT NOT NULL,
    account_type_id INT NULL,
    expense_category_id INT NULL,
    usage_date DATE NOT NULL,
    amount DECIMAL(10,2) NOT NULL,
    usage_reason NVARCHAR(100) NULL,
    linked_table NVARCHAR(50) NULL,
    linked_id INT NULL,
    CONSTRAINT FK_CardUsage_User FOREIGN KEY (user_id) REFERENCES Users(user_id),
    CONSTRAINT FK_CardUsage_Card FOREIGN KEY (card_id) REFERENCES CardAccount(card_id),
    CONSTRAINT FK_CardUsage_Type FOREIGN KEY (account_type_id) REFERENCES CardAccountTypes(account_type_id)
);

-- VehicleRequests 테이블 생성
CREATE TABLE VehicleRequests (
    request_id INT IDENTITY(1,1) PRIMARY KEY NOT NULL,
    user_id VARCHAR(30) NOT NULL,
    request_date DATE NOT NULL,
    purpose NVARCHAR(100) NOT NULL,
    start_location NVARCHAR(100) NOT NULL,
    destination NVARCHAR(100) NOT NULL,
    distance DECIMAL(10,2) NULL,
    total_amount DECIMAL(10,2) NULL,
    approval_status NVARCHAR(20) DEFAULT '작성중',
    approval_log_id INT NULL,
    is_deleted BIT DEFAULT 0,
    CONSTRAINT FK_VehicleRequests_User FOREIGN KEY (user_id) REFERENCES Users(user_id)
);

-- FuelRate 테이블 생성
CREATE TABLE FuelRate (
    fuel_rate_id INT IDENTITY(1,1) PRIMARY KEY NOT NULL,
    rate DECIMAL(10,2) NOT NULL,
    date DATE DEFAULT GETDATE()
);

-- ApprovalLogs 테이블 생성
CREATE TABLE ApprovalLogs (
    approval_log_id INT IDENTITY(1,1) PRIMARY KEY NOT NULL,
    approver_id VARCHAR(30) NOT NULL,
    target_table_name NVARCHAR(50) NOT NULL,
    target_id INT NOT NULL,
    approval_step INT NULL,
    status NVARCHAR(20) DEFAULT '대기',
    approved_at DATETIME NULL,
    CONSTRAINT FK_ApprovalLogs_User FOREIGN KEY (approver_id) REFERENCES Users(user_id)
);

-- ActivityLogs 테이블 생성
CREATE TABLE ActivityLogs (
    log_id INT IDENTITY(1,1) PRIMARY KEY NOT NULL,
    user_id VARCHAR(30) NOT NULL,
    action NVARCHAR(50) NOT NULL,
    description NVARCHAR(200) NOT NULL,
    created_at DATETIME DEFAULT GETDATE(),
    CONSTRAINT FK_ActivityLogs_User FOREIGN KEY (user_id) REFERENCES Users(user_id)
);

-- 기본 데이터 입력
-- 부서 데이터 입력
INSERT INTO Department (department_id, name) VALUES (1, '관리부');
INSERT INTO Department (department_id, name) VALUES (2, '영업부');
INSERT INTO Department (department_id, name) VALUES (3, '기술부');

-- 카드 계정 유형 입력
INSERT INTO CardAccountTypes (account_type_id, type_name) VALUES (1, '식대');
INSERT INTO CardAccountTypes (account_type_id, type_name) VALUES (2, '교통비');
INSERT INTO CardAccountTypes (account_type_id, type_name) VALUES (3, '접대비');
INSERT INTO CardAccountTypes (account_type_id, type_name) VALUES (4, '유류비');
INSERT INTO CardAccountTypes (account_type_id, type_name) VALUES (5, '소모품비');

-- 카드 계정 입력
INSERT INTO CardAccount (card_id, account_name, account_type_id, issuer) VALUES (1, '법인카드1', 1, '신한카드');
INSERT INTO CardAccount (card_id, account_name, account_type_id, issuer) VALUES (2, '법인카드2', 1, '삼성카드');

-- 유류비 단가 입력
INSERT INTO FuelRate (rate) VALUES (2000.00);

-- 기본 사용자 등록 (admin 먼저 등록해야 참조 무결성 위반 없음)
INSERT INTO Users (user_id, password, name, department_id) VALUES ('admin', 'admin123', '관리자', 1); 