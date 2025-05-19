-- 계정과목 테이블 생성
IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='AccountSubject' AND xtype='U')
BEGIN
    CREATE TABLE [dbo].[AccountSubject](
        [subject_code] [nvarchar](20) NOT NULL,
        [subject_name] [nvarchar](100) NOT NULL,
        [created_at] [datetime] DEFAULT GETDATE(),
        CONSTRAINT [PK_AccountSubject] PRIMARY KEY CLUSTERED 
        (
            [subject_code] ASC
        )
    )
    
    -- 기존 CardAccountTypes 테이블에서 데이터 마이그레이션
    INSERT INTO [dbo].[AccountSubject] ([subject_code], [subject_name])
    SELECT CAST(account_type_id AS nvarchar(20)), type_name
    FROM [dbo].[CardAccountTypes]
    
    -- 기존 CardAccountType 테이블에서 데이터 마이그레이션
    INSERT INTO [dbo].[AccountSubject] ([subject_code], [subject_name])
    SELECT CAST(id AS nvarchar(20)), name
    FROM [dbo].[CardAccountType]
    WHERE NOT EXISTS (
        SELECT 1 FROM [dbo].[AccountSubject] a 
        WHERE a.subject_code = CAST(CardAccountType.id AS nvarchar(20))
    )
    
    -- 기존 ExpenseCategory 테이블에서 데이터 마이그레이션
    INSERT INTO [dbo].[AccountSubject] ([subject_code], [subject_name])
    SELECT CAST(expense_category_id AS nvarchar(20)), category_name
    FROM [dbo].[ExpenseCategory]
    WHERE NOT EXISTS (
        SELECT 1 FROM [dbo].[AccountSubject] a 
        WHERE a.subject_code = CAST(ExpenseCategory.expense_category_id AS nvarchar(20))
    )
    
    -- 기본 계정과목 데이터 추가
    IF NOT EXISTS (SELECT 1 FROM [dbo].[AccountSubject])
    BEGIN
        INSERT INTO [dbo].[AccountSubject] ([subject_code], [subject_name]) VALUES ('1001', N'식대')
        INSERT INTO [dbo].[AccountSubject] ([subject_code], [subject_name]) VALUES ('1002', N'교통비')
        INSERT INTO [dbo].[AccountSubject] ([subject_code], [subject_name]) VALUES ('1003', N'접대비')
        INSERT INTO [dbo].[AccountSubject] ([subject_code], [subject_name]) VALUES ('1004', N'소모품비')
        INSERT INTO [dbo].[AccountSubject] ([subject_code], [subject_name]) VALUES ('1005', N'회의비')
        INSERT INTO [dbo].[AccountSubject] ([subject_code], [subject_name]) VALUES ('1006', N'출장비')
    END
    
    PRINT '계정과목 테이블 생성 및 데이터 마이그레이션 완료'
END
ELSE
BEGIN
    PRINT '계정과목 테이블이 이미 존재합니다'
END 