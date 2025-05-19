-- 직급 테이블 생성
IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='JobGrade' AND xtype='U')
BEGIN
    CREATE TABLE [dbo].[JobGrade](
        [job_grade_id] [int] IDENTITY(1,1) NOT NULL,
        [name] [nvarchar](50) NOT NULL,
        [created_at] [datetime] DEFAULT GETDATE(),
        CONSTRAINT [PK_JobGrade] PRIMARY KEY CLUSTERED 
        (
            [job_grade_id] ASC
        )
    )
    
    -- 기본 직급 데이터 추가
    INSERT INTO [dbo].[JobGrade] ([name]) VALUES (N'사원')
    INSERT INTO [dbo].[JobGrade] ([name]) VALUES (N'대리')
    INSERT INTO [dbo].[JobGrade] ([name]) VALUES (N'과장')
    INSERT INTO [dbo].[JobGrade] ([name]) VALUES (N'차장')
    INSERT INTO [dbo].[JobGrade] ([name]) VALUES (N'부장')
    INSERT INTO [dbo].[JobGrade] ([name]) VALUES (N'이사')
    INSERT INTO [dbo].[JobGrade] ([name]) VALUES (N'상무')
    INSERT INTO [dbo].[JobGrade] ([name]) VALUES (N'전무')
    INSERT INTO [dbo].[JobGrade] ([name]) VALUES (N'사장')
    
    PRINT '직급 테이블 생성 및 기본 데이터 추가 완료'
END
ELSE
BEGIN
    PRINT '직급 테이블이 이미 존재합니다'
END 