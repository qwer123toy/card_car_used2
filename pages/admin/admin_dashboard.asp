<%@ Language="VBScript" CodePage="65001" %>
<% 
Response.CodePage = 65001
Response.CharSet = "utf-8"
%>

<!--#include file="../../db.asp"-->
<!--#include file="../../includes/functions.asp"-->
<%
' 로그인 체크
If Not IsAuthenticated() Then
    RedirectTo("../../index.asp")
End If

' 사용자 정보 조회
Dim userSQL, userRS
userSQL = "SELECT name, department_id, job_grade FROM " & dbSchema & ".Users WHERE user_id = '" & Session("user_id") & "'"
Set userRS = db99.Execute(userSQL)

Dim userName
If Not userRS.EOF Then
    userName = userRS("name")
Else
    userName = Session("user_id")
End If

' 부서명 가져오기
Function GetDepartmentName(deptId)
    If IsNull(deptId) Or deptId = "" Then
        GetDepartmentName = "-"
        Exit Function
    End If
    
    Dim deptName, deptSQL, deptRS
    deptSQL = "SELECT name FROM " & dbSchema & ".Department WHERE department_id = " & deptId
    
    On Error Resume Next
    Set deptRS = db99.Execute(deptSQL)
    
    If Err.Number = 0 And Not deptRS.EOF Then
        deptName = deptRS("name")
    Else
        deptName = deptId
    End If
    
    If Not deptRS Is Nothing Then
        If deptRS.State = 1 Then
            deptRS.Close
        End If
        Set deptRS = Nothing
    End If
    
    GetDepartmentName = deptName
End Function

' 직급명 가져오기
Function GetJobGradeName(gradeId)
    If IsNull(gradeId) Or gradeId = "" Then
        GetJobGradeName = "-"
        Exit Function
    End If
    
    Dim gradeName, gradeSQL, gradeRS
    gradeSQL = "SELECT name FROM " & dbSchema & ".Job_Grade WHERE job_grade_id = " & gradeId
    
    On Error Resume Next
    Set gradeRS = db99.Execute(gradeSQL)
    
    If Err.Number = 0 And Not gradeRS.EOF Then
        gradeName = gradeRS("name")
    Else
        gradeName = gradeId
    End If
    
    If Not gradeRS Is Nothing Then
        If gradeRS.State = 1 Then
            gradeRS.Close
        End If
        Set gradeRS = Nothing
    End If
    
    GetJobGradeName = gradeName
End Function

' 통계 정보 가져오기
Dim statSQL, statRS
Dim userCount, cardCount, vehicleCount, approvalCount

' 사용자 수
statSQL = "SELECT COUNT(*) AS cnt FROM " & dbSchema & ".Users WHERE is_active = 1"
Set statRS = db99.Execute(statSQL)
If Not statRS.EOF Then
    userCount = statRS("cnt")
Else
    userCount = 0
End If
Set statRS = Nothing

' 카드 계정 수
statSQL = "SELECT COUNT(*) AS cnt FROM " & dbSchema & ".CardAccount"
Set statRS = db99.Execute(statSQL)
If Not statRS.EOF Then
    cardCount = statRS("cnt")
Else
    cardCount = 0
End If
Set statRS = Nothing

' 차량 신청 수 (최근 30일)
statSQL = "SELECT COUNT(*) AS cnt FROM " & dbSchema & ".VehicleRequests WHERE request_date >= DATEADD(day, -30, GETDATE())"
Set statRS = db99.Execute(statSQL)
If Not statRS.EOF Then
    vehicleCount = statRS("cnt")
Else
    vehicleCount = 0
End If
Set statRS = Nothing

' 결재 수 (최근 30일)
statSQL = "SELECT COUNT(*) AS cnt FROM " & dbSchema & ".ApprovalLogs WHERE created_at >= DATEADD(day, -30, GETDATE())"
Set statRS = db99.Execute(statSQL)
If Not statRS.EOF Then
    approvalCount = statRS("cnt")
Else
    approvalCount = 0
End If
Set statRS = Nothing
%>

<!--#include file="../../includes/header.asp"-->

<style>
.admin-container {
    max-width: 1400px;
    margin: 0 auto;
    padding: 2rem 1rem;
}

.admin-nav {
    background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
    border-radius: 16px;
    padding: 1.5rem;
    margin-bottom: 2rem;
    box-shadow: 0 8px 32px rgba(0,0,0,0.1);
}

.admin-nav-title {
    color: white;
    font-size: 1.25rem;
    font-weight: 600;
    margin-bottom: 1.5rem;
    display: flex;
    align-items: center;
}

.admin-nav-grid {
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
    gap: 0.75rem;
}

.admin-nav-item {
    background: rgba(255,255,255,0.1);
    border: 1px solid rgba(255,255,255,0.2);
    border-radius: 12px;
    padding: 1rem;
    color: white;
    text-decoration: none;
    transition: all 0.3s ease;
    display: flex;
    align-items: center;
    font-size: 0.9rem;
    font-weight: 500;
}

.admin-nav-item:hover {
    background: rgba(255,255,255,0.2);
    transform: translateY(-2px);
    color: white;
    text-decoration: none;
}

.admin-nav-item.active {
    background: rgba(255,255,255,0.25);
    border-color: rgba(255,255,255,0.4);
}

.admin-nav-item i {
    margin-right: 0.75rem;
    font-size: 1.1rem;
}

.welcome-section {
    background: linear-gradient(135deg, #f093fb 0%, #f5576c 100%);
    border-radius: 16px;
    padding: 2rem;
    margin-bottom: 2rem;
    color: white;
    box-shadow: 0 8px 32px rgba(0,0,0,0.1);
}

.welcome-title {
    font-size: 1.75rem;
    font-weight: 700;
    margin-bottom: 0.5rem;
}

.welcome-subtitle {
    font-size: 1.1rem;
    opacity: 0.9;
    margin-bottom: 0;
}

.stats-grid {
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(280px, 1fr));
    gap: 1.5rem;
    margin-bottom: 2rem;
}

.stat-card {
    background: white;
    border-radius: 16px;
    padding: 2rem;
    box-shadow: 0 4px 20px rgba(0,0,0,0.08);
    transition: all 0.3s ease;
    border: none;
    position: relative;
    overflow: hidden;
}

.stat-card::before {
    content: '';
    position: absolute;
    top: 0;
    left: 0;
    right: 0;
    height: 4px;
    background: var(--card-color);
}

.stat-card:hover {
    transform: translateY(-4px);
    box-shadow: 0 8px 30px rgba(0,0,0,0.12);
}

.stat-card.primary {
    --card-color: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
}

.stat-card.success {
    --card-color: linear-gradient(135deg, #4facfe 0%, #00f2fe 100%);
}

.stat-card.warning {
    --card-color: linear-gradient(135deg, #43e97b 0%, #38f9d7 100%);
}

.stat-card.info {
    --card-color: linear-gradient(135deg, #fa709a 0%, #fee140 100%);
}

.stat-header {
    display: flex;
    justify-content: space-between;
    align-items: flex-start;
    margin-bottom: 1rem;
}

.stat-icon {
    width: 60px;
    height: 60px;
    border-radius: 12px;
    display: flex;
    align-items: center;
    justify-content: center;
    font-size: 1.5rem;
    color: white;
}

.stat-icon.primary {
    background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
}

.stat-icon.success {
    background: linear-gradient(135deg, #4facfe 0%, #00f2fe 100%);
}

.stat-icon.warning {
    background: linear-gradient(135deg, #43e97b 0%, #38f9d7 100%);
}

.stat-icon.info {
    background: linear-gradient(135deg, #fa709a 0%, #fee140 100%);
}

.stat-content {
    flex: 1;
}

.stat-title {
    font-size: 0.9rem;
    color: #64748B;
    font-weight: 600;
    margin-bottom: 0.5rem;
    text-transform: uppercase;
    letter-spacing: 0.5px;
}

.stat-number {
    font-size: 2.5rem;
    font-weight: 700;
    color: #1E293B;
    margin-bottom: 0.5rem;
    line-height: 1;
}

.stat-footer {
    display: flex;
    justify-content: space-between;
    align-items: center;
    margin-top: 1.5rem;
    padding-top: 1rem;
    border-top: 1px solid #E2E8F0;
}

.stat-label {
    font-size: 0.875rem;
    color: #64748B;
    font-weight: 500;
}

.stat-link {
    color: #3B82F6;
    text-decoration: none;
    font-weight: 600;
    font-size: 0.875rem;
    display: flex;
    align-items: center;
    transition: all 0.2s ease;
}

.stat-link:hover {
    color: #1D4ED8;
    text-decoration: none;
    transform: translateX(2px);
}

.stat-link i {
    margin-left: 0.5rem;
    transition: transform 0.2s ease;
}

.stat-link:hover i {
    transform: translateX(2px);
}

.quick-actions {
    background: white;
    border-radius: 16px;
    padding: 2rem;
    box-shadow: 0 4px 20px rgba(0,0,0,0.08);
}

.quick-actions-title {
    font-size: 1.25rem;
    font-weight: 600;
    color: #1E293B;
    margin-bottom: 1.5rem;
    display: flex;
    align-items: center;
}

.quick-actions-grid {
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
    gap: 1rem;
}

.quick-action-btn {
    background: #F8FAFC;
    border: 2px solid #E2E8F0;
    border-radius: 12px;
    padding: 1.5rem;
    text-decoration: none;
    color: #475569;
    transition: all 0.3s ease;
    display: flex;
    flex-direction: column;
    align-items: center;
    text-align: center;
}

.quick-action-btn:hover {
    background: #F1F5F9;
    border-color: #CBD5E1;
    transform: translateY(-2px);
    color: #334155;
    text-decoration: none;
}

.quick-action-btn i {
    font-size: 2rem;
    margin-bottom: 1rem;
    color: #64748B;
}

.quick-action-btn span {
    font-weight: 600;
    font-size: 0.9rem;
}
</style>

<div class="admin-container">
    <!-- 관리자 네비게이션 -->
    <div class="admin-nav">
        <div class="admin-nav-title">
            <i class="fas fa-cog me-2"></i>관리자 메뉴
        </div>
        <div class="admin-nav-grid">
            <a href="admin_dashboard.asp" class="admin-nav-item active">
                <i class="fas fa-tachometer-alt"></i>대시보드
            </a>
            <a href="admin_cardaccount.asp" class="admin-nav-item">
                <i class="fas fa-credit-card"></i>카드 계정 관리
            </a>
            <a href="admin_cardaccounttypes.asp" class="admin-nav-item">
                <i class="fas fa-tags"></i>카드 계정 유형 관리
            </a>
            <a href="admin_fuelrate.asp" class="admin-nav-item">
                <i class="fas fa-gas-pump"></i>유류비 단가 관리
            </a>
            <a href="admin_job_grade.asp" class="admin-nav-item">
                <i class="fas fa-user-tie"></i>직급 관리
            </a>
            <a href="admin_department.asp" class="admin-nav-item">
                <i class="fas fa-sitemap"></i>부서 관리
            </a>
            <a href="admin_users.asp" class="admin-nav-item">
                <i class="fas fa-users"></i>사용자 관리
            </a>
            <a href="admin_card_usage.asp" class="admin-nav-item">
                <i class="fas fa-receipt"></i>카드 사용 내역 관리
            </a>
            <a href="admin_vehicle_requests.asp" class="admin-nav-item">
                <i class="fas fa-car"></i>차량 이용 신청 관리
            </a>
            <a href="admin_approvals.asp" class="admin-nav-item">
                <i class="fas fa-file-signature"></i>결재 로그 관리
            </a>
        </div>
    </div>

    <!-- 환영 섹션 -->
    <div class="welcome-section">
        <div class="welcome-title">
            <i class="fas fa-tachometer-alt me-3"></i>관리자 대시보드
        </div>
        <div class="welcome-subtitle">
            시스템의 주요 설정과 통계를 한눈에 확인하고 관리할 수 있습니다.
        </div>
    </div>

    <!-- 통계 카드 -->
    <div class="stats-grid">
        <div class="stat-card primary">
            <div class="stat-header">
                <div class="stat-content">
                    <div class="stat-title">등록 사용자</div>
                    <div class="stat-number"><%= userCount %></div>
                </div>
                <div class="stat-icon primary">
                    <i class="fas fa-users"></i>
                </div>
            </div>
            <div class="stat-footer">
                <span class="stat-label">활성 사용자</span>
                <a href="admin_users.asp" class="stat-link">
                    관리하기 <i class="fas fa-arrow-right"></i>
                </a>
            </div>
        </div>

        <div class="stat-card success">
            <div class="stat-header">
                <div class="stat-content">
                    <div class="stat-title">카드 계정</div>
                    <div class="stat-number"><%= cardCount %></div>
                </div>
                <div class="stat-icon success">
                    <i class="fas fa-credit-card"></i>
                </div>
            </div>
            <div class="stat-footer">
                <span class="stat-label">등록된 카드</span>
                <a href="admin_cardaccount.asp" class="stat-link">
                    관리하기 <i class="fas fa-arrow-right"></i>
                </a>
            </div>
        </div>

        <div class="stat-card warning">
            <div class="stat-header">
                <div class="stat-content">
                    <div class="stat-title">차량 신청</div>
                    <div class="stat-number"><%= vehicleCount %></div>
                </div>
                <div class="stat-icon warning">
                    <i class="fas fa-car"></i>
                </div>
            </div>
            <div class="stat-footer">
                <span class="stat-label">최근 30일</span>
                <a href="admin_vehicle_requests.asp" class="stat-link">
                    관리하기 <i class="fas fa-arrow-right"></i>
                </a>
            </div>
        </div>

        <div class="stat-card info">
            <div class="stat-header">
                <div class="stat-content">
                    <div class="stat-title">결재 처리</div>
                    <div class="stat-number"><%= approvalCount %></div>
                </div>
                <div class="stat-icon info">
                    <i class="fas fa-file-signature"></i>
                </div>
            </div>
            <div class="stat-footer">
                <span class="stat-label">최근 30일</span>
                <a href="admin_approvals.asp" class="stat-link">
                    관리하기 <i class="fas fa-arrow-right"></i>
                </a>
            </div>
        </div>
    </div>

    <!-- 빠른 작업 -->
    <div class="quick-actions">
        <div class="quick-actions-title">
            <i class="fas fa-bolt me-2"></i>빠른 작업
        </div>
        <div class="quick-actions-grid">
            <a href="admin_users.asp" class="quick-action-btn">
                <i class="fas fa-user-plus"></i>
                <span>사용자 추가</span>
            </a>
            <a href="admin_cardaccount.asp" class="quick-action-btn">
                <i class="fas fa-credit-card"></i>
                <span>카드 계정 추가</span>
            </a>
            <a href="admin_department.asp" class="quick-action-btn">
                <i class="fas fa-sitemap"></i>
                <span>부서 관리</span>
            </a>
            <a href="admin_job_grade.asp" class="quick-action-btn">
                <i class="fas fa-user-tie"></i>
                <span>직급 관리</span>
            </a>
            <a href="admin_fuelrate.asp" class="quick-action-btn">
                <i class="fas fa-gas-pump"></i>
                <span>유류비 설정</span>
            </a>
            <a href="admin_approvals.asp" class="quick-action-btn">
                <i class="fas fa-file-signature"></i>
                <span>결재 로그</span>
            </a>
        </div>
    </div>
</div>

<%
' 사용한 객체 해제
If Not userRS Is Nothing Then
    If userRS.State = 1 Then
        userRS.Close
    End If
    Set userRS = Nothing
End If
%>

<!--#include file="../../includes/footer.asp"--> 