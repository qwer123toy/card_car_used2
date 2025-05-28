<%@ Language="VBScript" CodePage="65001" %>
<% 
Response.CodePage = 65001
Response.CharSet = "utf-8"
%>

<!--#include file="../../db.asp"-->
<!--#include file="../../includes/functions.asp"-->
<%
If Not IsAuthenticated() Then RedirectTo("../../index.asp")
If Not IsAdmin() Then
    Response.Write("<script>alert('관리자 권한이 필요합니다.'); window.location.href='../dashboard.asp';</script>")
    Response.End
End If

Dim targetId, targetTable

targetId = Request.QueryString("target_id")
targetTable = Request.QueryString("target_table_name")

If targetId = "" Or targetTable = "" Then
    Response.Write("<script>alert('잘못된 접근입니다.'); location.href='admin_approvals.asp';</script>")
    Response.End
End If

Dim docSQL, docRS
Select Case LCase(targetTable)
    Case "cardusage"
        docSQL = "SELECT cu.*, u.name AS requester_name, u.email AS requester_email, d.name AS requester_department, " & _
                 "ca.account_name, j.name AS job_grade_name " & _
                 "FROM " & dbSchema & ".CardUsage cu " & _
                 "LEFT JOIN " & dbSchema & ".Users u ON cu.user_id = u.user_id " & _
                 "LEFT JOIN " & dbSchema & ".Department d ON u.department_id = d.department_id " & _
                 "LEFT JOIN " & dbSchema & ".Job_Grade j ON u.job_grade = j.job_grade_id " & _
                 "LEFT JOIN " & dbSchema & ".CardAccount ca ON cu.card_id = ca.card_id " & _
                 "WHERE cu.usage_id = " & targetId
    Case "vehiclerequests"
        docSQL = "SELECT vr.*, u.name AS requester_name, u.email AS requester_email, d.name AS requester_department, " & _
                 "j.name AS job_grade_name, fr.rate AS fuel_rate " & _
                 "FROM " & dbSchema & ".VehicleRequests vr " & _
                 "LEFT JOIN " & dbSchema & ".Users u ON vr.user_id = u.user_id " & _
                 "LEFT JOIN " & dbSchema & ".Department d ON u.department_id = d.department_id " & _
                 "LEFT JOIN " & dbSchema & ".Job_Grade j ON u.job_grade = j.job_grade_id " & _
                 "LEFT JOIN " & dbSchema & ".FuelRate fr ON fr.date <= vr.start_date " & _
                 "WHERE vr.request_id = " & targetId & " ORDER BY fr.date DESC"
    Case Else
        Response.Write("<script>alert('지원하지 않는 문서 유형입니다.'); location.href='admin_approvals.asp';</script>")
        Response.End
End Select

Set docRS = db.Execute(docSQL)
If docRS.EOF Then
    Response.Write("<script>alert('문서를 찾을 수 없습니다.'); location.href='admin_approvals.asp';</script>")
    Response.End
End If

' 결재 정보 조회
Dim approvalRS, approvalSQL
approvalSQL = "SELECT al.*, u.name AS approver_name, u.department_id, " & _
             "d.name AS department_name, u.job_grade, j.name AS job_grade_name " & _
             "FROM " & dbSchema & ".ApprovalLogs al " & _
             "JOIN " & dbSchema & ".Users u ON al.approver_id = u.user_id " & _
             "LEFT JOIN " & dbSchema & ".Department d ON u.department_id = d.department_id " & _
             "LEFT JOIN " & dbSchema & ".Job_Grade j ON u.job_grade = j.job_grade_id " & _
             "WHERE al.target_table_name = '" & targetTable & "' AND al.target_id = " & targetId & " " & _
             "ORDER BY al.approval_step"

Set approvalRS = db.Execute(approvalSQL)

Function FormatDate(dateValue)
    If IsNull(dateValue) Or Not IsDate(dateValue) Then
        FormatDate = "-"
    Else
        FormatDate = FormatDateTime(dateValue, 2)
    End If
End Function

Function FormatDateTimeValue(dateValue)
    If IsNull(dateValue) Or Not IsDate(dateValue) Then
        FormatDateTimeValue = "-"
    Else
        FormatDateTimeValue = FormatDateTime(dateValue, 0)
    End If
End Function

Function GetStatusBadge(status)
    Select Case status
        Case "승인"
            GetStatusBadge = "<span class='badge bg-success'><i class='fas fa-check me-1'></i>승인</span>"
        Case "반려"
            GetStatusBadge = "<span class='badge bg-danger'><i class='fas fa-times me-1'></i>반려</span>"
        Case "대기"
            GetStatusBadge = "<span class='badge bg-warning'><i class='fas fa-clock me-1'></i>대기</span>"
        Case "완료"
            GetStatusBadge = "<span class='badge bg-primary'><i class='fas fa-check-double me-1'></i>완료</span>"
        Case Else
            GetStatusBadge = "<span class='badge bg-secondary'>" & status & "</span>"
    End Select
End Function

Dim isCardUsage, isVehicleRequest
isCardUsage = (LCase(targetTable) = "cardusage")
isVehicleRequest = (LCase(targetTable) = "vehiclerequests")
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

.card {
    border: none;
    box-shadow: 0 0 20px rgba(0,0,0,0.05);
    border-radius: 16px;
    margin-bottom: 2rem;
    background: #fff;
    overflow: hidden;
}

.card-header {
    background: linear-gradient(to right, #4A90E2, #5A9EEA);
    border-bottom: none;
    padding: 1.5rem;
}

.card-header h5 {
    color: #fff;
    font-weight: 600;
    margin: 0;
    font-size: 1.25rem;
}

.card-body {
    padding: 2rem;
}

.page-header {
    display: flex;
    justify-content: space-between;
    align-items: center;
    margin-bottom: 2rem;
    padding: 1.5rem;
    background: white;
    border-radius: 12px;
    box-shadow: 0 2px 4px rgba(0,0,0,0.05);
}

.page-title {
    font-size: 1.5rem;
    font-weight: 600;
    color: #2C3E50;
    margin: 0;
}

.btn {
    padding: 0.875rem 1.5rem;
    font-weight: 600;
    border-radius: 8px;
    transition: all 0.2s ease;
    margin: 0 0.25rem;
}

.btn-secondary {
    background: #F8FAFC;
    border: 2px solid #E9ECEF;
    color: #2C3E50;
}

.btn-secondary:hover {
    background: #E9ECEF;
    transform: translateY(-2px);
}

.info-grid {
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(300px, 1fr));
    gap: 2rem;
    margin-bottom: 2rem;
}

.detail-section {
    background: #F8FAFC;
    border-radius: 12px;
    padding: 1.5rem;
}

.detail-section h6 {
    color: #2C3E50;
    font-weight: 600;
    margin-bottom: 1rem;
    padding-bottom: 0.5rem;
    border-bottom: 2px solid #E9ECEF;
}

.detail-item {
    display: flex;
    justify-content: space-between;
    align-items: center;
    padding: 0.75rem 0;
    border-bottom: 1px solid #E9ECEF;
}

.detail-item:last-child {
    border-bottom: none;
}

.detail-label {
    font-weight: 600;
    color: #64748B;
    min-width: 120px;
}

.detail-value {
    color: #2C3E50;
    text-align: right;
}

.badge {
    padding: 0.5rem 1rem;
    font-weight: 500;
    border-radius: 6px;
    font-size: 0.875rem;
}

.bg-success {
    background: #E3F9E5 !important;
    color: #1B873F;
}

.bg-danger {
    background: #FFE9E9 !important;
    color: #DA3633;
}

.bg-warning {
    background: #FFF3CD !important;
    color: #856404;
}

.bg-primary {
    background: #D1ECF1 !important;
    color: #0C5460;
}

.bg-secondary {
    background: #F1F5F9 !important;
    color: #475569;
}

.approval-line {
    background: #F8FAFC;
    border-radius: 12px;
    padding: 1.5rem;
    margin-bottom: 2rem;
}

.approval-steps {
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(280px, 1fr));
    gap: 1rem;
    margin-bottom: 1rem;
}

.approval-step {
    background: white;
    border: 2px solid #E9ECEF;
    border-radius: 10px;
    padding: 1.25rem;
    transition: all 0.2s ease;
}

.approval-step:hover {
    border-color: #4A90E2;
    box-shadow: 0 4px 12px rgba(74,144,226,0.1);
}

.step-label {
    display: inline-block;
    font-weight: 600;
    color: #2C3E50;
    margin-bottom: 1rem;
    font-size: 0.95rem;
    padding: 0.5rem 1rem;
    background: #E9ECEF;
    border-radius: 6px;
}

.approver-info {
    margin-top: 0.5rem;
}

.approver-name {
    font-weight: 600;
    color: #2C3E50;
    margin-bottom: 0.25rem;
}

.approver-dept {
    font-size: 0.9rem;
    color: #64748B;
}

.approval-status {
    display: flex;
    align-items: center;
    justify-content: space-between;
    margin-top: 1rem;
    padding-top: 1rem;
    border-top: 1px solid #E9ECEF;
}

.approval-date {
    font-size: 0.875rem;
    color: #64748B;
}

.approval-comment {
    margin-top: 0.75rem;
    padding: 0.75rem;
    background: #F1F5F9;
    border-radius: 6px;
    font-size: 0.875rem;
    color: #475569;
}

.approval-comment i {
    margin-right: 0.5rem;
    color: #64748B;
}

/* 결재 진행 상황 */
.approval-progress {
    background: #F8FAFC;
    border-radius: 12px;
    padding: 1.5rem;
    border: 1px solid #E9ECEF;
}

.progress-header h6 {
    color: #2C3E50;
    font-weight: 600;
    margin-bottom: 1rem;
}

.progress {
    background-color: #E9ECEF;
}

.progress-bar {
    background: linear-gradient(to right, #2ECC71, #27AE60) !important;
}

.progress-info {
    text-align: center;
}

/* 결재 단계 상세 */
.approval-step.completed {
    border-color: #2ECC71;
    background: #F0FDF4;
}

.approval-step.rejected {
    border-color: #E74C3C;
    background: #FEF2F2;
}

.approval-step.pending {
    border-color: #F59E0B;
    background: #FFFBEB;
}

.step-header {
    display: flex;
    justify-content: space-between;
    align-items: center;
    margin-bottom: 1rem;
    padding-bottom: 0.75rem;
    border-bottom: 1px solid #E9ECEF;
}

.approver-details {
    margin-bottom: 1rem;
}

.approver-main-info {
    display: flex;
    align-items: center;
    margin-bottom: 1rem;
}

.approver-avatar {
    margin-right: 1rem;
    font-size: 2rem;
    color: #64748B;
}

.approval-timeline {
    background: #F8FAFC;
    border-radius: 8px;
    padding: 1rem;
    margin-top: 1rem;
}

.timeline-item {
    display: flex;
    justify-content: space-between;
    align-items: center;
    padding: 0.5rem 0;
    border-bottom: 1px solid #E9ECEF;
}

.timeline-item:last-child {
    border-bottom: none;
}

.timeline-label {
    font-weight: 600;
    color: #64748B;
    font-size: 0.875rem;
}

.timeline-value {
    color: #2C3E50;
    font-size: 0.875rem;
}

.comment-header {
    font-weight: 600;
    color: #2C3E50;
    margin-bottom: 0.5rem;
    font-size: 0.9rem;
}

.comment-content {
    background: #F1F5F9;
    padding: 0.75rem;
    border-radius: 6px;
    font-size: 0.875rem;
    color: #475569;
    border-left: 3px solid #4A90E2;
}

.processing-time {
    margin-top: 0.75rem;
    text-align: right;
}

/* 결재 요약 */
.approval-summary {
    background: #F8FAFC;
    border-radius: 12px;
    padding: 1.5rem;
    border: 1px solid #E9ECEF;
}

.summary-header h6 {
    color: #2C3E50;
    font-weight: 600;
    margin-bottom: 1rem;
}

.summary-item {
    text-align: center;
    padding: 1rem;
    background: white;
    border-radius: 8px;
    border: 1px solid #E9ECEF;
}

.summary-label {
    font-size: 0.875rem;
    color: #64748B;
    font-weight: 500;
    margin-bottom: 0.5rem;
}

.summary-value {
    font-size: 1.25rem;
    font-weight: 700;
    color: #2C3E50;
}
</style>

<div class="admin-container">
    <!-- 관리자 네비게이션 -->
    <div class="admin-nav">
        <div class="admin-nav-title">
            <i class="fas fa-cog me-2"></i>관리자 메뉴
        </div>
        <div class="admin-nav-grid">
            <a href="admin_dashboard.asp" class="admin-nav-item">
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
            <a href="admin_approvals.asp" class="admin-nav-item active">
                <i class="fas fa-file-signature"></i>결재 로그 관리
            </a>
        </div>
    </div>

    <!-- 페이지 헤더 -->
    <div class="page-header">
        <h2 class="page-title">
            <i class="fas fa-file-signature me-2"></i>결재 문서 상세보기
        </h2>
        <div>
            <a href="admin_approvals.asp" class="btn btn-secondary">
                <i class="fas fa-arrow-left me-1"></i> 목록으로
            </a>
        </div>
    </div>

    <!-- 기본 정보 -->
    <div class="card">
        <div class="card-header">
            <h5><i class="fas fa-info-circle me-2"></i>기본 정보</h5>
        </div>
        <div class="card-body">
            <div class="info-grid">
                <div class="detail-section">
                    <h6><i class="fas fa-file-alt me-2"></i>문서 정보</h6>
                    <div class="detail-item">
                        <span class="detail-label">문서 유형</span>
                        <span class="detail-value">
                            <% If isCardUsage Then %>
                                <i class="fas fa-credit-card me-1"></i>카드 사용 내역
                            <% ElseIf isVehicleRequest Then %>
                                <i class="fas fa-car me-1"></i>차량 이용 신청
                            <% End If %>
                        </span>
                    </div>
                    <div class="detail-item">
                        <span class="detail-label">제목</span>
                        <span class="detail-value"><%= IIf(IsNull(docRS("title")), "-", docRS("title")) %></span>
                    </div>
                    <div class="detail-item">
                        <span class="detail-label">신청일</span>
                        <span class="detail-value"><%= FormatDate(docRS("created_at")) %></span>
                    </div>
                    <div class="detail-item">
                        <span class="detail-label">상태</span>
                        <span class="detail-value"><%= GetStatusBadge(docRS("approval_status")) %></span>
                    </div>
                    <% If Not IsNull(docRS("purpose")) And docRS("purpose") <> "" Then %>
                    <div class="detail-item">
                        <span class="detail-label">사용 목적</span>
                        <span class="detail-value"><%= docRS("purpose") %></span>
                    </div>
                    <% End If %>
                </div>

                <div class="detail-section">
                    <h6><i class="fas fa-user me-2"></i>신청자 정보</h6>
                    <div class="detail-item">
                        <span class="detail-label">이름</span>
                        <span class="detail-value"><%= IIf(IsNull(docRS("requester_name")), "-", docRS("requester_name")) %></span>
                    </div>
                    <div class="detail-item">
                        <span class="detail-label">이메일</span>
                        <span class="detail-value"><%= IIf(IsNull(docRS("requester_email")), "-", docRS("requester_email")) %></span>
                    </div>
                    <div class="detail-item">
                        <span class="detail-label">부서</span>
                        <span class="detail-value"><%= IIf(IsNull(docRS("requester_department")), "-", docRS("requester_department")) %></span>
                    </div>
                    <% If Not IsNull(docRS("job_grade_name")) Then %>
                    <div class="detail-item">
                        <span class="detail-label">직급</span>
                        <span class="detail-value"><%= docRS("job_grade_name") %></span>
                    </div>
                    <% End If %>
                </div>
            </div>
        </div>
    </div>

    <!-- 상세 정보 -->
    <% If isCardUsage Then %>
    <div class="card">
        <div class="card-header">
            <h5><i class="fas fa-credit-card me-2"></i>카드 사용 상세</h5>
        </div>
        <div class="card-body">
            <div class="info-grid">
                <div class="detail-section">
                    <h6><i class="fas fa-credit-card me-2"></i>카드 정보</h6>
                    <div class="detail-item">
                        <span class="detail-label">카드 계정</span>
                        <span class="detail-value"><%= IIf(IsNull(docRS("account_name")), "-", docRS("account_name")) %></span>
                    </div>
                    <% If Not IsNull(docRS("usage_date")) Then %>
                    <div class="detail-item">
                        <span class="detail-label">사용일</span>
                        <span class="detail-value"><%= FormatDate(docRS("usage_date")) %></span>
                    </div>
                    <% End If %>
                    <% If Not IsNull(docRS("amount")) Then %>
                    <div class="detail-item">
                        <span class="detail-label">사용 금액</span>
                        <span class="detail-value"><%= FormatNumber(docRS("amount")) %>원</span>
                    </div>
                    <% End If %>
                </div>
                <div class="detail-section">
                    <h6><i class="fas fa-map-marker-alt me-2"></i>사용 정보</h6>
                    <% If Not IsNull(docRS("store_name")) Then %>
                    <div class="detail-item">
                        <span class="detail-label">가맹점</span>
                        <span class="detail-value"><%= docRS("store_name") %></span>
                    </div>
                    <% End If %>
                    
                </div>
            </div>
        </div>
    </div>
    <% ElseIf isVehicleRequest Then %>
    <div class="card">
        <div class="card-header">
            <h5><i class="fas fa-car me-2"></i>차량 이용 상세</h5>
        </div>
        <div class="card-body">
            <div class="info-grid">
                <div class="detail-section">
                    <h6><i class="fas fa-calendar me-2"></i>이용 일정</h6>
                    <% If Not IsNull(docRS("start_date")) Then %>
                    <div class="detail-item">
                        <span class="detail-label">시작일</span>
                        <span class="detail-value"><%= FormatDate(docRS("start_date")) %></span>
                    </div>
                    <% End If %>
                    <% If Not IsNull(docRS("end_date")) Then %>
                    <div class="detail-item">
                        <span class="detail-label">종료일</span>
                        <span class="detail-value"><%= FormatDate(docRS("end_date")) %></span>
                    </div>
                    <% End If %>
                    <% If Not IsNull(docRS("destination")) Then %>
                    <div class="detail-item">
                        <span class="detail-label">목적지</span>
                        <span class="detail-value"><%= docRS("destination") %></span>
                    </div>
                    <% End If %>
                </div>
                <div class="detail-section">
                    <h6><i class="fas fa-gas-pump me-2"></i>비용 정보</h6>
                    <% If Not IsNull(docRS("distance")) Then %>
                    <div class="detail-item">
                        <span class="detail-label">거리</span>
                        <span class="detail-value"><%= docRS("distance") %>km</span>
                    </div>
                    <% End If %>
                    <% If Not IsNull(docRS("fuel_rate")) Then %>
                    <div class="detail-item">
                        <span class="detail-label">유류비 단가</span>
                        <span class="detail-value"><%= FormatNumber(docRS("fuel_rate")) %>원/km</span>
                    </div>
                    <% End If %>
                    <% If Not IsNull(docRS("distance")) And Not IsNull(docRS("fuel_rate")) Then %>
                    <div class="detail-item">
                        <span class="detail-label">예상 유류비</span>
                        <span class="detail-value"><%= FormatNumber((docRS("distance") * docRS("fuel_rate"))) %>원</span>
                    </div>
                    <% End If %>
                </div>
            </div>
        </div>
    </div>
    <% End If %>

    <!-- 결재 라인 -->
    <% If Not approvalRS.EOF Then %>
    <div class="card">
        <div class="card-header">
            <h5><i class="fas fa-route me-2"></i>결재 라인</h5>
        </div>
        <div class="card-body">
            <div class="approval-line">
                <!-- 결재 진행 상황 표시 -->
                <div class="approval-progress mb-4">
                    <div class="progress-header">
                        <h6><i class="fas fa-chart-line me-2"></i>결재 진행 상황</h6>
                    </div>
                    <div class="progress-bar-container">
                        <%
                        ' 결재 단계별 상태 확인
                        approvalRS.MoveFirst
                        Dim totalSteps, completedSteps, currentStep
                        totalSteps = 0
                        completedSteps = 0
                        currentStep = 0
                        
                        Do While Not approvalRS.EOF
                            totalSteps = totalSteps + 1
                            If approvalRS("status") = "승인" Then
                                completedSteps = completedSteps + 1
                            ElseIf approvalRS("status") = "대기" And currentStep = 0 Then
                                currentStep = totalSteps
                            End If
                            approvalRS.MoveNext
                        Loop
                        
                        Dim progressPercent
                        If totalSteps > 0 Then
                            progressPercent = (completedSteps * 100) / totalSteps
                        Else
                            progressPercent = 0
                        End If
                        %>
                        <div class="progress" style="height: 8px; border-radius: 4px;">
                            <div class="progress-bar bg-success" role="progressbar" style="width: <%= progressPercent %>%" aria-valuenow="<%= progressPercent %>" aria-valuemin="0" aria-valuemax="100"></div>
                        </div>
                        <div class="progress-info mt-2">
                            <span class="text-muted">진행률: <%= Int(progressPercent) %>% (<%= completedSteps %>/<%= totalSteps %> 단계 완료)</span>
                        </div>
                    </div>
                </div>

                <!-- 결재 단계별 상세 정보 -->
                <div class="approval-steps">
                    <% 
                    approvalRS.MoveFirst
                    Do While Not approvalRS.EOF
                    %>
                    <div class="approval-step <%= IIf(approvalRS("status") = "승인", "completed", IIf(approvalRS("status") = "반려", "rejected", "pending")) %>">
                        <div class="step-header">
                            <div class="step-label">
                                <% If approvalRS("status") = "승인" Then %>
                                    <i class="fas fa-check-circle text-success me-1"></i>
                                <% ElseIf approvalRS("status") = "반려" Then %>
                                    <i class="fas fa-times-circle text-danger me-1"></i>
                                <% Else %>
                                    <i class="fas fa-clock text-warning me-1"></i>
                                <% End If %>
                                <%= approvalRS("approval_step") %>단계 결재
                            </div>
                            <div class="step-status">
                                <%= GetStatusBadge(approvalRS("status")) %>
                            </div>
                        </div>
                        
                        <div class="approver-details">
                            <div class="approver-main-info">
                                <div class="approver-avatar">
                                    <i class="fas fa-user-circle"></i>
                                </div>
                                <div class="approver-info">
                                    <div class="approver-name"><%= approvalRS("approver_name") %></div>
                                    <div class="approver-dept">
                                        <%= IIf(IsNull(approvalRS("department_name")), "", approvalRS("department_name")) %>
                                        <%= IIf(IsNull(approvalRS("job_grade_name")), "", " / " & approvalRS("job_grade_name")) %>
                                    </div>
                                </div>
                            </div>
                            
                            <div class="approval-timeline">
                                <div class="timeline-item">
                                    <div class="timeline-label">결재 요청</div>
                                    <div class="timeline-value"><%= FormatDateTimeValue(approvalRS("created_at")) %></div>
                                </div>
                                <% If Not IsNull(approvalRS("approved_at")) Then %>
                                <div class="timeline-item">
                                    <div class="timeline-label">결재 완료</div>
                                    <div class="timeline-value"><%= FormatDateTimeValue(approvalRS("approved_at")) %></div>
                                </div>
                                <% End If %>
                            </div>
                        </div>
                        
                        <% If Not IsNull(approvalRS("comments")) And approvalRS("comments") <> "" Then %>
                        <div class="approval-comment">
                            <div class="comment-header">
                                <i class="fas fa-comment-alt me-1"></i>결재 의견
                            </div>
                            <div class="comment-content">
                                <%= approvalRS("comments") %>
                            </div>
                        </div>
                        <% End If %>
                        
                        <!-- 결재 처리 시간 계산 -->
                        <% If Not IsNull(approvalRS("approved_at")) And Not IsNull(approvalRS("created_at")) Then %>
                        <div class="processing-time">
                            <%
                            Dim processingHours
                            processingHours = DateDiff("h", approvalRS("created_at"), approvalRS("approved_at"))
                            %>
                            <small class="text-muted">
                                <i class="fas fa-stopwatch me-1"></i>처리 시간: 
                                <% If processingHours < 24 Then %>
                                    <%= processingHours %>시간
                                <% Else %>
                                    <%= Int(processingHours / 24) %>일 <%= processingHours Mod 24 %>시간
                                <% End If %>
                            </small>
                        </div>
                        <% End If %>
                    </div>
                    <%
                    approvalRS.MoveNext
                    Loop
                    %>
                </div>
                
                <!-- 결재 요약 정보 -->
                <div class="approval-summary mt-4">
                    <div class="summary-header">
                        <h6><i class="fas fa-info-circle me-2"></i>결재 요약</h6>
                    </div>
                    <div class="summary-content">
                        <div class="row">
                            <div class="col-md-3">
                                <div class="summary-item">
                                    <div class="summary-label">총 결재 단계</div>
                                    <div class="summary-value"><%= totalSteps %>단계</div>
                                </div>
                            </div>
                            <div class="col-md-3">
                                <div class="summary-item">
                                    <div class="summary-label">완료된 단계</div>
                                    <div class="summary-value text-success"><%= completedSteps %>단계</div>
                                </div>
                            </div>
                            <div class="col-md-3">
                                <div class="summary-item">
                                    <div class="summary-label">현재 단계</div>
                                    <div class="summary-value text-primary">
                                        <% If currentStep > 0 Then %>
                                            <%= currentStep %>단계
                                        <% Else %>
                                            완료
                                        <% End If %>
                                    </div>
                                </div>
                            </div>
                            <div class="col-md-3">
                                <div class="summary-item">
                                    <div class="summary-label">전체 진행률</div>
                                    <div class="summary-value text-info"><%= Int(progressPercent) %>%</div>
                                </div>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
        </div>
    </div>
    <% End If %>
</div>

<% If Not docRS Is Nothing Then If docRS.State = 1 Then docRS.Close : Set docRS = Nothing %>
<% If Not approvalRS Is Nothing Then If approvalRS.State = 1 Then approvalRS.Close : Set approvalRS = Nothing %>
<!--#include file="../../includes/footer.asp"-->
