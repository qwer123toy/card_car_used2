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

' 관리자 권한 체크
If Not IsAdmin() Then
    Response.Write("<script>alert('관리자 권한이 필요합니다.'); window.location.href='../dashboard.asp';</script>")
    Response.End
End If

' 사용자 ID 확인
Dim userId
userId = Request.QueryString("id")

If userId = "" Then
    Response.Write("<script>alert('사용자 ID가 필요합니다.'); window.location.href='admin_users.asp';</script>")
    Response.End
End If

' 사용자 정보 조회
Dim userSQL, userRS
userSQL = "SELECT u.user_id, u.name, u.email, u.department_id, " & _
          "u.job_grade, u.is_active, u.created_at,  " & _
          "d.name AS department_name, jg.name AS job_grade_name " & _
          "FROM " & dbSchema & ".Users u " & _
          "LEFT JOIN " & dbSchema & ".Department d ON u.department_id = d.department_id " & _
          "LEFT JOIN " & dbSchema & ".Job_Grade jg ON u.job_grade = jg.job_grade_id " & _
          "WHERE u.user_id = '" & Replace(userId, "'", "''") & "'"


Set userRS = db99.Execute(userSQL)

If userRS.EOF Then
    Response.Write("<script>alert('사용자를 찾을 수 없습니다.'); window.location.href='admin_users.asp';</script>")
    Response.End
End If

' 사용자의 카드 사용 내역 조회 (최근 10건)
Dim cardUsageSQL, cardUsageRS
cardUsageSQL = "SELECT TOP 10 cu.usage_id, cu.usage_date, cu.store_name, cu.amount, " & _
               "cu.purpose, cu.approval_status, ca.issuer " & _
               "FROM " & dbSchema & ".CardUsage cu " & _
               "LEFT JOIN " & dbSchema & ".CardAccount ca ON cu.card_id = ca.card_id " & _
               "WHERE cu.user_id = '" & Replace(userId, "'", "''") & "' " & _
               "ORDER BY cu.usage_date DESC"


Set cardUsageRS = db99.Execute(cardUsageSQL)

' 사용자의 차량 이용 신청 조회 (최근 10건)
Dim vehicleRequestSQL, vehicleRequestRS
vehicleRequestSQL = "SELECT TOP 10 vr.request_id, vr.request_date, vr.start_date, vr.end_date, " & _
                    "vr.destination, vr.purpose, vr.approval_status " & _
                    "FROM " & dbSchema & ".VehicleRequests vr " & _
                    "WHERE vr.user_id = '" & Replace(userId, "'", "''") & "' " & _
                    "ORDER BY vr.request_date DESC"

Set vehicleRequestRS = db99.Execute(vehicleRequestSQL)

' 승인 상태 표시
Function GetApprovalStatusBadge(status)
    Select Case status
        Case "승인"
            GetApprovalStatusBadge = "<span class=""badge bg-success"">승인</span>"
        Case "대기"
            GetApprovalStatusBadge = "<span class=""badge bg-warning"">대기</span>"
        Case "반려"
            GetApprovalStatusBadge = "<span class=""badge bg-danger"">반려</span>"
    End Select
End Function


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

.info-section {
    background: white;
    border-radius: 16px;
    padding: 2rem;
    margin-bottom: 2rem;
    box-shadow: 0 4px 20px rgba(0,0,0,0.08);
}

.info-title {
    font-size: 1.1rem;
    font-weight: 600;
    color: #2C3E50;
    margin-bottom: 1.5rem;
    display: flex;
    align-items: center;
}

.info-grid {
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(300px, 1fr));
    gap: 1.5rem;
}

.info-item {
    display: flex;
    flex-direction: column;
}

.info-label {
    font-size: 0.875rem;
    font-weight: 600;
    color: #64748B;
    margin-bottom: 0.5rem;
}

.info-value {
    font-size: 1rem;
    color: #2C3E50;
    font-weight: 500;
}

.table-section {
    background: white;
    border-radius: 16px;
    padding: 2rem;
    box-shadow: 0 4px 20px rgba(0,0,0,0.08);
    margin-bottom: 2rem;
}

.table-title {
    font-size: 1.1rem;
    font-weight: 600;
    color: #2C3E50;
    margin-bottom: 1.5rem;
    display: flex;
    align-items: center;
}

.table {
    margin-bottom: 0;
    border-radius: 12px;
    overflow: hidden;
    box-shadow: 0 2px 8px rgba(0,0,0,0.05);
}

.table th {
    background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
    color: white;
    font-weight: 600;
    border: none;
    padding: 1rem;
    font-size: 0.95rem;
}

.table td {
    padding: 1rem;
    vertical-align: middle;
    border-bottom: 1px solid #E9ECEF;
    color: #2C3E50;
}

.table tbody tr:hover {
    background-color: #F8FAFC;
    transition: background-color 0.2s ease;
}

.btn {
    padding: 0.875rem 1.5rem;
    font-weight: 600;
    border-radius: 8px;
    transition: all 0.2s ease;
    margin: 0 0.25rem;
}

.btn-primary {
    background: linear-gradient(to right, #4A90E2, #5A9EEA);
    border: none;
    color: white;
}

.btn-primary:hover {
    transform: translateY(-2px);
    box-shadow: 0 4px 12px rgba(74,144,226,0.2);
}

.btn-secondary {
    background: linear-gradient(to right, #6C757D, #5A6268);
    border: none;
    color: white;
}

.btn-secondary:hover {
    transform: translateY(-2px);
    box-shadow: 0 4px 12px rgba(108,117,125,0.2);
}

.empty-state {
    text-align: center;
    padding: 3rem;
    color: #64748B;
}

.empty-state i {
    font-size: 3rem;
    margin-bottom: 1rem;
    color: #CBD5E1;
}

.badge {
    font-size: 0.75rem;
    padding: 0.5rem 0.75rem;
    border-radius: 6px;
}

.status-badge {
    display: inline-flex;
    align-items: center;
    padding: 0.5rem 1rem;
    border-radius: 8px;
    font-size: 0.875rem;
    font-weight: 600;
}

.status-active {
    background: #D1FAE5;
    color: #065F46;
}

.status-inactive {
    background: #FEE2E2;
    color: #991B1B;
}

.status-admin {
    background: #DBEAFE;
    color: #1E40AF;
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
            <a href="admin_users.asp" class="admin-nav-item active">
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

    <!-- 페이지 헤더 -->
    <div class="page-header">
        <h2 class="page-title">
            <i class="fas fa-user me-2"></i>사용자 상세 정보
        </h2>
        <div>
            <a href="admin_users.asp" class="btn btn-secondary">
                <i class="fas fa-arrow-left me-1"></i> 목록으로
            </a>
        </div>
    </div>

    <!-- 사용자 기본 정보 -->
    <div class="info-section">
        <div class="info-title">
            <i class="fas fa-user me-2"></i>기본 정보
        </div>
        <div class="info-grid">
            <div class="info-item">
                <div class="info-label">사용자 ID</div>
                <div class="info-value"><%= userRS("user_id") %></div>
            </div>

            <div class="info-item">
                <div class="info-label">이름</div>
                <div class="info-value"><%= userRS("name") %></div>
            </div>
                        
            <div class="info-item">
                <div class="info-label">부서</div>
                <div class="info-value"><%= IIf(IsNull(userRS("department_name")), "-", userRS("department_name")) %></div>
            </div>
            <div class="info-item">
                <div class="info-label">직급</div>
                <div class="info-value"><%= IIf(IsNull(userRS("job_grade_name")), "-", userRS("job_grade_name")) %></div>
            </div>
            <div class="info-item">
                <div class="info-label">상태</div>
                <div class="info-value">
                    <% If userRS("is_active") Then %>
                    <span class="status-badge status-active">활성</span>
                    <% Else %>
                    <span class="status-badge status-inactive">비활성</span>
                    <% End If %>

                </div>
            </div>
            <div class="info-item">
                <div class="info-label">가입일</div>
                <div class="info-value"><%= FormatDateTime(userRS("created_at"), 2) %></div>
            </div>

        </div>
    </div>

    <!-- 카드 사용 내역 -->
    <div class="table-section">
        <div class="table-title">
            <i class="fas fa-credit-card me-2"></i>최근 카드 사용 내역
        </div>
        
        <% If cardUsageRS.EOF Then %>
        <div class="empty-state">
            <i class="fas fa-credit-card"></i>
            <h5>카드 사용 내역이 없습니다</h5>
        </div>
        <% Else %>
        <div class="table-responsive">
            <table class="table">
                <thead>
                    <tr>
                        <th style="text-align: center;">사용일</th>
                        <th style="text-align: center;">카드</th>
                        <th style="text-align: center;">가맹점</th>
                        <th style="text-align: center;">금액</th>
                        <th style="text-align: center;">목적</th>
                        <th style="text-align: center;">상태</th>
                    </tr>
                </thead>
                <tbody>
                    <% Do While Not cardUsageRS.EOF %>
                    <tr>
                        <td style="text-align: center;"><%= FormatDateTime(cardUsageRS("usage_date"), 2) %></td>
                        <td style="text-align: center;"><%= cardUsageRS("issuer") %></td>
                        <td style="text-align: center;"><%= cardUsageRS("store_name") %></td>
                        <td style="text-align: center;"><strong><%= FormatCurrency(cardUsageRS("amount")) %></strong></td>
                        <td style="text-align: center;"><%= cardUsageRS("purpose") %></td>
                        <td style="text-align: center;"><%= GetApprovalStatusBadge(cardUsageRS("approval_status")) %></td>
                    </tr>
                    <% 
                    cardUsageRS.MoveNext
                    Loop
                    %>
                </tbody>
            </table>
        </div>
        <% End If %>
    </div>

    <!-- 차량 이용 신청 -->
    <div class="table-section">
        <div class="table-title">
            <i class="fas fa-car me-2"></i>최근 차량 이용 신청
        </div>
        
        <% If vehicleRequestRS.EOF Then %>
        <div class="empty-state">
            <i class="fas fa-car"></i>
            <h5>차량 이용 신청이 없습니다</h5>
        </div>
        <% Else %>
        <div class="table-responsive">
            <table class="table">
                <thead>
                    <tr>
                        <th style="text-align: center;">출발일</th>
                        <th style="text-align: center;">반납일</th>
                        <th style="text-align: center;">목적지</th>
                        <th style="text-align: center;">목적</th>
                        <th style="text-align: center;">상태</th>
                    </tr>
                </thead>
                <tbody>
                    <% Do While Not vehicleRequestRS.EOF %>
                    <tr>
                        <td style="text-align: center;"><%= FormatDateTime(vehicleRequestRS("start_date"), 2) %></td>
                        <td style="text-align: center;"><%= FormatDateTime(vehicleRequestRS("end_date"), 2) %></td>
                        <td style="text-align: center;"><%= vehicleRequestRS("destination") %></td>
                        <td style="text-align: center;"><%= vehicleRequestRS("purpose") %></td>
                        <td style="text-align: center;"><%= GetApprovalStatusBadge(vehicleRequestRS("approval_status")) %></td>
                    </tr>
                    <% 
                    vehicleRequestRS.MoveNext
                    Loop
                    %>
                </tbody>
            </table>
        </div>
        <% End If %>
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

If Not cardUsageRS Is Nothing Then
    If cardUsageRS.State = 1 Then
        cardUsageRS.Close
    End If
    Set cardUsageRS = Nothing
End If

If Not vehicleRequestRS Is Nothing Then
    If vehicleRequestRS.State = 1 Then
        vehicleRequestRS.Close
    End If
    Set vehicleRequestRS = Nothing
End If
%>

<!--#include file="../../includes/footer.asp"--> 