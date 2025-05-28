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

' 사용 내역 ID 확인
Dim usageId
usageId = Request.QueryString("id")

If usageId = "" Then
    Response.Write("<script>alert('사용 내역 ID가 필요합니다.'); window.location.href='admin_card_usage.asp';</script>")
    Response.End
End If

' 카드 사용 내역 상세 정보 조회
Dim usageSQL, usageRS
usageSQL = "SELECT cu.usage_id, cu.user_id, cu.title, cu.usage_date, cu.store_name, " & _
           "cu.amount, cu.purpose, cu.receipt_file, cu.approval_status, cu.created_at, " & _
           "u.name AS user_name, ca.account_name, ca.issuer, " & _
           "cat.type_name AS category_name " & _
           "FROM " & dbSchema & ".CardUsage cu " & _
           "LEFT JOIN " & dbSchema & ".Users u ON cu.user_id = u.user_id " & _
           "LEFT JOIN " & dbSchema & ".CardAccount ca ON cu.card_id = ca.card_id " & _
           "LEFT JOIN " & dbSchema & ".CardAccountTypes cat ON cu.expense_category_id = cat.account_type_id " & _
           "WHERE cu.usage_id = " & usageId

Set usageRS = db99.Execute(usageSQL)

If usageRS.EOF Then
    Response.Write("<script>alert('카드 사용 내역을 찾을 수 없습니다.'); window.location.href='admin_card_usage.asp';</script>")
    Response.End
End If

' 승인 상태 표시
Function GetApprovalStatusBadge(status)
    Select Case status
        Case "승인"
            GetApprovalStatusBadge = "<span class=""badge bg-success"">승인</span>"
        Case "대기"
            GetApprovalStatusBadge = "<span class=""badge bg-warning"">대기</span>"
        Case "반려"
            GetApprovalStatusBadge = "<span class=""badge bg-danger"">반려</span>"
        Case Else
            GetApprovalStatusBadge = "<span class=""badge bg-secondary"">미정</span>"
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

.detail-card {
    background: white;
    border-radius: 16px;
    padding: 2rem;
    margin-bottom: 2rem;
    box-shadow: 0 4px 20px rgba(0,0,0,0.08);
}

.detail-title {
    font-size: 1.1rem;
    font-weight: 600;
    color: #2C3E50;
    margin-bottom: 1.5rem;
    display: flex;
    align-items: center;
}

.detail-grid {
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(300px, 1fr));
    gap: 1.5rem;
}

.detail-item {
    display: flex;
    flex-direction: column;
}

.detail-label {
    font-size: 0.875rem;
    font-weight: 600;
    color: #64748B;
    margin-bottom: 0.5rem;
}

.detail-value {
    font-size: 1rem;
    color: #2C3E50;
    font-weight: 500;
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

.badge {
    font-size: 0.75rem;
    padding: 0.5rem 0.75rem;
    border-radius: 6px;
}

.receipt-section {
    background: #F8FAFC;
    border-radius: 12px;
    padding: 1.5rem;
    margin-top: 1rem;
}

.receipt-title {
    font-size: 1rem;
    font-weight: 600;
    color: #2C3E50;
    margin-bottom: 1rem;
}

.receipt-image {
    max-width: 100%;
    height: auto;
    border-radius: 8px;
    box-shadow: 0 2px 8px rgba(0,0,0,0.1);
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
            <a href="admin_card_usage.asp" class="admin-nav-item active">
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
            <i class="fas fa-receipt me-2"></i>카드 사용 내역 상세
        </h2>
        <div>
            <a href="admin_card_usage.asp" class="btn btn-secondary">
                <i class="fas fa-arrow-left me-1"></i> 목록으로
            </a>
        </div>
    </div>

    <!-- 카드 사용 내역 상세 정보 -->
    <div class="detail-card">
        <div class="detail-title">
            <i class="fas fa-info-circle me-2"></i>사용 내역 정보
        </div>
        <div class="detail-grid">
            <div class="detail-item">
                <div class="detail-label">사용 내역 ID</div>
                <div class="detail-value"><%= usageRS("usage_id") %></div>
            </div>
            <div class="detail-item">
                <div class="detail-label">제목</div>
                <div class="detail-value"><%= usageRS("title") %></div>
            </div>
            <div class="detail-item">
                <div class="detail-label">사용자</div>
                <div class="detail-value"><%= usageRS("user_name") %></div>
            </div>
            <div class="detail-item">
                <div class="detail-label">사용일</div>
                <div class="detail-value"><%= FormatDateTime(usageRS("usage_date"), 2) %></div>
            </div>
            <div class="detail-item">
                <div class="detail-label">카드</div>
                <div class="detail-value">
                    <%= usageRS("account_name") %><br>
                    <small class="text-muted"><%= usageRS("issuer") %></small>
                </div>
            </div>
            <div class="detail-item">
                <div class="detail-label">가맹점</div>
                <div class="detail-value"><%= usageRS("store_name") %></div>
            </div>
            <div class="detail-item">
                <div class="detail-label">사용 금액</div>
                <div class="detail-value"><strong><%= FormatCurrency(usageRS("amount")) %></strong></div>
            </div>
            <div class="detail-item">
                <div class="detail-label">지출 카테고리</div>
                <div class="detail-value"><%= IIf(IsNull(usageRS("category_name")), "-", usageRS("category_name")) %></div>
            </div>
            <div class="detail-item">
                <div class="detail-label">사용 목적</div>
                <div class="detail-value"><%= usageRS("purpose") %></div>
            </div>
            <div class="detail-item">
                <div class="detail-label">승인 상태</div>
                <div class="detail-value"><%= GetApprovalStatusBadge(usageRS("approval_status")) %></div>
            </div>
            <div class="detail-item">
                <div class="detail-label">등록일</div>
                <div class="detail-value"><%= FormatDateTime(usageRS("created_at"), 0) %></div>
            </div>
        </div>

        <!-- 영수증 이미지 -->
        <% If Not IsNull(usageRS("receipt_file")) And usageRS("receipt_file") <> "" Then %>
        <div class="receipt-section">
            <div class="receipt-title">
                <i class="fas fa-file-image me-2"></i>영수증 이미지
            </div>
            <img src="../../uploads/receipts/<%= usageRS("receipt_file") %>" alt="영수증" class="receipt-image">
        </div>
        <% End If %>
    </div>
</div>

<%
' 사용한 객체 해제
If Not usageRS Is Nothing Then
    If usageRS.State = 1 Then
        usageRS.Close
    End If
    Set usageRS = Nothing
End If
%>

<!--#include file="../../includes/footer.asp"--> 