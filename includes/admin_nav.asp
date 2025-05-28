<%
' 현재 페이지 파일명 추출
Dim currentPage
currentPage = LCase(Mid(Request.ServerVariables("SCRIPT_NAME"), InStrRev(Request.ServerVariables("SCRIPT_NAME"), "/") + 1))

Function GetActiveClass(pageName)
    If InStr(currentPage, pageName) > 0 Then
        GetActiveClass = " active"
    Else
        GetActiveClass = ""
    End If
End Function
%>

<link rel="stylesheet" href="../../includes/admin_styles.css">

<div class="admin-container">
    <!-- 관리자 네비게이션 -->
    <div class="admin-nav">
        <div class="admin-nav-title">
            <i class="fas fa-cog me-2"></i>관리자 메뉴
        </div>
        <div class="admin-nav-grid">
            <a href="admin_dashboard.asp" class="admin-nav-item<%= GetActiveClass("dashboard") %>">
                <i class="fas fa-tachometer-alt"></i>대시보드
            </a>
            <a href="admin_cardaccount.asp" class="admin-nav-item<%= GetActiveClass("cardaccount") %>">
                <i class="fas fa-credit-card"></i>카드 계정 관리
            </a>
            <a href="admin_cardaccounttypes.asp" class="admin-nav-item<%= GetActiveClass("cardaccounttypes") %>">
                <i class="fas fa-tags"></i>카드 계정 유형 관리
            </a>
            <a href="admin_fuelrate.asp" class="admin-nav-item<%= GetActiveClass("fuelrate") %>">
                <i class="fas fa-gas-pump"></i>유류비 단가 관리
            </a>
            <a href="admin_job_grade.asp" class="admin-nav-item<%= GetActiveClass("job_grade") %>">
                <i class="fas fa-user-tie"></i>직급 관리
            </a>
            <a href="admin_department.asp" class="admin-nav-item<%= GetActiveClass("department") %>">
                <i class="fas fa-sitemap"></i>부서 관리
            </a>
            <a href="admin_users.asp" class="admin-nav-item<%= GetActiveClass("users") %>">
                <i class="fas fa-users"></i>사용자 관리
            </a>
            <a href="admin_card_usage.asp" class="admin-nav-item<%= GetActiveClass("card_usage") %>">
                <i class="fas fa-receipt"></i>카드 사용 내역 관리
            </a>
            <a href="admin_vehicle_requests.asp" class="admin-nav-item<%= GetActiveClass("vehicle_requests") %>">
                <i class="fas fa-car"></i>차량 이용 신청 관리
            </a>
            <a href="admin_approvals.asp" class="admin-nav-item<%= GetActiveClass("approvals") %>">
                <i class="fas fa-file-signature"></i>결재 로그 관리
            </a>
        </div>
    </div>
</div> 