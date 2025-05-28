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

' 결재 문서 삭제 처리
If Request.QueryString("action") = "delete" And Request.QueryString("id") <> "" Then
    Dim deleteId
    deleteId = PreventSQLInjection(Request.QueryString("id"))
    
    ' 삭제 쿼리 실행
    Dim deleteSQL
    deleteSQL = "DELETE FROM " & dbSchema & ".Approvals WHERE approval_id = " & deleteId
    
    On Error Resume Next
    db.Execute(deleteSQL)
    
    If Err.Number <> 0 Then
        Response.Write("<script>alert('결재 문서 삭제 중 오류가 발생했습니다: " & Server.HTMLEncode(Err.Description) & "'); window.location.href='admin_approvals.asp';</script>")
    Else
        ' 활동 로그 기록
        LogActivity Session("user_id"), "결재문서삭제", "결재 문서 삭제 (ID: " & deleteId & ")"
        Response.Write("<script>alert('결재 문서가 삭제되었습니다.'); window.location.href='admin_approvals.asp';</script>")
    End If
    On Error GoTo 0
    Response.End
End If

' 페이징 처리
Dim pageNo, pageSize, totalCount, totalPages
pageSize = 15 ' 페이지당 표시할 레코드 수

' 현재 페이지 번호
If Request.QueryString("page") = "" Then
    pageNo = 1
Else
    pageNo = CInt(Request.QueryString("page"))
End If

' 검색 조건에 따른 SQL 쿼리 구성
Dim searchKeyword, searchField, searchDateFrom, searchDateTo, searchStatus, whereClause
searchKeyword = Trim(Request.QueryString("keyword"))
searchField = Request.QueryString("field")
searchDateFrom = Request.QueryString("date_from")
searchDateTo = Request.QueryString("date_to")
searchStatus = Request.QueryString("status")

whereClause = ""
Dim whereConditions : whereConditions = Array()
Dim conditionIndex : conditionIndex = 0

' 키워드 검색 조건
If searchKeyword <> "" Then
    If searchField = "requester_id" Then
        ReDim Preserve whereConditions(conditionIndex)
        whereConditions(conditionIndex) = "u1.name LIKE '%" & PreventSQLInjection(searchKeyword) & "%'"
        conditionIndex = conditionIndex + 1
    ElseIf searchField = "approver_id" Then
        ReDim Preserve whereConditions(conditionIndex)
        whereConditions(conditionIndex) = "u2.name LIKE '%" & PreventSQLInjection(searchKeyword) & "%'"
        conditionIndex = conditionIndex + 1
    ElseIf searchField = "title" Then
        ReDim Preserve whereConditions(conditionIndex)
        whereConditions(conditionIndex) = "a.title LIKE '%" & PreventSQLInjection(searchKeyword) & "%'"
        conditionIndex = conditionIndex + 1
    End If
End If

' 날짜 범위 검색 조건
If IsDate(searchDateFrom) Then
    ReDim Preserve whereConditions(conditionIndex)
    whereConditions(conditionIndex) = "a.request_date >= '" & CDate(searchDateFrom) & "'"
    conditionIndex = conditionIndex + 1
End If

If IsDate(searchDateTo) Then
    ReDim Preserve whereConditions(conditionIndex)
    whereConditions(conditionIndex) = "a.request_date <= '" & CDate(searchDateTo) & " 23:59:59'"
    conditionIndex = conditionIndex + 1
End If

' 상태 검색 조건
If searchStatus <> "" Then
    ReDim Preserve whereConditions(conditionIndex)
    whereConditions(conditionIndex) = "a.status = '" & PreventSQLInjection(searchStatus) & "'"
    conditionIndex = conditionIndex + 1
End If

' WHERE 절 구성
If conditionIndex > 0 Then
    whereClause = " WHERE " & Join(whereConditions, " AND ")
End If

' 전체 레코드 수
Dim countSQL, countRS
countSQL = "SELECT COUNT(*) AS cnt " & _
           "FROM " & dbSchema & ".Approvallogs a " & _
           "LEFT JOIN " & dbSchema & ".Users u1 ON a.approver_id = u1.user_id " & _
           "LEFT JOIN " & dbSchema & ".Users u2 ON a.approver_id = u2.user_id " & _
           IIf(whereClause <> "", whereClause, "")

Set countRS = db99.Execute(countSQL)
totalCount = countRS("cnt")
totalPages = (totalCount + pageSize - 1) \ pageSize

' 결재 문서 목록 조회
Dim listSQL, listRS
' WHERE 절이 있을 경우 "AND"로 연결 (앞의 WHERE 제거하고)
listSQL = "SELECT * FROM (" & _
          "SELECT TOP " & pageSize & " * FROM (" & _
          "SELECT TOP " & (pageNo * pageSize) & " a.approval_log_id, a.approver_id, " & _
          "a.target_table_name, a.target_id, a.approval_step, a.status, a.approved_at, " & _
          "a.comments, a.created_at, u1.name AS approver_name, " & _
          "ISNULL(cu.title, vr.title) AS title " & _
          "FROM " & dbSchema & ".Approvallogs a " & _
          "INNER JOIN (" & _
          "    SELECT MAX(approval_log_id) AS approval_log_id " & _
          "    FROM " & dbSchema & ".Approvallogs " & _
          "    GROUP BY target_id" & _
          ") AS grouped ON a.approval_log_id = grouped.approval_log_id " & _
          "LEFT JOIN " & dbSchema & ".Users u1 ON a.approver_id = u1.user_id " & _
          "LEFT JOIN " & dbSchema & ".CardUsage cu ON a.target_table_name = 'CardUsage' AND a.target_id = cu.usage_id " & _
          "LEFT JOIN " & dbSchema & ".VehicleRequests vr ON a.target_table_name = 'VehicleRequests' AND a.target_id = vr.request_id " & _
          IIf(whereClause <> "", whereClause, "") & " " & _
          "ORDER BY a.created_at DESC) AS T1 " & _
          "ORDER BY created_at ASC) AS T2 " & _
          "ORDER BY created_at DESC"

Set listRS = db99.Execute(listSQL)

' 상태 옵션
Dim statusOptions : statusOptions = Array("대기", "승인", "반려")

' 사용자 목록 조회
Dim userSQL, userRS
userSQL = "SELECT user_id, name FROM " & dbSchema & ".Users WHERE is_active = 1 ORDER BY name"
Set userRS = db99.Execute(userSQL)

' 날짜 포맷
Function FormatDate(dateValue)
    If IsNull(dateValue) Or Not IsDate(dateValue) Then
        FormatDate = "-"
    Else
        FormatDate = FormatDateTime(dateValue, 2)
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

.search-section {
    background: white;
    border-radius: 16px;
    padding: 2rem;
    margin-bottom: 2rem;
    box-shadow: 0 4px 20px rgba(0,0,0,0.08);
}

.search-title {
    font-size: 1.1rem;
    font-weight: 600;
    color: #2C3E50;
    margin-bottom: 1.5rem;
    display: flex;
    align-items: center;
}

.form-control, .form-select {
    border-radius: 8px;
    border: 2px solid #E9ECEF;
    padding: 0.875rem 1rem;
    font-size: 1rem;
    transition: all 0.2s ease;
}

.form-control:focus, .form-select:focus {
    border-color: #4A90E2;
    box-shadow: 0 0 0 4px rgba(74,144,226,0.1);
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

.btn-danger {
    background: linear-gradient(to right, #E74C3C, #C0392B);
    border: none;
    color: white;
}

.btn-danger:hover {
    transform: translateY(-2px);
    box-shadow: 0 4px 12px rgba(231,76,60,0.2);
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

.btn-sm {
    padding: 0.5rem 1rem;
    font-size: 0.875rem;
    margin: 0 0.125rem;
}

.pagination {
    margin-top: 2rem;
}

.page-link {
    border-radius: 8px;
    border: 2px solid #E9ECEF;
    color: #4A90E2;
    padding: 0.75rem 1rem;
    margin: 0 0.125rem;
    font-weight: 500;
}

.page-link:hover {
    background-color: #4A90E2;
    border-color: #4A90E2;
    color: white;
}

.page-item.active .page-link {
    background-color: #4A90E2;
    border-color: #4A90E2;
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
            <i class="fas fa-file-signature me-2"></i>결재 로그 관리
        </h2>
    </div>

    <!-- 검색 섹션 -->
    <div class="search-section">
        <div class="search-title">
            <i class="fas fa-search me-2"></i>결재 로그 검색
        </div>
        <form action="admin_approvals.asp" method="get">
            <div class="row g-3">
                <div class="col-md-2">
                    <label class="form-label">검색 필드</label>
                    <select class="form-select" name="field">
                        <option value="title" <% If searchField = "title" Then %>selected<% End If %>>제목</option>
                        <option value="requester_id" <% If searchField = "requester_id" Then %>selected<% End If %>>결재자</option>
                        
                    </select>
                </div>
                <div class="col-md-3">
                    <label class="form-label">검색어</label>
                    <input type="text" class="form-control" name="keyword" value="<%= searchKeyword %>" placeholder="검색어를 입력하세요">
                </div>
                <div class="col-md-2">
                    <label class="form-label">상태</label>
                    <select class="form-select" name="status">
                        <option value="">전체</option>
                        <% For Each status In statusOptions %>
                        <option value="<%= status %>" <% If searchStatus = status Then %>selected<% End If %>><%= status %></option>
                        <% Next %>
                    </select>
                </div>
                <div class="col-md-2">
                    <label class="form-label">시작일</label>
                    <input type="date" class="form-control" name="date_from" value="<%= searchDateFrom %>">
                </div>
                <div class="col-md-2">
                    <label class="form-label">종료일</label>
                    <input type="date" class="form-control" name="date_to" value="<%= searchDateTo %>">
                </div>
                <div class="col-md-1">
                    <label class="form-label">&nbsp;</label>
                    <button type="submit" class="btn btn-primary w-100">
                        <i class="fas fa-search"></i>
                    </button>
                </div>
            </div>
        </form>
    </div>

    <!-- 결재 로그 목록 -->
    <div class="table-section">
        <div class="table-title">
            <i class="fas fa-list me-2"></i>결재 로그 목록 (총 <%= totalCount %>건)
        </div>
        
        <% If listRS.EOF Then %>
        <div class="empty-state">
            <i class="fas fa-file-signature"></i>
            <h5>결재 로그가 없습니다</h5>
            <p>검색 조건을 변경해보세요.</p>
        </div>
        <% Else %>
        <div class="table-responsive">
            <table class="table">
                <thead>
                    <tr>
                        <th style="text-align: center;">문서 유형</th>
                        <th style="text-align: center;">제목</th>
                        <th style="text-align: center;">결재자</th>
                        <th style="text-align: center;">결재 단계</th>
                        <th style="text-align: center;">상태</th>
                        <th style="text-align: center;">결재일</th>
                        <th style="text-align: center;">관리</th>
                    </tr>
                </thead>
                <tbody>
                    <% Do While Not listRS.EOF %>
                    <tr>
                        <td style="text-align: center;">
                            <% If listRS("target_table_name") = "CardUsage" Then %>
                                <i class="fas fa-credit-card me-1"></i>카드 사용
                            <% ElseIf listRS("target_table_name") = "VehicleRequests" Then %>
                                <i class="fas fa-car me-1"></i>차량 신청
                            <% Else %>
                                <i class="fas fa-file me-1"></i><%= listRS("target_table_name") %>
                            <% End If %>
                        </td>
                        <td style="text-align: center;"><%= IIf(IsNull(listRS("title")), "-", listRS("title")) %></td>
                        <td style="text-align: center;"><%= IIf(IsNull(listRS("approver_name")), "-", listRS("approver_name")) %></td>
                        <td style="text-align: center;"><%= listRS("approval_step") %>단계</td>
                        <td style="text-align: center;"><%= GetStatusBadge(listRS("status")) %></td>
                        <td style="text-align: center;"><%= FormatDate(listRS("approved_at")) %></td>
                        <td style="text-align: center;">
                            <a href="admin_approval_view.asp?target_id=<%= listRS("target_id") %>&target_table_name=<%= listRS("target_table_name") %>" class="btn btn-sm btn-primary">
                                <i class="fas fa-eye"></i> 상세
                            </a>
                        </td>
                    </tr>
                    <% 
                    listRS.MoveNext
                    Loop
                    %>
                </tbody>
            </table>
        </div>

        <!-- 페이징 -->
        <% If totalPages > 1 Then %>
        <nav aria-label="Page navigation">
            <ul class="pagination justify-content-center">
                <% If pageNo > 1 Then %>
                <li class="page-item">
                    <a class="page-link" href="admin_approvals.asp?page=<%= pageNo - 1 %>&field=<%= searchField %>&keyword=<%= searchKeyword %>&status=<%= searchStatus %>&date_from=<%= searchDateFrom %>&date_to=<%= searchDateTo %>">
                        <i class="fas fa-chevron-left"></i> 이전
                    </a>
                </li>
                <% End If %>
                
                <% 
                Dim startPage, endPage
                If pageNo - 5 > 1 Then
                    startPage = pageNo - 5
                Else
                    startPage = 1
                End If
                
                If pageNo + 5 < totalPages Then
                    endPage = pageNo + 5
                Else
                    endPage = totalPages
                End If
                
                For i = startPage To endPage
                %>
                <li class="page-item <% If i = pageNo Then %>active<% End If %>">
                    <a class="page-link" href="admin_approvals.asp?page=<%= i %>&field=<%= searchField %>&keyword=<%= searchKeyword %>&status=<%= searchStatus %>&date_from=<%= searchDateFrom %>&date_to=<%= searchDateTo %>"><%= i %></a>
                </li>
                <% Next %>
                
                <% If pageNo < totalPages Then %>
                <li class="page-item">
                    <a class="page-link" href="admin_approvals.asp?page=<%= pageNo + 1 %>&field=<%= searchField %>&keyword=<%= searchKeyword %>&status=<%= searchStatus %>&date_from=<%= searchDateFrom %>&date_to=<%= searchDateTo %>">
                        다음 <i class="fas fa-chevron-right"></i>
                    </a>
                </li>
                <% End If %>
            </ul>
        </nav>
        <% End If %>
        <% End If %>
    </div>
</div>

<%
' 사용한 객체 해제
If Not listRS Is Nothing Then
    If listRS.State = 1 Then
        listRS.Close
    End If
    Set listRS = Nothing
End If

If Not userRS Is Nothing Then
    If userRS.State = 1 Then
        userRS.Close
    End If
    Set userRS = Nothing
End If
%>

<!--#include file="../../includes/footer.asp"--> 