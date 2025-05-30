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

' 엑셀 다운로드 처리
If Request.QueryString("action") = "excel" Then
    ' 검색 조건 가져오기
    Dim excelSearchKeyword, excelSearchField, excelSearchDateFrom, excelSearchDateTo, excelWhereClause
    excelSearchKeyword = Trim(Request.QueryString("keyword"))
    excelSearchField = Request.QueryString("field")
    excelSearchDateFrom = Request.QueryString("date_from")
    excelSearchDateTo = Request.QueryString("date_to")

    excelWhereClause = ""
    Dim excelWhereConditions : excelWhereConditions = Array()
    Dim excelConditionIndex : excelConditionIndex = 0

    ' 키워드 검색 조건
    If excelSearchKeyword <> "" Then
        If excelSearchField = "card_id" Then
            ReDim Preserve excelWhereConditions(excelConditionIndex)
            excelWhereConditions(excelConditionIndex) = "ca.issuer LIKE '%" & PreventSQLInjection(excelSearchKeyword) & "%'"
            excelConditionIndex = excelConditionIndex + 1
        ElseIf excelSearchField = "user_id" Then
            ReDim Preserve excelWhereConditions(excelConditionIndex)
            excelWhereConditions(excelConditionIndex) = "u.name LIKE '%" & PreventSQLInjection(excelSearchKeyword) & "%'"
            excelConditionIndex = excelConditionIndex + 1
        ElseIf excelSearchField = "expense_category_id" Then
            ReDim Preserve excelWhereConditions(excelConditionIndex)
            excelWhereConditions(excelConditionIndex) = "cat.type_name LIKE '%" & PreventSQLInjection(excelSearchKeyword) & "%'"
            excelConditionIndex = excelConditionIndex + 1
        End If
    End If

    ' 날짜 범위 검색 조건
    If IsDate(excelSearchDateFrom) Then
        ReDim Preserve excelWhereConditions(excelConditionIndex)
        excelWhereConditions(excelConditionIndex) = "cu.usage_date >= '" & CDate(excelSearchDateFrom) & "'"
        excelConditionIndex = excelConditionIndex + 1
    End If

    If IsDate(excelSearchDateTo) Then
        ReDim Preserve excelWhereConditions(excelConditionIndex)
        excelWhereConditions(excelConditionIndex) = "cu.usage_date <= '" & CDate(excelSearchDateTo) & " 23:59:59'"
        excelConditionIndex = excelConditionIndex + 1
    End If

    ' WHERE 절 구성
    If excelConditionIndex > 0 Then
        excelWhereClause = " WHERE " & Join(excelWhereConditions, " AND ")
    End If

    ' 전체 데이터 조회 (페이징 없이)
    Dim excelSQL, excelRS
    excelSQL = "SELECT cu.usage_id, cu.user_id, cu.title, ca.account_name as card_id, ca.issuer as issuer, cu.department_id, " & _
               "cu.expense_category_id as account_type_id, cu.usage_date, cu.store_name, cu.amount, cu.purpose, " & _
               "cu.linked_table, cu.linked_id, cu.receipt_file, cu.created_at, cu.approval_status, " & _
               "u.name AS user_name, cat.type_name AS category_name " & _
               "FROM " & dbSchema & ".CardUsage cu " & _
               "LEFT JOIN " & dbSchema & ".Users u ON cu.user_id = u.user_id " & _
               "LEFT JOIN " & dbSchema & ".CardAccountTypes cat ON cu.expense_category_id = cat.account_type_id " & _
               "LEFT JOIN " & dbSchema & ".CardAccount ca ON cu.card_id = ca.card_id " & _
               IIf(excelWhereClause <> "", " " & excelWhereClause, "") & " " & _
               "ORDER BY cu.usage_date DESC"

    Set excelRS = db99.Execute(excelSQL)

    ' 엑셀 파일 헤더 설정
    Response.ContentType = "application/vnd.ms-excel"
    Response.AddHeader "Content-Disposition", "attachment; filename=card_usage_" & Replace(Replace(Replace(Now(), "/", ""), ":", ""), " ", "_") & ".xls"
    Response.CharSet = "utf-8"
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
</head>
<body>
<table border="1">
    <tr>
        <th>사용자</th>
        <th>제목</th>
        <th>카드</th>
        <th>발급사</th>
        <th>사용일</th>
        <th>사용처</th>
        <th>금액</th>
        <th>계정구분</th>
        <th>사용목적</th>
        <th>승인상태</th>
        <th>등록일</th>
    </tr>
    <% Do While Not excelRS.EOF %>
    <tr>
        <td><%= excelRS("user_name") %></td>
        <td><%= excelRS("title") %></td>
        <td><%= excelRS("card_id") %></td>
        <td><%= excelRS("issuer") %></td>
        <td><%= FormatDateTime(excelRS("usage_date"), 2) %></td>
        <td><%= excelRS("store_name") %></td>
        <td><%= FormatNumber(excelRS("amount")) %></td>
        <td><%= IIf(IsNull(excelRS("category_name")), "-", excelRS("category_name")) %></td>
        <td><%= excelRS("purpose") %></td>
        <td><%= excelRS("approval_status") %></td>
        <td><%= FormatDateTime(excelRS("created_at"), 2) %></td>
    </tr>
    <% 
    excelRS.MoveNext
    Loop
    %>
</table>
</body>
</html>
<%
    ' 사용한 객체 해제
    If Not excelRS Is Nothing Then
        If excelRS.State = 1 Then
            excelRS.Close
        End If
        Set excelRS = Nothing
    End If
    Response.End
End If

' 카드 사용 내역 삭제 처리
If Request.QueryString("action") = "delete" And Request.QueryString("id") <> "" Then
    Dim deleteId
    deleteId = PreventSQLInjection(Request.QueryString("id"))
    
    ' 삭제 쿼리 실행
    Dim deleteSQL
    deleteSQL = "DELETE FROM " & dbSchema & ".CardUsage WHERE usage_id = " & deleteId
    
    On Error Resume Next
    db.Execute(deleteSQL)
    
    If Err.Number <> 0 Then
        Response.Write("<script>alert('카드 사용 내역 삭제 중 오류가 발생했습니다.'); window.location.href='admin_card_usage.asp';</script>")
    Else
        ' 활동 로그 기록
        LogActivity Session("user_id"), "카드사용내역삭제", "카드 사용 내역 삭제 (ID: " & deleteId & ")"
        Response.Write("<script>alert('카드 사용 내역이 삭제되었습니다.'); window.location.href='admin_card_usage.asp';</script>")
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
Dim searchKeyword, searchField, searchDateFrom, searchDateTo, whereClause
searchKeyword = Trim(Request.QueryString("keyword"))
searchField = Request.QueryString("field")
searchDateFrom = Request.QueryString("date_from")
searchDateTo = Request.QueryString("date_to")

whereClause = ""
Dim whereConditions : whereConditions = Array()
Dim conditionIndex : conditionIndex = 0

' 키워드 검색 조건
If searchKeyword <> "" Then
    If searchField = "card_id" Then
        ReDim Preserve whereConditions(conditionIndex)
        whereConditions(conditionIndex) = "ca.issuer LIKE '%" & PreventSQLInjection(searchKeyword) & "%'"
        conditionIndex = conditionIndex + 1
    ElseIf searchField = "user_id" Then
        ReDim Preserve whereConditions(conditionIndex)
        whereConditions(conditionIndex) = "u.name LIKE '%" & PreventSQLInjection(searchKeyword) & "%'"
        conditionIndex = conditionIndex + 1
    ElseIf searchField = "expense_category_id" Then
        ReDim Preserve whereConditions(conditionIndex)
        whereConditions(conditionIndex) = "cat.type_name LIKE '%" & PreventSQLInjection(searchKeyword) & "%'"
        conditionIndex = conditionIndex + 1
    End If
End If

' 날짜 범위 검색 조건
If IsDate(searchDateFrom) Then
    ReDim Preserve whereConditions(conditionIndex)
    whereConditions(conditionIndex) = "cu.usage_date >= '" & CDate(searchDateFrom) & "'"
    conditionIndex = conditionIndex + 1
End If

If IsDate(searchDateTo) Then
    ReDim Preserve whereConditions(conditionIndex)
    whereConditions(conditionIndex) = "cu.usage_date <= '" & CDate(searchDateTo) & " 23:59:59'"
    conditionIndex = conditionIndex + 1
End If

' WHERE 절 구성
If conditionIndex > 0 Then
    whereClause = " WHERE " & Join(whereConditions, " AND ")
End If

' 전체 레코드 수
Dim countSQL, countRS
countSQL = "SELECT COUNT(*) AS cnt " & _
           "FROM " & dbSchema & ".CardUsage cu " & _
           "LEFT JOIN " & dbSchema & ".Users u ON cu.user_id = u.user_id " & _
           "LEFT JOIN " & dbSchema & ".CardAccount ca ON cu.card_id = ca.card_id " & _
           "LEFT JOIN " & dbSchema & ".CardAccountTypes cat ON cu.expense_category_id = cat.account_type_id " & _
           IIf(whereClause <> "", " " & whereClause, "")

Set countRS = db99.Execute(countSQL)
totalCount = countRS("cnt")
totalPages = (totalCount + pageSize - 1) \ pageSize

' 카드 사용 내역 목록 조회
Dim listSQL, listRS
listSQL = "SELECT * FROM (" & _
          "SELECT TOP " & pageSize & " * FROM (" & _
          "SELECT TOP " & (pageNo * pageSize) & " cu.usage_id, cu.user_id, cu.title, ca.account_name as card_id, ca.issuer as issuer, cu.department_id, " & _
          "cu.expense_category_id as account_type_id, cu.usage_date, cu.store_name, cu.amount, cu.purpose, " & _
          "cu.linked_table, cu.linked_id, cu.receipt_file, cu.created_at, cu.approval_status, " & _
          "u.name AS user_name " & _
          "FROM " & dbSchema & ".CardUsage cu " & _
          "LEFT JOIN " & dbSchema & ".Users u ON cu.user_id = u.user_id " & _
          "LEFT JOIN " & dbSchema & ".CardAccountTypes cat ON cu.expense_category_id = cat.account_type_id " & _
          "LEFT JOIN " & dbSchema & ".CardAccount ca ON cu.card_id = ca.card_id " & _
          IIf(whereClause <> "", " " & whereClause, "") & " " & _
          "ORDER BY cu.usage_date DESC) AS T1 " & _
          "ORDER BY usage_date ASC) AS T2 " & _
          "ORDER BY usage_date DESC"

Set listRS = db99.Execute(listSQL)

' 지출 카테고리 이름 가져오기
Function GetCategoryName(categoryId)
    If IsNull(categoryId) Or categoryId = "" Then
        GetCategoryName = "-"
        Exit Function
    End If
    
    Dim categoryName, catSQL, catRS
    catSQL = "SELECT type_name FROM " & dbSchema & ".CardAccountTypes WHERE account_type_id = '" & categoryId & "'"
    
    On Error Resume Next
    Set catRS = db99.Execute(catSQL)
    
    If Err.Number = 0 And Not catRS.EOF Then
        categoryName = catRS("type_name")
    Else
        categoryName = categoryId
    End If
    
    If Not catRS Is Nothing Then
        If catRS.State = 1 Then
            catRS.Close
        End If
        Set catRS = Nothing
    End If
    
    GetCategoryName = categoryName
End Function

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

.btn-success {
    background: linear-gradient(to right, #28a745, #20c997);
    border: none;
    color: white;
}

.btn-success:hover {
    transform: translateY(-2px);
    box-shadow: 0 4px 12px rgba(40,167,69,0.2);
}

.btn-sm {
    padding: 0.5rem 1rem;
    font-size: 0.875rem;
    margin: 0 0.125rem;
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

.badge {
    font-size: 0.75rem;
    padding: 0.5rem 0.75rem;
    border-radius: 6px;
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
            <i class="fas fa-receipt me-2"></i>카드 사용 내역 관리
        </h2>
    </div>

    <!-- 검색 섹션 -->
    <div class="search-section">
        <div class="search-title">
            <i class="fas fa-search me-2"></i>카드 사용 내역 검색
        </div>
        <form action="admin_card_usage.asp" method="get">
            <div class="row g-3">
                <div class="col-md-3">
                    <label class="form-label">검색 필드</label>
                    <select name="field" class="form-select">
                        <option value="">전체</option>
                        <option value="user_id" <% If searchField = "user_id" Then %>selected<% End If %>>사용자</option>
                        <option value="card_id" <% If searchField = "card_id" Then %>selected<% End If %>>카드</option>
                        <option value="expense_category_id" <% If searchField = "expense_category_id" Then %>selected<% End If %>>지출 카테고리</option>
                    </select>
                </div>
                <div class="col-md-3">
                    <label class="form-label">검색어</label>
                    <input type="text" class="form-control" name="keyword" value="<%= searchKeyword %>" placeholder="검색어를 입력하세요">
                </div>
                <div class="col-md-2">
                    <label class="form-label">시작일</label>
                    <input type="date" class="form-control" name="date_from" value="<%= searchDateFrom %>">
                </div>
                <div class="col-md-2">
                    <label class="form-label">종료일</label>
                    <input type="date" class="form-control" name="date_to" value="<%= searchDateTo %>">
                </div>
                <div class="col-md-2">
                    <label class="form-label">&nbsp;</label>
                    <div class="d-flex gap-2">
                        <button type="submit" class="btn btn-primary flex-fill">
                            <i class="fas fa-search me-1"></i>검색
                        </button>
                        <button type="button" class="btn btn-success" onclick="exportToExcel()">
                            <i class="fas fa-file-excel me-1"></i>엑셀
                        </button>
                    </div>
                </div>
            </div>
        </form>
    </div>

    <!-- 카드 사용 내역 목록 -->
    <div class="table-section">
        <div class="table-title">
            <i class="fas fa-list me-2"></i>카드 사용 내역 목록 (총 <%= totalCount %>개)
        </div>
        
        <% If listRS.EOF Then %>
        <div class="empty-state">
            <i class="fas fa-receipt"></i>
            <h5>등록된 카드 사용 내역이 없습니다</h5>
            <p>검색 조건을 변경해보세요.</p>
        </div>
        <% Else %>
        <div class="table-responsive">
            <table class="table">
                <thead>
                    <tr>
                        <th style="text-align: center;">사용자</th>
                        <th style="text-align: center;">카드</th>
                        <th style="text-align: center;">사용일</th>
                        <th style="text-align: center;">사용처</th>
                        <th style="text-align: center;">금액</th>
                        <th style="text-align: center;">계정구분</th>
                        <th style="text-align: center;">승인상태</th>
                        <th style="text-align: center;">관리</th>
                    </tr>
                </thead>
                <tbody>
                    <% Do While Not listRS.EOF %>
                    <tr>
                        <td style="text-align: center;"><strong><%= listRS("user_name") %></strong></td>
                        <td style="text-align: center;"><%= listRS("card_id") %><br><%= listRS("issuer") %></td>
                        <td style="text-align: center;"><%= FormatDateTime(listRS("usage_date"), 2) %></td>
                        <td style="text-align: center;"><%= listRS("store_name") %></td>
                        <td style="text-align: center;"><strong><%= FormatCurrency(listRS("amount")) %></strong></td>
                        <td style="text-align: center;"><%= GetCategoryName(listRS("account_type_id")) %></td>
                        <td style="text-align: center;"><%= GetApprovalStatusBadge(listRS("approval_status")) %></td>
                        <td style="text-align: center;">
                            <a href="admin_card_usage_view.asp?id=<%= listRS("usage_id") %>" class="btn btn-sm btn-primary">
                                <i class="fas fa-eye"></i> 상세
                            </a>
                            <button class="btn btn-sm btn-danger" onclick="confirmDelete('<%= listRS("usage_id") %>')">
                                <i class="fas fa-trash"></i> 삭제
                            </button>
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
                    <a class="page-link" href="admin_card_usage.asp?page=<%= pageNo - 1 %>&field=<%= searchField %>&keyword=<%= searchKeyword %>&date_from=<%= searchDateFrom %>&date_to=<%= searchDateTo %>">
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
                    <a class="page-link" href="admin_card_usage.asp?page=<%= i %>&field=<%= searchField %>&keyword=<%= searchKeyword %>&date_from=<%= searchDateFrom %>&date_to=<%= searchDateTo %>"><%= i %></a>
                </li>
                <% Next %>
                
                <% If pageNo < totalPages Then %>
                <li class="page-item">
                    <a class="page-link" href="admin_card_usage.asp?page=<%= pageNo + 1 %>&field=<%= searchField %>&keyword=<%= searchKeyword %>&date_from=<%= searchDateFrom %>&date_to=<%= searchDateTo %>">
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

<script>
function confirmDelete(id) {
    if (confirm('정말로 이 카드 사용 내역을 삭제하시겠습니까?')) {
        window.location.href = "admin_card_usage.asp?action=delete&id=" + id;
    }
}

function exportToExcel() {
    // 현재 검색 조건을 가져와서 엑셀 다운로드 URL 생성
    const urlParams = new URLSearchParams(window.location.search);
    const field = urlParams.get('field') || '';
    const keyword = urlParams.get('keyword') || '';
    const dateFrom = urlParams.get('date_from') || '';
    const dateTo = urlParams.get('date_to') || '';
    
    const excelUrl = `admin_card_usage.asp?action=excel&field=${encodeURIComponent(field)}&keyword=${encodeURIComponent(keyword)}&date_from=${encodeURIComponent(dateFrom)}&date_to=${encodeURIComponent(dateTo)}`;
    
    window.location.href = excelUrl;
}
</script>

<%
' 사용한 객체 해제
If Not listRS Is Nothing Then
    If listRS.State = 1 Then
        listRS.Close
    End If
    Set listRS = Nothing
End If
%>

<!--#include file="../../includes/footer.asp"--> 