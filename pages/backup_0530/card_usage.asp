<%@ Language="VBScript" CodePage="65001" %>
<% 
Response.CodePage = 65001
Response.CharSet = "utf-8"
' DB 연결 확인용 쿼리를 추가합니다
Dim tableCheckSql, tableExists, dbTestConnection, tableSchema

On Error Resume Next
Set dbTestConnection = Server.CreateObject("ADODB.Connection")
dbTestConnection.ConnectionString = "Provider=SQLOLEDB;Data Source=121.175.77.251;Initial Catalog=balju_new;user ID=sa;password=WJStksrhk5030!;"
dbTestConnection.Open

Dim checkMsg
If Err.Number <> 0 Then
    checkMsg = "DB 연결 실패: " & Err.Description
    tableExists = False
Else
    checkMsg = "DB 연결 성공!"
    
    ' CardUsage 테이블 존재 여부 확인
    tableCheckSql = "SELECT CASE WHEN OBJECT_ID('dbo.CardUsage', 'U') IS NOT NULL THEN 1 ELSE 0 END AS table_exists"
    Dim checkRS
    Set checkRS = dbTestConnection.Execute(tableCheckSql)
    
    If Not checkRS.EOF Then
        tableExists = (checkRS("table_exists") = 1)
        If tableExists Then
            checkMsg = checkMsg & " CardUsage 테이블이 존재합니다."
            
            ' 테이블 구조 확인
            tableSchema = ""
            Dim schemaSQL, schemaRS
            schemaSQL = "SELECT COLUMN_NAME, DATA_TYPE, CHARACTER_MAXIMUM_LENGTH FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = 'CardUsage' ORDER BY ORDINAL_POSITION"
            Set schemaRS = dbTestConnection.Execute(schemaSQL)
            
            If Not schemaRS.EOF Then
                Do While Not schemaRS.EOF
                    tableSchema = tableSchema & schemaRS("COLUMN_NAME") & " (" & schemaRS("DATA_TYPE")
                    If Not IsNull(schemaRS("CHARACTER_MAXIMUM_LENGTH")) Then
                        tableSchema = tableSchema & "(" & schemaRS("CHARACTER_MAXIMUM_LENGTH") & ")"
                    End If
                    tableSchema = tableSchema & "), "
                    schemaRS.MoveNext
                Loop
            End If
            
            checkMsg = checkMsg & "<br>컬럼 구조: " & tableSchema
        Else
            checkMsg = checkMsg & " CardUsage 테이블이 존재하지 않습니다."
        End If
    Else
        tableExists = False
        checkMsg = checkMsg & " 테이블 확인 실패"
    End If
    
    checkRS.Close
    dbTestConnection.Close
End If
Err.Clear
On Error GoTo 0
%>

<!--#include file="../db.asp"-->
<!--#include file="../includes/functions.asp"-->
<%
' 로그인 체크
If Not IsAuthenticated() Then
    RedirectTo("../index.asp")
End If

' 엑셀 다운로드 처리
If Request.QueryString("action") = "excel" Then
    ' 검색 조건 가져오기
    Dim excelSearchCardId, excelSearchStartDate, excelSearchEndDate, excelSearchAccountType, excelSearchCondition
    excelSearchCardId = PreventSQLInjection(Request.QueryString("card_id"))
    excelSearchStartDate = PreventSQLInjection(Request.QueryString("start_date"))
    excelSearchEndDate = PreventSQLInjection(Request.QueryString("end_date"))
    excelSearchAccountType = PreventSQLInjection(Request.QueryString("account_type_id"))

    ' 검색 조건 SQL 생성
    If Session("user_id") <> "" Then 
        excelSearchCondition = " WHERE user_id = '" & Session("user_id") & "' "
    Else
        excelSearchCondition = " WHERE 1=1 "
    End If

    If excelSearchCardId <> "" Then
        excelSearchCondition = excelSearchCondition & " AND card_id = " & excelSearchCardId
    End If

    If excelSearchStartDate <> "" Then
        excelSearchCondition = excelSearchCondition & " AND usage_date >= '" & excelSearchStartDate & "'"
    End If

    If excelSearchEndDate <> "" Then
        excelSearchCondition = excelSearchCondition & " AND usage_date <= '" & excelSearchEndDate & "'"
    End If

    If excelSearchAccountType <> "" Then
        excelSearchCondition = excelSearchCondition & " AND expense_category_id = " & excelSearchAccountType
    End If

    ' 전체 데이터 조회 (페이징 없이)
    Dim excelSQL, excelRS
    excelSQL = "SELECT cu.usage_id, cu.user_id, cu.title, ca.account_name as card_name, ca.issuer, cu.usage_date, " & _
               "cu.store_name, cu.amount, cu.purpose, cu.approval_status, cu.expense_category_id, " & _
               "cat.type_name AS category_name " & _
               "FROM " & dbSchema & ".CardUsage cu " & _
               "LEFT JOIN " & dbSchema & ".CardAccount ca ON cu.card_id = ca.card_id " & _
               "LEFT JOIN " & dbSchema & ".CardAccountTypes cat ON cu.expense_category_id = cat.account_type_id " & _
               excelSearchCondition & " " & _
               "ORDER BY cu.usage_date DESC"

    Set excelRS = db99.Execute(excelSQL)

    ' 엑셀 파일 헤더 설정
    Response.ContentType = "application/vnd.ms-excel"
    Response.AddHeader "Content-Disposition", "attachment; filename=my_card_usage_" & Replace(Replace(Replace(Now(), "/", ""), ":", ""), " ", "_") & ".xls"
    Response.CharSet = "utf-8"
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
</head>
<body>
<table border="1">
    <tr>
        <th>제목</th>
        <th>카드</th>
        <th>발급사</th>
        <th>사용일</th>
        <th>사용처</th>
        <th>금액</th>
        <th>계정구분</th>
        <th>사용목적</th>
        <th>승인상태</th>
    </tr>
    <% Do While Not excelRS.EOF %>
    <tr>
        <td><%= excelRS("title") %></td>
        <td><%= excelRS("card_name") %></td>
        <td><%= excelRS("issuer") %></td>
        <td><%= FormatDateTime(excelRS("usage_date"), 2) %></td>
        <td><%= excelRS("store_name") %></td>
        <td><%= FormatNumber(excelRS("amount")) %></td>
        <td><%= IIf(IsNull(excelRS("category_name")), "-", excelRS("category_name")) %></td>
        <td><%= excelRS("purpose") %></td>
        <td><%= excelRS("approval_status") %></td>
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

' 세션에서 메시지 가져오기
Dim errorMsg, successMsg
If Session("success_msg") <> "" Then
    successMsg = Session("success_msg")
    Session("success_msg") = ""
End If

If Session("error_msg") <> "" Then
    errorMsg = Session("error_msg")
    Session("error_msg") = ""
End If

' 페이징 처리
Dim pageSize, currentPage, startRow, totalRows
pageSize = 10 ' 페이지당 표시할 레코드 수
currentPage = Request.QueryString("page")
If currentPage = "" Or Not IsNumeric(currentPage) Then
    currentPage = 1
Else
    currentPage = CInt(currentPage)
End If
startRow = (currentPage - 1) * pageSize

' 검색 조건
Dim searchCardId, searchStartDate, searchEndDate, searchAccountType
searchCardId = PreventSQLInjection(Request.QueryString("card_id"))
searchStartDate = PreventSQLInjection(Request.QueryString("start_date"))
searchEndDate = PreventSQLInjection(Request.QueryString("end_date"))
searchAccountType = PreventSQLInjection(Request.QueryString("account_type_id"))

' 검색 조건 SQL 생성
Dim searchCondition
If Session("user_id") <> "" Then 
    searchCondition = " WHERE user_id = '" & Session("user_id") & "' "
Else
    searchCondition = " WHERE 1=1 "
End If

If searchCardId <> "" Then
    searchCondition = searchCondition & " AND card_id = " & searchCardId
End If

If searchStartDate <> "" Then
    searchCondition = searchCondition & " AND usage_date >= '" & searchStartDate & "'"
End If

If searchEndDate <> "" Then
    searchCondition = searchCondition & " AND usage_date <= '" & searchEndDate & "'"
End If

If searchAccountType <> "" Then
    searchCondition = searchCondition & " AND expense_category_id = " & searchAccountType
End If

' 총 레코드 수 조회
Dim countSQL, countRS
countSQL = "SELECT COUNT(*) AS total FROM " & dbSchema & ".CardUsage" & searchCondition
Set countRS = db.Execute(countSQL)

If Err.Number <> 0 Then
    totalRows = 0
    Err.Clear
Else
    totalRows = countRS("total")
    countRS.Close
End If

' 전체 페이지 수 계산
Dim totalPages
totalPages = Ceil(totalRows / pageSize)
If totalPages < 1 Then totalPages = 1

' 카드 사용 내역 조회 - 페이징 처리 포함
Dim SQL, rs, sqlDebugInfo

' 페이징 처리된 쿼리
If currentPage = 1 Then
    ' 첫 페이지
    SQL = "SELECT TOP " & pageSize & " usage_id, user_id, card_id, usage_date, amount, store_name, purpose, title, " & _
          "approval_status, department_id, expense_category_id, cost_type_id " & _
          "FROM " & dbSchema & ".CardUsage" & searchCondition & _
          " ORDER BY usage_date DESC, usage_id DESC"
Else
    ' 2페이지 이상 - NOT IN 방식 사용
    SQL = "SELECT TOP " & pageSize & " usage_id, user_id, card_id, usage_date, amount, store_name, purpose, title, " & _
          " approval_status, department_id, expense_category_id, cost_type_id " & _
          "FROM " & dbSchema & ".CardUsage" & searchCondition & _
          " AND usage_id NOT IN (" & _
          "SELECT TOP " & startRow & " usage_id FROM " & dbSchema & ".CardUsage" & searchCondition & _
          " ORDER BY usage_date DESC, usage_id DESC)" & _
          " ORDER BY usage_date DESC, usage_id DESC"
End If

' 디버깅용 SQL 정보 저장
sqlDebugInfo = "실행 쿼리: " & SQL

Set rs = db.Execute(SQL)
sqlDebugInfo = sqlDebugInfo & "<br>에러 여부: " & (Err.Number <> 0) & "<br>에러 번호: " & Err.Number & "<br>에러 설명: " & Err.Description
On Error GoTo 0

' 카드 목록 조회
Dim cardSQL, cardRS
cardSQL = "SELECT card_id, account_name, issuer FROM " & dbSchema & ".CardAccount ORDER BY account_name"

On Error Resume Next
Set cardRS = db99.Execute(cardSQL)

' 계정 과목 목록 조회
Dim accountTypeSQL, accountTypeRS
accountTypeSQL = "SELECT account_type_id, type_name FROM " & dbSchema & ".CardAccountTypes ORDER BY type_name"

On Error Resume Next
Set accountTypeRS = db.Execute(accountTypeSQL)
If Err.Number <> 0 Then
    Err.Clear
    ' 대체 테이블 또는 뷰로 시도
    accountTypeSQL = "SELECT id AS account_type_id, name AS type_name FROM " & dbSchema & ".CardAccountType ORDER BY name"
    Set accountTypeRS = db.Execute(accountTypeSQL)
End If
If Err.Number <> 0 Then
    Err.Clear
    ' 또 다른 대체 테이블 시도
    accountTypeSQL = "SELECT expense_category_id AS account_type_id, category_name AS type_name FROM " & dbSchema & ".ExpenseCategory ORDER BY category_name"
    Set accountTypeRS = db.Execute(accountTypeSQL)
End If

' 카드 계정 이름과 계정 타입 이름 조회를 위한 함수
Function GetCardName(cardId)
    Dim cardNumber, cardName
    
    cardNumber = "알 수 없음"
    cardName = "알 수 없음"
    
    ' 메모리 객체에서 먼저 찾기
    If Not cardRS Is Nothing And Not cardRS.EOF Then
        cardRS.MoveFirst
        Do While Not cardRS.EOF
            If CStr(cardRS("card_id")) = CStr(cardId) Then
                cardNumber = cardRS("account_name")
                cardName = cardRS("issuer")

                Exit Do
            End If
            cardRS.MoveNext
        Loop
    End If
    
    ' 메모리에 없으면 DB에서 직접 조회
    If cardNumber = "알 수 없음" Then
        Dim directSQL, directRS
        directSQL = "SELECT account_name, issuer FROM " & dbSchema & ".CardAccount WHERE card_id = " & cardId
        
        On Error Resume Next
        Set directRS = db99.Execute(directSQL)
        
        
        If Err.Number = 0 And Not directRS.EOF Then
            cardNumber = directRS("account_name")
            cardName = directRS("issuer")
        End If
        
        If Not directRS Is Nothing Then
            If directRS.State = 1 Then ' 1 = adStateOpen
                directRS.Close
            End If
            Set directRS = Nothing
        End If
        On Error GoTo 0
    End If
    
    GetCardName = cardName & " (" & cardNumber & ")"
End Function

' 계정 유형 정보는 expense_category_id를 통해 얻습니다
Function GetExpenseCategoryName(categoryId)
    If categoryId = "" Or Not IsNumeric(categoryId) Then
        GetExpenseCategoryName = "-"
        Exit Function
    End If
    
    Dim typeName
    typeName = "-"
    
    ' 메모리 객체에서 먼저 찾기
    If Not accountTypeRS Is Nothing And Not accountTypeRS.EOF Then
        accountTypeRS.MoveFirst
        Do While Not accountTypeRS.EOF
            If CStr(accountTypeRS("account_type_id")) = CStr(categoryId) Then
                typeName = accountTypeRS("type_name")
                Exit Do
            End If
            accountTypeRS.MoveNext
        Loop
    End If
    
    ' 메모리에 없으면 DB에서 직접 조회
    If typeName = "-" Then
        Dim catSQL, catRS
        catSQL = "SELECT type_name FROM " & dbSchema & ".CardAccountTypes WHERE account_type_id = " & categoryId
        
        On Error Resume Next
        Set catRS = db.Execute(catSQL)
        
        If Err.Number <> 0 Then
            Err.Clear
            ' 대체 테이블 시도
            catSQL = "SELECT name AS type_name FROM " & dbSchema & ".CardAccountType WHERE id = " & categoryId
            Set catRS = db.Execute(catSQL)
        End If
        
        If Err.Number <> 0 Then
            Err.Clear
            ' 또 다른 대체 테이블 시도
            catSQL = "SELECT category_name AS type_name FROM " & dbSchema & ".ExpenseCategory WHERE expense_category_id = " & categoryId
            Set catRS = db.Execute(catSQL)
        End If
        
        If Err.Number = 0 And Not catRS.EOF Then
            typeName = catRS("type_name")
        End If
        
        If Not catRS Is Nothing Then
            If catRS.State = 1 Then
                catRS.Close
            End If
            Set catRS = Nothing
        End If
        On Error GoTo 0
    End If
    
    GetExpenseCategoryName = typeName
End Function

' 페이징 함수
Function Ceil(number)
    Ceil = Int(number)
    If Ceil <> number Then
        Ceil = Ceil + 1
    End If
End Function

On Error GoTo 0
%>
<!--#include file="../includes/header.asp"-->

<style>
.container {
    max-width: 1400px;
    margin: 0 auto;
    padding: 2rem 1rem;
}

.page-header {
    display: flex;
    justify-content: space-between;
    align-items: center;
    margin-bottom: 2rem;
    padding: 1.5rem;
    background: white;
    border-radius: 16px;
    box-shadow: 0 4px 20px rgba(0,0,0,0.08);
}

.page-title {
    font-size: 1.5rem;
    font-weight: 600;
    color: #2C3E50;
    margin: 0;
    display: flex;
    align-items: center;
}

.btn-group-nav {
    display: flex;
    gap: 0.5rem;
}

.btn-nav {
    padding: 0.875rem 1.5rem;
    font-size: 0.9rem;
    font-weight: 600;
    border-radius: 8px;
    transition: all 0.2s ease;
}

.card {
    border: none;
    box-shadow: 0 4px 20px rgba(0,0,0,0.08);
    border-radius: 16px;
    margin-bottom: 2rem;
    background: #fff;
    overflow: hidden;
}

.card-header {
    background: linear-gradient(135deg, #E8F2FF 0%, #F0F8FF 100%);
    border-bottom: 1px solid #E2E8F0;
    padding: 1.5rem;
}

.card-header h5 {
    color: #475569;
    font-weight: 600;
    margin: 0;
    font-size: 1.1rem;
}

.filter-buttons {
    display: flex;
    gap: 0.5rem;
    margin-top: 1rem;
}

.filter-btn {
    padding: 0.5rem 1rem;
    border: 1px solid #CBD5E1;
    background: #F8FAFC;
    color: #64748B;
    text-decoration: none;
    border-radius: 6px;
    font-weight: 500;
    transition: all 0.2s ease;
    font-size: 0.875rem;
}

.filter-btn:hover {
    background: #E2E8F0;
    border-color: #94A3B8;
    color: #475569;
    text-decoration: none;
    transform: translateY(-1px);
}

.filter-btn.active {
    background: #E0F2FE;
    border-color: #0EA5E9;
    color: #0369A1;
    box-shadow: 0 2px 8px rgba(14,165,233,0.15);
}

.badge {
    padding: 0.375rem 0.75rem;
    font-weight: 500;
    border-radius: 6px;
    font-size: 0.8rem;
}

.badge-success {
    background: #DCFCE7;
    color: #166534;
    border: 1px solid #BBF7D0;
}

.badge-danger {
    background: #FEE2E2;
    color: #DC2626;
    border: 1px solid #FECACA;
}

.badge-primary {
    background: #DBEAFE;
    color: #1D4ED8;
    border: 1px solid #BFDBFE;
}

.badge-info {
    background: #E0F2FE;
    color: #0369A1;
    border: 1px solid #BAE6FD;
}

.badge-secondary {
    background: #F1F5F9;
    color: #475569;
    border: 1px solid #E2E8F0;
}

.badge-outline {
    background: transparent;
    border: 1px solid #E5E7EB;
    color: #6B7280;
}

.table {
    margin-bottom: 0;
}

.table th {
    background: linear-gradient(135deg, #F1F5F9 0%, #E2E8F0 100%);
    color: #475569;
    font-weight: 600;
    border: none;
    padding: 0.875rem;
    font-size: 0.9rem;
    white-space: nowrap;
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

.date-cell {
    font-size: 0.9rem;
    font-weight: 500;
    white-space: nowrap;
    min-width: 120px;
}

.amount-cell {
    font-weight: 600;
    color: #059669;
    text-align: right;
    white-space: nowrap;
}

.btn-sm {
    padding: 0.5rem 1rem;
    font-size: 0.875rem;
    border-radius: 6px;
    font-weight: 500;
}

.btn-outline-primary {
    border: 2px solid #4A90E2;
    color: #4A90E2;
    background: transparent;
}

.btn-outline-primary:hover {
    background: #4A90E2;
    color: white;
    transform: translateY(-2px);
    box-shadow: 0 4px 12px rgba(74,144,226,0.2);
}

.btn-secondary {
    background: #6B7280;
    color: white;
    border: none;
}

.btn-secondary:hover {
    background: #4B5563;
    transform: translateY(-2px);
}

.btn-primary {
    background: #4A90E2;
    color: white;
    border: none;
}

.btn-primary:hover {
    background: #357ABD;
    transform: translateY(-2px);
}

.pagination {
    margin-top: 2rem;
}

.page-link {
    border: none;
    padding: 1rem 1.25rem;
    margin: 0 0.25rem;
    border-radius: 8px;
    color: #2C3E50;
    background: #F8FAFC;
    transition: all 0.2s ease;
    font-weight: 500;
    min-height: 48px;
    display: flex;
    align-items: center;
    justify-content: center;
}

.page-link:hover {
    background: #E9ECEF;
    color: #2C3E50;
    transform: translateY(-2px);
}

.page-item.active .page-link {
    background: #4A90E2;
    color: white;
    box-shadow: 0 4px 12px rgba(74,144,226,0.2);
}

.empty-state {
    text-align: center;
    padding: 4rem 2rem;
    color: #64748B;
}

.empty-state i {
    font-size: 4rem;
    margin-bottom: 1rem;
    color: #CBD5E1;
}

.empty-state h5 {
    color: #64748B;
    margin-bottom: 0.5rem;
}

.empty-state p {
    color: #94A3B8;
}

.form-group {
    margin-bottom: 1rem;
}

.form-group label {
    display: block;
    margin-bottom: 0.5rem;
    font-weight: 600;
    color: #2C3E50;
}

.form-group input,
.form-group select {
    width: 100%;
    padding: 0.75rem;
    border: 2px solid #E9ECEF;
    border-radius: 8px;
    font-size: 1rem;
    transition: border-color 0.2s ease;
}

.form-group input:focus,
.form-group select:focus {
    outline: none;
    border-color: #4A90E2;
    box-shadow: 0 0 0 3px rgba(74,144,226,0.1);
}

.btn {
    padding: 0.75rem 1.5rem;
    border-radius: 8px;
    font-weight: 600;
    text-decoration: none;
    display: inline-block;
    transition: all 0.2s ease;
    border: none;
    cursor: pointer;
}

.btn-outline {
    background: transparent;
    border: 2px solid #E9ECEF;
    color: #6B7280;
}

.btn-outline:hover {
    background: #F3F4F6;
    color: #374151;
}

.alert {
    padding: 1rem;
    border-radius: 8px;
    margin-bottom: 1rem;
}

.alert-error {
    background: #FEF2F2;
    border: 1px solid #FECACA;
    color: #B91C1C;
}

.alert-success {
    background: #F0FDF4;
    border: 1px solid #BBF7D0;
    color: #166534;
}
</style>

<div class="container">
    <div class="page-header">
        <h2 class="page-title">
            <i class="fas fa-credit-card me-2"></i>카드 사용 내역
        </h2>
        <div class="btn-group-nav">
            <a href="dashboard.asp" class="btn btn-secondary btn-nav">
                <i class="fas fa-home me-1"></i> 대시보드
            </a>
            <a href="card_usage_add.asp" class="btn btn-primary btn-nav">
                <i class="fas fa-plus me-1"></i> 새 사용 내역 등록
            </a>
        </div>
    </div>

    <div class="card">
        <div class="card-header">
            <h5><i class="fas fa-search me-2"></i>검색 조건</h5>
        </div>
        <div class="card-body">
        
            <% If errorMsg <> "" Then %>
            <div class="alert alert-error">
                <div>
                    <span><strong>오류 : </strong> <%= errorMsg %></span>
                </div>
            </div>
            <% End If %>
            
            <% If successMsg <> "" Then %>
            <div class="alert alert-success">
                <div>
                    <span><strong>성공 : </strong> <%= successMsg %></span>
                </div>
            </div>
            <% End If %>
            
            <!-- 검색 폼 -->
            <form id="searchForm" method="get" action="card_usage.asp">
                <div class="row justify-content-center">
                    <div class="col-lg-12">
                        <div class="row g-3 mb-3">
                            <div class="col-md-3">
                                <div class="form-group">
                                    <label for="card_id">카드</label>
                                    <select id="card_id" name="card_id" class="form-select">
                                        <option value="">전체</option>
                                        <% 
                                        If Not cardRS Is Nothing And Not cardRS.EOF Then
                                            cardRS.MoveFirst
                                            Do While Not cardRS.EOF 
                                                Dim cardSelected
                                                cardSelected = ""
                                                If CStr(cardRS("card_id")) = searchCardId Then
                                                    cardSelected = "selected"
                                                End If
                                        %>
                                            <option value="<%= cardRS("card_id") %>" <%= cardSelected %>><%= cardRS("account_name") %></option>
                                        <% 
                                                cardRS.MoveNext
                                            Loop
                                        End If
                                        %>
                                    </select>
                                </div>
                            </div>
                            
                            <div class="col-md-3">
                                <div class="form-group">
                                    <label for="account_type_id">계정 과목</label>
                                    <select id="account_type_id" name="account_type_id" class="form-select">
                                        <option value="">전체</option>
                                        <% 
                                        If Not accountTypeRS Is Nothing And Not accountTypeRS.EOF Then
                                            accountTypeRS.MoveFirst
                                            Do While Not accountTypeRS.EOF 
                                                Dim typeSelected
                                                typeSelected = ""
                                                If CStr(accountTypeRS("account_type_id")) = searchAccountType Then
                                                    typeSelected = "selected"
                                                End If
                                        %>
                                            <option value="<%= accountTypeRS("account_type_id") %>" <%= typeSelected %>><%= accountTypeRS("type_name") %></option>
                                        <% 
                                                accountTypeRS.MoveNext
                                            Loop
                                        End If
                                        %>
                                    </select>
                                </div>
                            </div>
                            
                            <div class="col-md-2">
                                <div class="form-group">
                                    <label for="start_date">시작일</label>
                                    <input type="date" id="start_date" name="start_date" value="<%= searchStartDate %>" class="form-control">
                                </div>
                            </div>
                            
                            <div class="col-md-2">
                                <div class="form-group">
                                    <label for="end_date">종료일</label>
                                    <input type="date" id="end_date" name="end_date" value="<%= searchEndDate %>" class="form-control">
                                </div>
                            </div>
                            
                            <div class="col-md-2 d-flex align-items-end">
                                <div class="form-group w-100">
                                    <button type="submit" class="btn btn-primary w-100 mb-2">
                                        <i class="fas fa-search me-1"></i>검색
                                    </button>
                                </div>
                            </div>
                        </div>
                        
                        <!-- 버튼 그룹을 별도 행으로 분리 -->
                        <div class="row">
                            <div class="col-12">
                                <div class="d-flex gap-2 justify-content-center">
                                    <a href="card_usage.asp" class="btn btn-outline">
                                        <i class="fas fa-refresh me-1"></i>초기화
                                    </a>
                                    <button type="button" class="btn btn-success" onclick="exportToExcel()">
                                        <i class="fas fa-file-excel me-1"></i>엑셀
                                    </button>
                                    <button type="button" class="btn btn-info" onclick="printList()">
                                        <i class="fas fa-print me-1"></i>인쇄
                                    </button>
                                </div>
                            </div>
                        </div>
                    </div>
                </div>
            </form>
        </div>
    </div>
    
    <!-- 카드 사용 내역 목록 -->
    <div class="card" id="printArea">
        <div class="card-header">
            <h5><i class="fas fa-list me-2"></i>카드 사용 내역</h5>
        </div>
        <div class="card-body">
            <!-- 검색 조건 표시 (인쇄용) -->
            <div id="searchConditions" class="mb-3" style="display: none;">
                <h6>검색 조건</h6>
                <div class="row">
                    <div class="col-md-3">
                        <strong>카드:</strong> 
                        <span id="printCardName">
                            <% 
                            If searchCardId <> "" And Not cardRS Is Nothing Then
                                cardRS.MoveFirst
                                Do While Not cardRS.EOF
                                    If CStr(cardRS("card_id")) = searchCardId Then
                                        Response.Write cardRS("account_name")
                                        Exit Do
                                    End If
                                    cardRS.MoveNext
                                Loop
                            Else
                                Response.Write "전체"
                            End If
                            %>
                        </span>
                    </div>
                    <div class="col-md-3">
                        <strong>계정과목:</strong> 
                        <span id="printAccountType">
                            <% 
                            If searchAccountType <> "" And Not accountTypeRS Is Nothing Then
                                accountTypeRS.MoveFirst
                                Do While Not accountTypeRS.EOF
                                    If CStr(accountTypeRS("account_type_id")) = searchAccountType Then
                                        Response.Write accountTypeRS("type_name")
                                        Exit Do
                                    End If
                                    accountTypeRS.MoveNext
                                Loop
                            Else
                                Response.Write "전체"
                            End If
                            %>
                        </span>
                    </div>
                    <div class="col-md-3">
                        <strong>시작일:</strong> 
                        <span id="printStartDate"><%= IIf(searchStartDate <> "", searchStartDate, "전체") %></span>
                    </div>
                    <div class="col-md-3">
                        <strong>종료일:</strong> 
                        <span id="printEndDate"><%= IIf(searchEndDate <> "", searchEndDate, "전체") %></span>
                    </div>
                </div>
                <hr>
            </div>
            
            <% 
            ' 디버깅용 DB 연결 상태와 에러 표시
            If Not dbConnected Then 
            %>
                <div class="alert alert-error">
                    <div>
                        <span><strong>데이터베이스 연결 오류 : </strong> <%= dbErrorMsg %></span>
                    </div>
                </div>
            <% End If %>
            
            <% If Err.Number <> 0 Then %>
                <div class="alert alert-error">
                    <div>
                        <span><strong>SQL 오류 : </strong> <%= Err.Description %></span>
                    </div>
                </div>
            <% End If %>
            
            <% 
            If rs.EOF Then
            %>
                <div class="empty-state">
                    <i class="fas fa-credit-card"></i>
                    <h5>등록된 카드 사용 내역이 없습니다</h5>
                    <p>새로운 카드 사용 내역을 등록해보세요.</p>
                </div>
            <% Else %>
                <div class="table-responsive">
                    <table class="table table-hover">
                        <thead>
                            <tr>
                                <th style="text-align: center;">사용일자</th>
                                <th style="text-align: center;">카드</th>
                                <th style="text-align: center;">계정 과목</th>
                                <th style="text-align: center;">제목</th>
                                <th style="text-align: center;">사용처</th>
                                <th style="text-align: center;">사용 목적</th>
                                <th style="text-align: center;">금액</th>
                                <th style="text-align: center;">상태</th>
                                <th style="text-align: center;">관리</th>
                            </tr>
                        </thead>
                        <tbody>
                            <% 
                            Do While Not rs.EOF 
                            %>
                            <tr>
                                <td style="text-align: center;" class="date-cell"><%= FormatDate(rs("usage_date")) %></td>
                                <td style="text-align: center;"><%= GetCardName(rs("card_id")) %></td>
                                <td style="text-align: center;"><%= GetExpenseCategoryName(rs("expense_category_id")) %></td>
                                <td style="text-align: center;"><% 
                                    If Not IsNull(rs("title")) Then 
                                        Response.Write(rs("title"))
                                    ElseIf Not IsNull(rs("store_name")) Then
                                        Response.Write(rs("store_name"))
                                    Else
                                        Response.Write("-")
                                    End If
                                %></td>
                                <td style="text-align: center;"><% 
                                    If Not IsNull(rs("store_name")) Then 
                                        Response.Write(rs("store_name"))
                                    Else
                                        Response.Write("-")
                                    End If
                                %></td>
                                <td style="text-align: center;"><% 
                                    If Not IsNull(rs("purpose")) Then 
                                        Response.Write(rs("purpose"))
                                    Else
                                        Response.Write("-")
                                    End If
                                %></td>
                                <td style="text-align: center;" class="amount-cell"><%= FormatNumber(rs("amount")) %>원</td>
                                <td style="text-align: center;">
                                    <% 
                                    Dim statusClass, statusText
                                    If Not IsNull(rs("approval_status")) Then 
                                        statusText = rs("approval_status")
                                    Else
                                        statusText = "처리중"
                                    End If
                                    
                                    Select Case statusText
                                        Case "승인"
                                            statusClass = "badge badge-success"
                                        Case "반려"
                                            statusClass = "badge badge-danger"
                                        Case "대기"
                                            statusClass = "badge badge-info"
                                        Case Else
                                            statusClass = "badge badge-secondary"
                                    End Select
                                    %>
                                    <span class="<%= statusClass %>">
                                        <% If statusText = "승인" Then %>
                                            <i class="fas fa-check me-1"></i>
                                        <% ElseIf statusText = "반려" Then %>
                                            <i class="fas fa-times me-1"></i>
                                        <% ElseIf statusText = "대기" Then %>
                                            <i class="fas fa-clock me-1"></i>
                                        <% Else %>
                                            <i class="fas fa-edit me-1"></i>
                                        <% End If %>
                                        <%= statusText %>
                                    </span>
                                </td>
                                <td style="text-align: center;">
                                    <div style="display: flex; gap: 5px; justify-content: center;">
                                        <a href="card_usage_view.asp?id=<%= rs("usage_id") %>" class="btn btn-sm btn-outline-primary">
                                            <i class="fas fa-eye me-1"></i>상세
                                        </a>
                                        <% If rs("approval_status") <> "완료" Then %>
                                        <a href="card_usage_edit.asp?id=<%= rs("usage_id") %>" class="btn btn-sm btn-secondary">
                                            <i class="fas fa-edit me-1"></i>수정
                                        </a>
                                        <% End If %>
                                    </div>
                                </td>
                            </tr>
                            <% 
                                rs.MoveNext
                                Loop 
                            %>
                        </tbody>
                    </table>
                </div>
            <% End If %>
            
            <!-- 페이징 -->
            <% If totalPages > 1 Then %>
                <div class="d-flex justify-content-center mt-4">
                    <nav aria-label="Page navigation">
                        <ul class="pagination">
                            <% If currentPage > 1 Then %>
                                <li class="page-item">
                                    <a class="page-link" href="card_usage.asp?page=<%= currentPage - 1 %>&card_id=<%= searchCardId %>&start_date=<%= searchStartDate %>&end_date=<%= searchEndDate %>&account_type_id=<%= searchAccountType %>">
                                        <i class="fas fa-chevron-left"></i> 이전
                                    </a>
                                </li>
                            <% End If %>

                            <% 
                            Dim startPage, endPage
                            startPage = ((currentPage - 1) \ 5) * 5 + 1
                            endPage = startPage + 4
                            If endPage > totalPages Then endPage = totalPages

                            For i = startPage To endPage
                            %>
                                <li class="page-item <%= IIf(i = currentPage, "active", "") %>">
                                    <a class="page-link" href="card_usage.asp?page=<%= i %>&card_id=<%= searchCardId %>&start_date=<%= searchStartDate %>&end_date=<%= searchEndDate %>&account_type_id=<%= searchAccountType %>"><%= i %></a>
                                </li>
                            <% Next %>

                            <% If currentPage < totalPages Then %>
                                <li class="page-item">
                                    <a class="page-link" href="card_usage.asp?page=<%= currentPage + 1 %>&card_id=<%= searchCardId %>&start_date=<%= searchStartDate %>&end_date=<%= searchEndDate %>&account_type_id=<%= searchAccountType %>">
                                        다음 <i class="fas fa-chevron-right"></i>
                                    </a>
                                </li>
                            <% End If %>
                        </ul>
                    </nav>
                </div>
            <% End If %>
        </div>
    </div>
</div>

<script>
// 엑셀 저장 기능
function exportToExcel() {
    // 현재 검색 조건을 가져와서 엑셀 다운로드 URL 생성
    const urlParams = new URLSearchParams(window.location.search);
    const cardId = urlParams.get('card_id') || '';
    const startDate = urlParams.get('start_date') || '';
    const endDate = urlParams.get('end_date') || '';
    const accountType = urlParams.get('account_type_id') || '';
    
    const excelUrl = `card_usage.asp?action=excel&card_id=${encodeURIComponent(cardId)}&start_date=${encodeURIComponent(startDate)}&end_date=${encodeURIComponent(endDate)}&account_type_id=${encodeURIComponent(accountType)}`;
    
    window.location.href = excelUrl;
}

// 인쇄 기능
function printList() {
    // 현재 검색 조건을 가져오기
    const urlParams = new URLSearchParams(window.location.search);
    const cardId = urlParams.get('card_id') || '';
    const startDate = urlParams.get('start_date') || '';
    const endDate = urlParams.get('end_date') || '';
    const accountType = urlParams.get('account_type_id') || '';
    
    // 전체 데이터를 가져오기 위한 URL 생성 (페이징 없이)
    const printUrl = `card_usage_print_data.asp?card_id=${encodeURIComponent(cardId)}&start_date=${encodeURIComponent(startDate)}&end_date=${encodeURIComponent(endDate)}&account_type_id=${encodeURIComponent(accountType)}`;
    
    // AJAX로 전체 데이터 가져오기
    fetch(printUrl)
        .then(response => response.text())
        .then(data => {
            // 인쇄 창 생성
            const printWindow = window.open('', '', 'width=1200,height=900');
            
            printWindow.document.write('<html><head><title>카드 사용 내역</title>');
            printWindow.document.write('<style>');
            printWindow.document.write(`
                @media print {
                    @page {
                        margin: 15mm;
                        size: A4 landscape;
                    }
                }
                
                body { 
                    font-family: 'Malgun Gothic', Arial, sans-serif; 
                    padding: 20px; 
                    margin: 0;
                    font-size: 12px;
                    line-height: 1.4;
                    color: #000;
                }
                
                .print-header {
                    text-align: center;
                    margin-bottom: 30px;
                    border-bottom: 2px solid #000;
                    padding-bottom: 10px;
                }
                
                .print-header h2 {
                    margin: 0;
                    font-size: 20px;
                    font-weight: 600;
                    color: #000;
                }
                
                .print-date {
                    margin: 5px 0 0 0;
                    font-size: 12px;
                    color: #666;
                    text-align: right;
                }
                
                .search-conditions {
                    margin-bottom: 20px;
                    padding: 15px;
                    background: #f8f9fa;
                    border: 1px solid #000;
                }
                
                .search-conditions h6 {
                    font-size: 14px;
                    font-weight: 600;
                    margin-bottom: 10px;
                    color: #000;
                }
                
                .condition-row {
                    display: flex;
                    flex-wrap: wrap;
                    margin: 0 -5px;
                }
                
                .condition-col {
                    flex: 0 0 25%;
                    padding: 0 5px;
                    margin-bottom: 8px;
                }
                
                .table {
                    width: 100%;
                    border-collapse: collapse;
                    margin-bottom: 20px;
                }
                
                .table th {
                    background: #f8f9fa !important;
                    border: 2px solid #000 !important;
                    padding: 8px !important;
                    text-align: center;
                    font-weight: 600;
                    color: #000 !important;
                    font-size: 11px;
                }
                
                .table td {
                    border: 1px solid #000 !important;
                    padding: 6px !important;
                    text-align: center;
                    font-size: 10px;
                    color: #000;
                }
                
                .amount-cell {
                    text-align: right !important;
                    font-weight: 600;
                }
                
                .date-cell {
                    white-space: nowrap;
                    font-weight: 500;
                }
                
                .status-badge {
                    padding: 2px 6px;
                    border: 1px solid #000;
                    border-radius: 3px;
                    font-size: 9px;
                    background: #fff !important;
                    color: #000 !important;
                }
            `);
            printWindow.document.write('</style></head><body>');
            
            // 인쇄용 제목 추가
            printWindow.document.write('<div class="print-header">');
            printWindow.document.write('<h2>카드 사용 내역</h2>');
            printWindow.document.write('</div>');
            
            printWindow.document.write('<p class="print-date">출력일: ' + new Date().toLocaleDateString('ko-KR') + '</p>');
            
            // 검색 조건 표시
            printWindow.document.write('<div class="search-conditions">');
            printWindow.document.write('<h6>검색 조건</h6>');
            printWindow.document.write('<div class="condition-row">');
            printWindow.document.write('<div class="condition-col"><strong>카드:</strong> ' + (cardId ? getCardNameForPrint(cardId) : '전체') + '</div>');
            printWindow.document.write('<div class="condition-col"><strong>계정과목:</strong> ' + (accountType ? getAccountTypeNameForPrint(accountType) : '전체') + '</div>');
            printWindow.document.write('<div class="condition-col"><strong>시작일:</strong> ' + (startDate || '전체') + '</div>');
            printWindow.document.write('<div class="condition-col"><strong>종료일:</strong> ' + (endDate || '전체') + '</div>');
            printWindow.document.write('</div>');
            printWindow.document.write('</div>');
            
            // 데이터 표시
            printWindow.document.write(data);
            
            printWindow.document.write('</body></html>');
            printWindow.document.close();
            printWindow.focus();
            
            // 인쇄 실행
            setTimeout(function() {
                printWindow.print();
                printWindow.close();
            }, 500);
        })
        .catch(error => {
            console.error('인쇄 데이터 로드 실패:', error);
            alert('인쇄 데이터를 불러오는데 실패했습니다.');
        });
}

// 카드명 가져오기 (인쇄용)
function getCardNameForPrint(cardId) {
    const cardSelect = document.getElementById('card_id');
    if (cardSelect) {
        for (let option of cardSelect.options) {
            if (option.value === cardId) {
                return option.text;
            }
        }
    }
    return '알 수 없음';
}

// 계정과목명 가져오기 (인쇄용)
function getAccountTypeNameForPrint(accountTypeId) {
    const accountSelect = document.getElementById('account_type_id');
    if (accountSelect) {
        for (let option of accountSelect.options) {
            if (option.value === accountTypeId) {
                return option.text;
            }
        }
    }
    return '알 수 없음';
}
</script>

<!--#include file="../includes/footer.asp"--> 