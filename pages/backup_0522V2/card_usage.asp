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

On Error Resume Next

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

' 카드 사용 내역 조회
Dim SQL, rs, sqlDebugInfo

' 기본 쿼리 - 실제 DB 구조에 맞게 필드 지정
SQL = "SELECT usage_id, user_id, card_id, usage_date, amount, store_name, purpose, title, " & _
      "linked_table, linked_id, approval_status, department_id, expense_category_id " & _
      "FROM " & dbSchema & ".CardUsage" & searchCondition

' 정렬 추가
SQL = SQL & " ORDER BY usage_date DESC, usage_id DESC"

' 디버깅용 SQL 정보 저장
sqlDebugInfo = "실행 쿼리: " & SQL

Set rs = db.Execute(SQL)
sqlDebugInfo = sqlDebugInfo & "<br>에러 여부: " & (Err.Number <> 0) & "<br>에러 번호: " & Err.Number & "<br>에러 설명: " & Err.Description
On Error GoTo 0

' 카드 목록 조회
Dim cardSQL, cardRS
cardSQL = "SELECT card_id, account_name FROM " & dbSchema & ".CardAccount ORDER BY account_name"

On Error Resume Next
Set cardRS = db.Execute(cardSQL)
If Err.Number <> 0 Then
    Err.Clear
    ' 대체 테이블 또는 뷰로 시도
    cardSQL = "SELECT card_id, name AS account_name FROM " & dbSchema & ".Card ORDER BY name"
    Set cardRS = db.Execute(cardSQL)
End If

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
    Dim cardName
    cardName = "알 수 없음"
    
    ' 메모리 객체에서 먼저 찾기
    If Not cardRS Is Nothing And Not cardRS.EOF Then
        cardRS.MoveFirst
        Do While Not cardRS.EOF
            If CStr(cardRS("card_id")) = CStr(cardId) Then
                cardName = cardRS("account_name")
                Exit Do
            End If
            cardRS.MoveNext
        Loop
    End If
    
    ' 메모리에 없으면 DB에서 직접 조회
    If cardName = "알 수 없음" Then
        Dim directSQL, directRS
        directSQL = "SELECT account_name FROM " & dbSchema & ".CardAccount WHERE card_id = " & cardId
        
        On Error Resume Next
        Set directRS = db.Execute(directSQL)
        
        If Err.Number <> 0 Then
            Err.Clear
            ' 대체 테이블 시도
            directSQL = "SELECT name AS account_name FROM " & dbSchema & ".Card WHERE card_id = " & cardId
            Set directRS = db.Execute(directSQL)
        End If
        
        If Err.Number = 0 And Not directRS.EOF Then
            cardName = directRS("account_name")
        Else
            cardName = cardId
        End If
        
        If Not directRS Is Nothing Then
            If directRS.State = 1 Then ' 1 = adStateOpen
                directRS.Close
            End If
            Set directRS = Nothing
        End If
        On Error GoTo 0
    End If
    
    GetCardName = cardName
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

<div class="card-usage-container">
    <div class="shadcn-card" style="margin-bottom: 20px;">
        <div class="shadcn-card-header">
            <h2 class="shadcn-card-title">카드 사용 내역</h2>
            <p class="shadcn-card-description">법인 카드 사용 내역을 조회합니다.</p>
        </div>
        
        <% If errorMsg <> "" Then %>
        <div class="shadcn-alert shadcn-alert-error">
            <div>
                <span class="shadcn-alert-title">오류</span>
                <span class="shadcn-alert-description"><%= errorMsg %></span>
            </div>
        </div>
        <% End If %>
        
        <% If successMsg <> "" Then %>
        <div class="shadcn-alert shadcn-alert-success">
            <div>
                <span class="shadcn-alert-title">성공</span>
                <span class="shadcn-alert-description"><%= successMsg %></span>
            </div>
        </div>
        <% End If %>
        
        <!-- 검색 폼 -->
        <div class="shadcn-card-content">
            <form id="searchForm" method="get" action="card_usage.asp">
                <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 10px; margin-bottom: 15px;">
                    <div class="form-group">
                        <label class="shadcn-input-label" for="card_id">카드</label>
                        <select class="shadcn-select" id="card_id" name="card_id">
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
                    
                    <div class="form-group">
                        <label class="shadcn-input-label" for="account_type_id">계정 과목</label>
                        <select class="shadcn-select" id="account_type_id" name="account_type_id">
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
                    
                    <div class="form-group">
                        <label class="shadcn-input-label" for="start_date">시작일</label>
                        <input class="shadcn-input" type="date" id="start_date" name="start_date" value="<%= searchStartDate %>">
                    </div>
                    
                    <div class="form-group">
                        <label class="shadcn-input-label" for="end_date">종료일</label>
                        <input class="shadcn-input" type="date" id="end_date" name="end_date" value="<%= searchEndDate %>">
                    </div>
                </div>
                
                <div style="display: flex; justify-content: flex-end; gap: 10px;">
                    <button type="submit" class="shadcn-btn shadcn-btn-primary">검색</button>
                    <a href="card_usage.asp" class="shadcn-btn shadcn-btn-outline">초기화</a>
                    <a href="card_usage_add.asp" class="shadcn-btn shadcn-btn-secondary">새 내역 등록</a>
                </div>
            </form>
        </div>
    </div>
    
    <!-- 카드 사용 내역 목록 -->
    <div class="shadcn-card">
        <div class="shadcn-card-content">
            <% 
            ' 디버깅용 DB 연결 상태와 에러 표시
            If Not dbConnected Then 
            %>
                <div class="shadcn-alert shadcn-alert-error">
                    <div>
                        <span class="shadcn-alert-title">데이터베이스 연결 오류</span>
                        <span class="shadcn-alert-description"><%= dbErrorMsg %></span>
                    </div>
                </div>
            <% End If %>
            
            <% If Err.Number <> 0 Then %>
                <div class="shadcn-alert shadcn-alert-error">
                    <div>
                        <span class="shadcn-alert-title">SQL 오류 (번호: <%= Err.Number %>)</span>
                        <span class="shadcn-alert-description"><%= Err.Description %></span>
                    </div>
                </div>
            <% End If %>
            
            <table class="shadcn-table">
                <thead class="shadcn-table-header">
                    <tr>
                        <th>사용일자</th>
                        <th>카드</th>
                        <th>계정 과목</th>
                        <th>제목</th>
                        <th>사용처</th>
                        <th>사용 목적</th>
                        <th>금액</th>
                        <th>상태</th>
                        <th>관리</th>
                    </tr>
                </thead>
                <tbody>
                    <% 
                    If Not rs.EOF Then
                        Do While Not rs.EOF 
                    %>
                    <tr>
                        <td><%= FormatDate(rs("usage_date")) %></td>
                        <td><%= GetCardName(rs("card_id")) %></td>
                        <td><%= GetExpenseCategoryName(rs("expense_category_id")) %></td>
                        <td><% 
                            If Not IsNull(rs("title")) Then 
                                Response.Write(rs("title"))
                            ElseIf Not IsNull(rs("store_name")) Then
                                Response.Write(rs("store_name"))
                            Else
                                Response.Write("-")
                            End If
                        %></td>
                        <td><% 
                            If Not IsNull(rs("store_name")) Then 
                                Response.Write(rs("store_name"))
                            Else
                                Response.Write("-")
                            End If
                        %></td>
                        <td><% 
                            If Not IsNull(rs("purpose")) Then 
                                Response.Write(rs("purpose"))
                            Else
                                Response.Write("-")
                            End If
                        %></td>
                        <td><%= FormatNumber(rs("amount")) %></td>
                        <td><% 
                            If Not IsNull(rs("approval_status")) Then 
                                Response.Write(rs("approval_status"))
                            Else
                                Response.Write("처리중")
                            End If
                        %></td>
                        <td>
                            <div style="display: flex; gap: 5px;">
                                <a href="card_usage_view.asp?id=<%= rs("usage_id") %>" class="shadcn-btn shadcn-btn-outline" style="padding: 2px 8px; font-size: 0.75rem;">상세</a>
                                <% If rs("approval_status") <> "완료" Then %>
                                <a href="card_usage_edit.asp?id=<%= rs("usage_id") %>" class="shadcn-btn shadcn-btn-secondary" style="padding: 2px 8px; font-size: 0.75rem;">수정</a>
                                <% End If %>
                            </div>
                        </td>
                    </tr>
                    <% 
                        rs.MoveNext
                        Loop 
                    Else 
                    %>
                    <tr>
                        <td colspan="9" class="text-center">등록된 카드 사용 내역이 없습니다.</td>
                    </tr>
                    <% End If %>
                </tbody>
            </table>
            
            <!-- 페이징 -->
            <% If totalPages > 1 Then %>
            <div style="display: flex; justify-content: center; margin-top: 20px;">
                <div class="pagination">
                    <% 
                    Dim pageStart, pageEnd, i
                    
                    ' 표시할 페이지 범위 설정
                    pageStart = currentPage - 5
                    If pageStart < 1 Then pageStart = 1
                    
                    pageEnd = pageStart + 9
                    If pageEnd > totalPages Then pageEnd = totalPages
                    
                    ' 페이지 링크 생성
                    If currentPage > 1 Then
                    %>
                        <a href="card_usage.asp?page=1&card_id=<%= searchCardId %>&start_date=<%= searchStartDate %>&end_date=<%= searchEndDate %>&account_type_id=<%= searchAccountType %>" class="shadcn-btn shadcn-btn-outline" style="padding: 5px 10px; margin: 0 2px;">처음</a>
                        <a href="card_usage.asp?page=<%= currentPage - 1 %>&card_id=<%= searchCardId %>&start_date=<%= searchStartDate %>&end_date=<%= searchEndDate %>&account_type_id=<%= searchAccountType %>" class="shadcn-btn shadcn-btn-outline" style="padding: 5px 10px; margin: 0 2px;">이전</a>
                    <% End If %>
                    
                    <% For i = pageStart To pageEnd %>
                        <% If i = CInt(currentPage) Then %>
                            <span class="shadcn-btn shadcn-btn-primary" style="padding: 5px 10px; margin: 0 2px;"><%= i %></span>
                        <% Else %>
                            <a href="card_usage.asp?page=<%= i %>&card_id=<%= searchCardId %>&start_date=<%= searchStartDate %>&end_date=<%= searchEndDate %>&account_type_id=<%= searchAccountType %>" class="shadcn-btn shadcn-btn-outline" style="padding: 5px 10px; margin: 0 2px;"><%= i %></a>
                        <% End If %>
                    <% Next %>
                    
                    <% If CInt(currentPage) < totalPages Then %>
                        <a href="card_usage.asp?page=<%= currentPage + 1 %>&card_id=<%= searchCardId %>&start_date=<%= searchStartDate %>&end_date=<%= searchEndDate %>&account_type_id=<%= searchAccountType %>" class="shadcn-btn shadcn-btn-outline" style="padding: 5px 10px; margin: 0 2px;">다음</a>
                        <a href="card_usage.asp?page=<%= totalPages %>&card_id=<%= searchCardId %>&start_date=<%= searchStartDate %>&end_date=<%= searchEndDate %>&account_type_id=<%= searchAccountType %>" class="shadcn-btn shadcn-btn-outline" style="padding: 5px 10px; margin: 0 2px;">마지막</a>
                    <% End If %>
                </div>
            </div>
            <% End If %>
        </div>
    </div>
</div>

<!--#include file="../includes/footer.asp"--> 