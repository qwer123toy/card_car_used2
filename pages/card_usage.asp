<!--#include file="../includes/connection.asp"-->
<!--#include file="../includes/functions.asp"-->
<%
' 로그인 체크
If Not IsAuthenticated() Then
    RedirectTo("../index.asp")
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
searchCondition = " WHERE cu.user_id = '" & Session("user_id") & "' "

If searchCardId <> "" Then
    searchCondition = searchCondition & " AND cu.card_id = " & searchCardId
End If

If searchStartDate <> "" Then
    searchCondition = searchCondition & " AND cu.usage_date >= '" & searchStartDate & "'"
End If

If searchEndDate <> "" Then
    searchCondition = searchCondition & " AND cu.usage_date <= '" & searchEndDate & "'"
End If

If searchAccountType <> "" Then
    searchCondition = searchCondition & " AND cu.account_type_id = " & searchAccountType
End If

' 총 레코드 수 조회
Dim countSQL, countRS
countSQL = "SELECT COUNT(*) AS total FROM CardUsage cu" & searchCondition
Set countRS = dbConn.Execute(countSQL)
totalRows = countRS("total")
countRS.Close

' 전체 페이지 수 계산
Dim totalPages
totalPages = Ceil(totalRows / pageSize)
If totalPages < 1 Then totalPages = 1

' 카드 사용 내역 조회
Dim SQL, rs
SQL = "SELECT cu.usage_id, cu.usage_date, cu.amount, cu.usage_reason, " & _
      "ca.account_name, cat.type_name " & _
      "FROM CardUsage cu " & _
      "JOIN CardAccount ca ON cu.card_id = ca.card_id " & _
      "JOIN CardAccountTypes cat ON cu.account_type_id = cat.account_type_id " & _
      searchCondition & " " & _
      "ORDER BY cu.usage_date DESC, cu.usage_id DESC " & _
      "OFFSET " & startRow & " ROWS FETCH NEXT " & pageSize & " ROWS ONLY"

Set rs = dbConn.Execute(SQL)

' 페이지네이션 함수
Function Ceil(number)
    Ceil = Int(number)
    If Ceil <> number Then
        Ceil = Ceil + 1
    End If
End Function

' 카드 목록 조회
Dim cardSQL, cardRS
cardSQL = "SELECT card_id, account_name FROM CardAccount ORDER BY account_name"
Set cardRS = dbConn.Execute(cardSQL)

' 계정 과목 목록 조회
Dim accountTypeSQL, accountTypeRS
accountTypeSQL = "SELECT account_type_id, type_name FROM CardAccountTypes ORDER BY type_name"
Set accountTypeRS = dbConn.Execute(accountTypeSQL)
%>
<!--#include file="../includes/header.asp"-->

<div class="card-usage-container">
    <div class="shadcn-card" style="margin-bottom: 20px;">
        <div class="shadcn-card-header">
            <h2 class="shadcn-card-title">카드 사용 내역</h2>
            <p class="shadcn-card-description">법인 카드 사용 내역을 조회합니다.</p>
        </div>
        
        <!-- 검색 폼 -->
        <div class="shadcn-card-content">
            <form id="searchForm" method="get" action="card_usage.asp">
                <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 10px; margin-bottom: 15px;">
                    <div class="form-group">
                        <label class="shadcn-input-label" for="card_id">카드</label>
                        <select class="shadcn-select" id="card_id" name="card_id">
                            <option value="">전체</option>
                            <% 
                            If Not cardRS.EOF Then
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
                            cardRS.Close
                            %>
                        </select>
                    </div>
                    
                    <div class="form-group">
                        <label class="shadcn-input-label" for="account_type_id">계정 과목</label>
                        <select class="shadcn-select" id="account_type_id" name="account_type_id">
                            <option value="">전체</option>
                            <% 
                            If Not accountTypeRS.EOF Then
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
                            accountTypeRS.Close
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
            <% If totalRows = 0 Then %>
                <div style="text-align: center; padding: 30px;">
                    <p>등록된 카드 사용 내역이 없습니다.</p>
                </div>
            <% Else %>
                <table class="shadcn-table">
                    <thead class="shadcn-table-header">
                        <tr>
                            <th>사용일자</th>
                            <th>카드</th>
                            <th>계정 과목</th>
                            <th>금액</th>
                            <th>사용 사유</th>
                            <th>관리</th>
                        </tr>
                    </thead>
                    <tbody>
                        <% Do While Not rs.EOF %>
                        <tr>
                            <td><%= FormatDate(rs("usage_date")) %></td>
                            <td><%= rs("account_name") %></td>
                            <td><%= rs("type_name") %></td>
                            <td><%= FormatNumber(rs("amount")) %></td>
                            <td><%= rs("usage_reason") %></td>
                            <td>
                                <div style="display: flex; gap: 5px;">
                                    <a href="card_usage_print.asp?id=<%= rs("usage_id") %>" class="shadcn-btn shadcn-btn-outline" style="padding: 2px 8px; font-size: 0.75rem;">출력</a>
                                    <a href="card_usage_edit.asp?id=<%= rs("usage_id") %>" class="shadcn-btn shadcn-btn-secondary" style="padding: 2px 8px; font-size: 0.75rem;">수정</a>
                                </div>
                            </td>
                        </tr>
                        <% 
                            rs.MoveNext
                            Loop 
                        %>
                    </tbody>
                </table>
            <% End If %>
            
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