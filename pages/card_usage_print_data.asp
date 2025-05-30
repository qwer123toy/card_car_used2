<%@ Language="VBScript" CodePage="65001" %>
<% 
Response.CodePage = 65001
Response.CharSet = "utf-8"
%>

<!--#include file="../db.asp"-->
<!--#include file="../includes/functions.asp"-->
<%
' 로그인 체크
If Not IsAuthenticated() Then
    Response.Write "로그인이 필요합니다."
    Response.End
End If

' 검색 조건 가져오기
Dim searchCardId, searchStartDate, searchEndDate, searchAccountType, searchCondition
searchCardId = PreventSQLInjection(Request.QueryString("card_id"))
searchStartDate = PreventSQLInjection(Request.QueryString("start_date"))
searchEndDate = PreventSQLInjection(Request.QueryString("end_date"))
searchAccountType = PreventSQLInjection(Request.QueryString("account_type_id"))

' 검색 조건 SQL 생성
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

' 전체 데이터 조회 (페이징 없이)
Dim SQL, rs
SQL = "SELECT usage_id, user_id, card_id, usage_date, amount, store_name, purpose, title, " & _
      "approval_status, department_id, expense_category_id, cost_type_id " & _
      "FROM " & dbSchema & ".CardUsage" & searchCondition & _
      " ORDER BY usage_date DESC, usage_id DESC"

Set rs = db.Execute(SQL)

' 카드 목록 조회
Dim cardSQL, cardRS
cardSQL = "SELECT card_id, account_name, issuer FROM " & dbSchema & ".CardAccount ORDER BY account_name"
Set cardRS = db99.Execute(cardSQL)

' 계정 과목 목록 조회
Dim accountTypeSQL, accountTypeRS
accountTypeSQL = "SELECT account_type_id, type_name FROM " & dbSchema & ".CardAccountTypes ORDER BY type_name"
Set accountTypeRS = db.Execute(accountTypeSQL)

' 카드 계정 이름 조회 함수
Function GetCardNameForPrint(cardId)
    Dim cardNumber, cardName
    cardNumber = "알 수 없음"
    cardName = "알 수 없음"
    
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
    
    GetCardNameForPrint = cardName & " (" & cardNumber & ")"
End Function

' 계정 유형 이름 조회 함수
Function GetExpenseCategoryNameForPrint(categoryId)
    If categoryId = "" Or Not IsNumeric(categoryId) Then
        GetExpenseCategoryNameForPrint = "-"
        Exit Function
    End If
    
    Dim typeName
    typeName = "-"
    
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
    
    GetExpenseCategoryNameForPrint = typeName
End Function

' 총 건수 계산
Dim totalCount
totalCount = 0
If Not rs.EOF Then
    rs.MoveLast
    totalCount = rs.RecordCount
    rs.MoveFirst
End If
%>

<div class="table-container">
    <div class="total-info">
        <p><strong>총 건수:</strong> <%= totalCount %>건</p>
    </div>
    
    <% If rs.EOF Then %>
        <div class="empty-message">
            <p>검색 조건에 해당하는 카드 사용 내역이 없습니다.</p>
        </div>
    <% Else %>
        <table class="table">
            <thead>
                <tr>
                    <th>사용일자</th>
                    <th>카드</th>
                    <th>계정 과목</th>
                    <th>제목</th>
                    <th>사용처</th>
                    <th>사용 목적</th>
                    <th>금액</th>
                    <th>상태</th>
                </tr>
            </thead>
            <tbody>
                <% Do While Not rs.EOF %>
                <tr>
                    <td class="date-cell"><%= FormatDate(rs("usage_date")) %></td>
                    <td><%= GetCardNameForPrint(rs("card_id")) %></td>
                    <td><%= GetExpenseCategoryNameForPrint(rs("expense_category_id")) %></td>
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
                    <td class="amount-cell"><%= FormatNumber(rs("amount")) %>원</td>
                    <td>
                        <% 
                        Dim statusText
                        If Not IsNull(rs("approval_status")) Then 
                            statusText = rs("approval_status")
                        Else
                            statusText = "처리중"
                        End If
                        %>
                        <span class="status-badge"><%= statusText %></span>
                    </td>
                </tr>
                <% 
                    rs.MoveNext
                    Loop 
                %>
            </tbody>
        </table>
    <% End If %>
</div>

<%
' 사용한 객체 해제
If Not rs Is Nothing Then
    If rs.State = 1 Then
        rs.Close
    End If
    Set rs = Nothing
End If

If Not cardRS Is Nothing Then
    If cardRS.State = 1 Then
        cardRS.Close
    End If
    Set cardRS = Nothing
End If

If Not accountTypeRS Is Nothing Then
    If accountTypeRS.State = 1 Then
        accountTypeRS.Close
    End If
    Set accountTypeRS = Nothing
End If
%> 