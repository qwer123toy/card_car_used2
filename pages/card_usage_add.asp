<!--#include file="../includes/connection.asp"-->
<!--#include file="../includes/functions.asp"-->
<%
' 로그인 체크
If Not IsAuthenticated() Then
    RedirectTo("../index.asp")
End If

' 카드 계정 목록 조회
Dim cardSQL, cardRS
cardSQL = "SELECT card_id, account_name FROM CardAccount ORDER BY account_name"
Set cardRS = dbConn.Execute(cardSQL)

' 계정 과목 목록 조회
Dim accountTypeSQL, accountTypeRS
accountTypeSQL = "SELECT account_type_id, type_name FROM CardAccountTypes ORDER BY type_name"
Set accountTypeRS = dbConn.Execute(accountTypeSQL)

' 카드 사용 등록 처리
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    Dim cardId, usageDate, amount, usageReason, accountTypeId, errorMsg, successMsg, insertSQL
    
    cardId = PreventSQLInjection(Request.Form("card_id"))
    usageDate = PreventSQLInjection(Request.Form("usage_date"))
    amount = PreventSQLInjection(Request.Form("amount"))
    usageReason = PreventSQLInjection(Request.Form("usage_reason"))
    accountTypeId = PreventSQLInjection(Request.Form("account_type_id"))
    
    ' 금액에서 콤마 제거
    amount = Replace(amount, ",", "")
    
    ' 입력값 검증
    If cardId = "" Or usageDate = "" Or amount = "" Or accountTypeId = "" Then
        errorMsg = "필수 항목을 모두 입력해주세요."
    ElseIf Not IsNumeric(amount) Then
        errorMsg = "금액은 숫자만 입력 가능합니다."
    Else
        ' 카드 사용 내역 등록
        insertSQL = "INSERT INTO CardUsage (user_id, card_id, usage_date, amount, usage_reason, account_type_id) VALUES ('" & _
                   Session("user_id") & "', " & cardId & ", '" & usageDate & "', " & amount & ", '" & usageReason & "', " & accountTypeId & ")"
        
        On Error Resume Next
        dbConn.Execute insertSQL
        
        If Err.Number <> 0 Then
            errorMsg = "카드 사용 내역 등록 중 오류가 발생했습니다: " & Err.Description
        Else
            successMsg = "카드 사용 내역이 성공적으로 등록되었습니다."
            
            ' 활동 로그 기록
            LogActivity Session("user_id"), "카드사용등록", "카드 사용 내역 등록 (금액: " & amount & "원)"
        End If
        On Error GoTo 0
    End If
End If
%>
<!--#include file="../includes/header.asp"-->

<div class="card-usage-add-container">
    <div class="shadcn-card" style="max-width: 700px; margin: 30px auto;">
        <div class="shadcn-card-header">
            <h2 class="shadcn-card-title">카드 사용 내역 등록</h2>
            <p class="shadcn-card-description">법인 카드 사용 내역을 등록합니다.</p>
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
        
        <div class="shadcn-card-content">
            <form id="cardUsageForm" method="post" action="card_usage_add.asp" onsubmit="return validateForm('cardUsageForm', cardUsageRules)">
                <div class="form-group">
                    <label class="shadcn-input-label" for="card_id">카드 선택</label>
                    <select class="shadcn-select" id="card_id" name="card_id">
                        <option value="">카드를 선택하세요</option>
                        <% 
                        If Not cardRS.EOF Then
                            Do While Not cardRS.EOF 
                        %>
                            <option value="<%= cardRS("card_id") %>"><%= cardRS("account_name") %></option>
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
                        <option value="">계정 과목을 선택하세요</option>
                        <% 
                        If Not accountTypeRS.EOF Then
                            Do While Not accountTypeRS.EOF 
                        %>
                            <option value="<%= accountTypeRS("account_type_id") %>"><%= accountTypeRS("type_name") %></option>
                        <% 
                                accountTypeRS.MoveNext
                            Loop
                        End If
                        accountTypeRS.Close
                        %>
                    </select>
                </div>
                
                <div class="form-group">
                    <label class="shadcn-input-label" for="usage_date">사용일자</label>
                    <input class="shadcn-input" type="date" id="usage_date" name="usage_date">
                </div>
                
                <div class="form-group">
                    <label class="shadcn-input-label" for="amount">금액</label>
                    <input class="shadcn-input" type="text" id="amount" name="amount" placeholder="금액을 입력하세요" onkeyup="formatCurrency(this)">
                </div>
                
                <div class="form-group">
                    <label class="shadcn-input-label" for="usage_reason">사용 사유</label>
                    <textarea class="shadcn-input" id="usage_reason" name="usage_reason" rows="3" placeholder="사용 사유를 입력하세요"></textarea>
                </div>
                
                <div class="shadcn-card-footer" style="margin-top: 1.5rem;">
                    <button type="submit" class="shadcn-btn shadcn-btn-primary">등록하기</button>
                    <a href="card_usage.asp" class="shadcn-btn shadcn-btn-outline">취소</a>
                </div>
            </form>
        </div>
    </div>
</div>

<script>
    const cardUsageRules = {
        card_id: {
            required: true,
            message: '카드를 선택해주세요.'
        },
        account_type_id: {
            required: true,
            message: '계정 과목을 선택해주세요.'
        },
        usage_date: {
            required: true,
            message: '사용일자를 입력해주세요.'
        },
        amount: {
            required: true,
            message: '금액을 입력해주세요.'
        }
    };
</script>

<!--#include file="../includes/footer.asp"--> 