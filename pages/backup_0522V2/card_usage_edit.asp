<%@ Language="VBScript" CodePage="65001" %>
<% 
Response.CodePage = 65001
Response.CharSet = "utf-8"
%>

<!--#include file="../db.asp"-->
<!--#include file="../includes/functions.asp"-->
<%
' 숫자를 한국 통화 형식으로 변환하는 사용자 정의 함수
Function FormatKoreanCurrency(value)
    If IsNull(value) Then
        FormatKoreanCurrency = "0"
        Exit Function
    End If
    
    ' 숫자를 문자열로 변환하고 천 단위 콤마 추가
    If IsNumeric(value) Then
        FormatKoreanCurrency = FormatNumber(value, 0, -1, -1, -1)
    Else
        FormatKoreanCurrency = "0"
    End If
End Function

' 로그인 체크
If Not IsAuthenticated() Then
    RedirectTo("../index.asp")
End If

On Error Resume Next

' URL 파라미터에서 ID 가져오기
Dim usageId, errorMsg, successMsg
usageId = PreventSQLInjection(Request.QueryString("id"))

If usageId = "" Then
    errorMsg = "잘못된 접근입니다. 카드 사용 내역 ID가 필요합니다."
    Response.Write("<script>alert('" & errorMsg & "');window.location.href='card_usage.asp';</script>")
    Response.End
End If

' dbSchema가 설정되지 않은 경우를 대비해 기본값 설정
If Not(IsObject(dbSchema)) And (TypeName(dbSchema) <> "String" Or Len(dbSchema) = 0) Then
    dbSchema = "dbo"
End If

' 카드 사용 내역 조회
Dim SQL, rs
SQL = "SELECT u.usage_id, u.user_id, u.card_id, u.usage_date, u.amount, u.store_name, " & _
      "u.purpose, u.approval_status, u.department_id, u.expense_category_id, u.title " & _
      "FROM " & dbSchema & ".CardUsage u " & _
      "WHERE u.usage_id = " & usageId

Set rs = db.Execute(SQL)

If Err.Number <> 0 Then
    Response.Write("데이터베이스 조회 오류: " & Err.Description)
    Response.End
End If

' 레코드가 없는 경우 처리
If rs.EOF Then
    Response.Write("<script>alert('해당 카드 사용 내역을 찾을 수 없습니다.');window.location.href='card_usage.asp';</script>")
    Response.End
End If

' 상태가 완료일 때 접근 거부
If rs("approval_status") = "완료" Then
    Response.Write("<script>alert('완료 상태의 카드 사용 내역은 수정할 수 없습니다.');window.location.href='card_usage.asp';</script>")
    Response.End
End If

' 카드 계정 목록 조회
Dim cardSQL, cardRS
cardSQL = "SELECT card_id, account_name FROM " & dbSchema & ".CardAccount ORDER BY account_name"
Set cardRS = db.Execute(cardSQL)

' 계정 과목 목록 조회
Dim accountTypeSQL, accountTypeRS
accountTypeSQL = "SELECT account_type_id, type_name FROM " & dbSchema & ".CardAccountTypes ORDER BY type_name"
Set accountTypeRS = db.Execute(accountTypeSQL)

' 폼 제출 처리
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    Dim cardId, usageDate, amount, purpose, expenseCategoryId, storeName, title
    
    cardId = PreventSQLInjection(Request.Form("card_id"))
    usageDate = PreventSQLInjection(Request.Form("usage_date"))
    amount = PreventSQLInjection(Request.Form("amount"))
    purpose = PreventSQLInjection(Request.Form("purpose"))
    expenseCategoryId = PreventSQLInjection(Request.Form("expense_category_id"))
    storeName = PreventSQLInjection(Request.Form("store_name"))
    title = PreventSQLInjection(Request.Form("title"))
    
    ' 금액에서 콤마 제거
    amount = Replace(amount, ",", "")
    
    ' 입력값 검증
    If cardId = "" Or usageDate = "" Or amount = "" Or expenseCategoryId = "" Or storeName = "" Or title = "" Then
        errorMsg = "필수 항목을 모두 입력해주세요."
    ElseIf Not IsNumeric(amount) Then
        errorMsg = "금액은 숫자만 입력 가능합니다."
    Else
        ' SQL 쿼리 수정
        Dim cmd
        Set cmd = Server.CreateObject("ADODB.Command")
        cmd.ActiveConnection = db
        
        ' SQL 명령문 - title 필드 추가
        cmd.CommandText = "UPDATE " & dbSchema & ".CardUsage SET card_id = ?, usage_date = ?, amount = ?, " & _
                         "expense_category_id = ?, store_name = ?, purpose = ?, title = ? " & _
                         "WHERE usage_id = ?"
        
        ' 파라미터 추가
        cmd.Parameters.Append cmd.CreateParameter("@card_id", 3, 1, , cardId)
        cmd.Parameters.Append cmd.CreateParameter("@usage_date", 7, 1, , usageDate)
        cmd.Parameters.Append cmd.CreateParameter("@amount", 6, 1, , CDbl(amount))
        cmd.Parameters.Append cmd.CreateParameter("@expense_category_id", 3, 1, , expenseCategoryId)
        cmd.Parameters.Append cmd.CreateParameter("@store_name", 200, 1, 100, storeName)
        cmd.Parameters.Append cmd.CreateParameter("@purpose", 200, 1, 200, purpose)
        cmd.Parameters.Append cmd.CreateParameter("@title", 200, 1, 200, title)
        cmd.Parameters.Append cmd.CreateParameter("@usage_id", 3, 1, , usageId)
        
        On Error Resume Next
        cmd.Execute
        
        If Err.Number <> 0 Then
            errorMsg = "카드 사용 내역 수정 중 오류가 발생했습니다: " & Err.Description
            LogActivity Session("user_id"), "오류발생", "오류코드: " & Err.Number & ", 설명: " & Err.Description
        Else
            successMsg = "카드 사용 내역이 성공적으로 수정되었습니다."
            LogActivity Session("user_id"), "카드사용수정", "카드 사용 내역 수정 (금액: " & amount & "원, 사유: " & purpose & ")"
            Response.Redirect("card_usage_view.asp?id=" & usageId)
        End If
        On Error GoTo 0
    End If
End If

On Error GoTo 0
%>
<!--#include file="../includes/header.asp"-->

<div class="card-usage-edit-container">
    <div class="shadcn-card" style="max-width: 700px; margin: 30px auto;">
        <div class="shadcn-card-header">
            <h2 class="shadcn-card-title">카드 사용 내역 수정</h2>
            <p class="shadcn-card-description">법인 카드 사용 내역을 수정합니다.</p>
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
            <form id="cardUsageForm" method="post" action="card_usage_edit.asp?id=<%= usageId %>" onsubmit="prepareFormSubmission(); return validateForm('cardUsageForm', cardUsageRules)">
                <div class="form-group">
                    <label class="shadcn-input-label" for="card_id">카드 선택</label>
                    <select class="shadcn-select" id="card_id" name="card_id">
                        <option value="">카드를 선택하세요</option>
                        <% 
                        ' 실제 데이터베이스에서 카드 정보 표시 
                        If cardRS.RecordCount > 0 Then
                            ' 실제 데이터베이스에서 가져온 카드 계정 목록 표시
                            cardRS.MoveFirst
                            Do While Not cardRS.EOF 
                                Dim cardSelected
                                cardSelected = ""
                                If CStr(cardRS("card_id")) = CStr(rs("card_id")) Then
                                    cardSelected = "selected"
                                End If
                        %>
                            <option value="<%= cardRS("card_id") %>" <%= cardSelected %>><%= cardRS("account_name") %></option>
                        <% 
                                cardRS.MoveNext
                            Loop
                        Else 
                            ' 데이터베이스에 데이터가 없는 경우 기본 옵션
                        %>
                            <option value="1">법인카드1</option>
                            <option value="2">법인카드2</option>
                        <% End If %>
                    </select>
                </div>
                
                <div class="form-group">
                    <label class="shadcn-input-label" for="expense_category_id">계정 과목</label>
                    <select class="shadcn-select" id="expense_category_id" name="expense_category_id">
                        <option value="">계정 과목을 선택하세요</option>
                        <% 
                        ' 실제 데이터베이스에서 계정 과목 정보 표시
                        If accountTypeRS.RecordCount > 0 Then
                            ' 실제 데이터베이스에서 가져온 계정 과목 목록 표시
                            accountTypeRS.MoveFirst
                            Do While Not accountTypeRS.EOF 
                                Dim typeSelected
                                typeSelected = ""
                                If CStr(accountTypeRS("account_type_id")) = CStr(rs("expense_category_id")) Then
                                    typeSelected = "selected"
                                End If
                        %>
                            <option value="<%= accountTypeRS("account_type_id") %>" <%= typeSelected %>><%= accountTypeRS("type_name") %></option>
                        <% 
                                accountTypeRS.MoveNext
                            Loop
                        Else 
                            ' 데이터베이스에 데이터가 없는 경우 기본 옵션
                        %>
                            <option value="1">식대</option>
                            <option value="2">교통비</option>
                            <option value="3">접대비</option>
                        <% End If %>
                    </select>
                </div>
                
                <div class="form-group">
                    <label class="shadcn-input-label" for="usage_date">사용일자</label>
                    <input class="shadcn-input" type="date" id="usage_date" name="usage_date" value="<%= FormatDateTime(rs("usage_date"), 2) %>">
                </div>
                
                <div class="form-group">
                    <label class="shadcn-input-label" for="title">제목</label>
                    <input class="shadcn-input" type="text" id="title" name="title" placeholder="제목을 입력하세요" value="<%= rs("title") %>">
                </div>
                
                <div class="form-group">
                    <label class="shadcn-input-label" for="store_name">사용처</label>
                    <input class="shadcn-input" type="text" id="store_name" name="store_name" placeholder="사용처를 입력하세요" value="<%= rs("store_name") %>">
                </div>
                
                <div class="form-group">
                    <label class="shadcn-input-label" for="amount">금액</label>
                    <input class="shadcn-input" type="text" id="amount" name="amount" placeholder="금액을 입력하세요" value="<%= rs("amount") %>" onkeyup="cleanNumberInput(this)">
                </div>
                
                <div class="form-group">
                    <label class="shadcn-input-label" for="purpose">사용 목적</label>
                    <textarea class="shadcn-input" id="purpose" name="purpose" rows="3" placeholder="사용 목적을 입력하세요"><%= rs("purpose") %></textarea>
                </div>
                
                <div class="shadcn-card-footer" style="margin-top: 1.5rem;">
                    <button type="submit" class="shadcn-btn shadcn-btn-primary">수정하기</button>
                    <a href="card_usage_view.asp?id=<%= usageId %>" class="shadcn-btn shadcn-btn-outline">취소</a>
                    <a href="card_usage.asp" class="shadcn-btn shadcn-btn-outline">목록</a>
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
        expense_category_id: {
            required: true,
            message: '계정 과목을 선택해주세요.'
        },
        usage_date: {
            required: true,
            message: '사용일자를 입력해주세요.'
        },
        store_name: {
            required: true,
            message: '사용처를 입력해주세요.'
        },
        amount: {
            required: true,
            numeric: true,
            message: '금액을 숫자로 입력해주세요.'
        },
        purpose: {
            required: true,
            message: '사용 목적을 입력해주세요.'
        }
    };
    
    // 폼 제출 전 숫자 필드의 쉼표 제거
    function prepareFormSubmission() {
        // 숫자 입력 필드의 쉼표 제거
        const numericFields = ['amount'];
        numericFields.forEach(fieldId => {
            const field = document.getElementById(fieldId);
            if (field) {
                field.value = field.value.replace(/,/g, '');
            }
        });
    }
    
    // 숫자 입력 필드에서 쉼표 제거하는 함수
    function cleanNumberInput(input) {
        // 현재 선택 위치 저장
        const start = input.selectionStart;
        const end = input.selectionEnd;
        
        // 통화 기호와 쉼표(,) 제거 및 숫자만 남기기
        let value = input.value.replace(/[₩\\₩₩₩₩,]/g, ''); // 원화 기호, \ 제거
        value = value.replace(/[^\d.]/g, ''); // 숫자와 마침표만 허용
        
        // 천 단위 콤마 추가
        if (value) {
            value = parseFloat(value).toLocaleString('ko-KR', {maximumFractionDigits: 0});
        }
        
        // 입력 값이 바뀌었는지 확인
        const hasChanged = input.value !== value;
        
        // 값 갱신
        input.value = value;
        
        // 선택 위치 복원 (값이 바뀌었을 경우)
        if (hasChanged) {
            // 콤마가 추가된 경우 위치 조정 필요
            const newCursorPos = Math.max(
                0,
                value.length - (input.value.length - end)
            );
            input.setSelectionRange(newCursorPos, newCursorPos);
        }
    }

    // 페이지 로드 시 금액 필드 초기화
    window.onload = function() {
        // 페이지 로드 시 자동으로 금액 필드 정리 (통화 기호 제거)
        const amountField = document.getElementById('amount');
        if (amountField) {
            // 기존 값 저장
            const originalValue = amountField.value;
            // 통화 기호 및 쉼표 처리
            cleanNumberInput(amountField);
        }
    };
</script>

<%
' 사용한 Recordset 닫기
If IsObject(cardRS) Then
    If cardRS.State = 1 Then ' adStateOpen
        cardRS.Close
    End If
    Set cardRS = Nothing
End If

If IsObject(accountTypeRS) Then
    If accountTypeRS.State = 1 Then ' adStateOpen
        accountTypeRS.Close
    End If
    Set accountTypeRS = Nothing
End If

If IsObject(rs) Then
    If rs.State = 1 Then ' adStateOpen
        rs.Close
    End If
    Set rs = Nothing
End If
%>

<!--#include file="../includes/footer.asp"--> 