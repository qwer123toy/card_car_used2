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
    RedirectTo("../index.asp")
End If

On Error Resume Next

' 카드 계정 목록 조회
Dim cardSQL, cardRS
cardSQL = "SELECT card_id, account_name FROM CardAccount ORDER BY account_name"
Set cardRS = db.Execute(cardSQL)

' 테이블이 없는 경우 빈 레코드셋 생성
If Err.Number <> 0 Then
    Set cardRS = Server.CreateObject("ADODB.Recordset")
    cardRS.Fields.Append "card_id", 3 ' adInteger
    cardRS.Fields.Append "account_name", 200, 100 ' adVarChar
    cardRS.Open
    
    ' 샘플 데이터 추가
    cardRS.AddNew
    cardRS("card_id") = 1
    cardRS("account_name") = "법인카드1"
    cardRS.Update
    
    cardRS.AddNew
    cardRS("card_id") = 2
    cardRS("account_name") = "법인카드2"
    cardRS.Update
    
    cardRS.MoveFirst
End If

' 계정 과목 목록 조회
Dim accountTypeSQL, accountTypeRS
accountTypeSQL = "SELECT account_type_id, type_name FROM CardAccountTypes ORDER BY type_name"
Set accountTypeRS = db.Execute(accountTypeSQL)

' 테이블이 없는 경우 빈 레코드셋 생성
If Err.Number <> 0 Then
    Set accountTypeRS = Server.CreateObject("ADODB.Recordset")
    accountTypeRS.Fields.Append "account_type_id", 3 ' adInteger
    accountTypeRS.Fields.Append "type_name", 200, 50 ' adVarChar
    accountTypeRS.Open
    
    ' 샘플 데이터 추가
    accountTypeRS.AddNew
    accountTypeRS("account_type_id") = 1
    accountTypeRS("type_name") = "식대"
    accountTypeRS.Update
    
    accountTypeRS.AddNew
    accountTypeRS("account_type_id") = 2
    accountTypeRS("type_name") = "교통비"
    accountTypeRS.Update
    
    accountTypeRS.AddNew
    accountTypeRS("account_type_id") = 3
    accountTypeRS("type_name") = "접대비"
    accountTypeRS.Update
    
    accountTypeRS.MoveFirst
End If

' 카드 사용 등록 처리
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    Dim cardId, usageDate, amount, purpose, accountTypeId, errorMsg, successMsg, storeName
    
    cardId = PreventSQLInjection(Request.Form("card_id"))
    usageDate = PreventSQLInjection(Request.Form("usage_date"))
    amount = PreventSQLInjection(Request.Form("amount"))
    purpose = PreventSQLInjection(Request.Form("purpose"))
    accountTypeId = PreventSQLInjection(Request.Form("account_type_id"))
    storeName = PreventSQLInjection(Request.Form("store_name"))
    
    ' 금액에서 콤마 제거
    amount = Replace(amount, ",", "")
    
    ' 입력값 검증
    If cardId = "" Or usageDate = "" Or amount = "" Or accountTypeId = "" Or storeName = "" Then
        errorMsg = "필수 항목을 모두 입력해주세요."
    ElseIf Not IsNumeric(amount) Then
        errorMsg = "금액은 숫자만 입력 가능합니다."
    Else
        ' SQL 쿼리 수정 - department_id 필드 추가
        Dim cmd, departmentId
        Set cmd = Server.CreateObject("ADODB.Command")
        cmd.ActiveConnection = db
        
        ' 사용자의 부서 ID 가져오기 (세션에서)
        If Session("department_id") <> "" Then
            departmentId = Session("department_id")
        Else
            departmentId = 1 ' 기본값 설정 (관리부)
        End If
        
        ' SQL 명령문 - 실제 DB 구조에 맞게 수정
        cmd.CommandText = "INSERT INTO CardUsage (user_id, card_id, usage_date, amount, expense_category_id, department_id, store_name, purpose, approval_status) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)"
        
        ' 파라미터 추가
        cmd.Parameters.Append cmd.CreateParameter("@user_id", 200, 1, 30, Session("user_id"))
        cmd.Parameters.Append cmd.CreateParameter("@card_id", 3, 1, , cardId)
        cmd.Parameters.Append cmd.CreateParameter("@usage_date", 7, 1, , usageDate)
        cmd.Parameters.Append cmd.CreateParameter("@amount", 6, 1, , CDbl(amount))
        cmd.Parameters.Append cmd.CreateParameter("@expense_category_id", 3, 1, , accountTypeId)
        cmd.Parameters.Append cmd.CreateParameter("@department_id", 3, 1, , departmentId)
        cmd.Parameters.Append cmd.CreateParameter("@store_name", 200, 1, 100, storeName)
        cmd.Parameters.Append cmd.CreateParameter("@purpose", 200, 1, 200, purpose)
        cmd.Parameters.Append cmd.CreateParameter("@approval_status", 200, 1, 20, "대기")
        
        ' 쿼리 정보 로깅
        LogActivity Session("user_id"), "SQL실행", "카드 사용 내역 등록: " & cmd.CommandText
        
        ' 명령 실행
        On Error Resume Next
        cmd.Execute
        
        If Err.Number <> 0 Then
            errorMsg = "카드 사용 내역 등록 중 오류가 발생했습니다: " & Err.Description
            
            ' 오류 로깅
            LogActivity Session("user_id"), "오류발생", "오류코드: " & Err.Number & ", 설명: " & Err.Description
        Else
            successMsg = "카드 사용 내역이 성공적으로 등록되었습니다."
            
            ' 활동 로그 기록
            LogActivity Session("user_id"), "카드사용등록", "카드 사용 내역 등록 (금액: " & amount & "원, 사유: " & purpose & ")"
            
            ' dashboard 페이지로 리다이렉트
            Response.Redirect("dashboard.asp")
        End If
        On Error GoTo 0
    End If
End If

On Error GoTo 0
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
            <form id="cardUsageForm" method="post" action="card_usage_add.asp" onsubmit="prepareFormSubmission(); return validateForm('cardUsageForm', cardUsageRules)">
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
                    <label class="shadcn-input-label" for="store_name">사용처</label>
                    <input class="shadcn-input" type="text" id="store_name" name="store_name" placeholder="사용처를 입력하세요">
                </div>
                
                <div class="form-group">
                    <label class="shadcn-input-label" for="amount">금액</label>
                    <input class="shadcn-input" type="text" id="amount" name="amount" placeholder="금액을 입력하세요" onkeyup="cleanNumberInput(this)">
                </div>
                
                <div class="form-group">
                    <label class="shadcn-input-label" for="purpose">사용 목적</label>
                    <textarea class="shadcn-input" id="purpose" name="purpose" rows="3" placeholder="사용 목적을 입력하세요"></textarea>
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
        
        // 쉼표(,) 제거 및 숫자만 남기기
        let value = input.value.replace(/,/g, '');
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
</script>

<!--#include file="../includes/footer.asp"--> 