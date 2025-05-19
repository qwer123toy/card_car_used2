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
cardSQL = "SELECT card_id, account_name FROM " & dbSchema & ".CardAccount ORDER BY account_name"
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
accountTypeSQL = "SELECT account_type_id, type_name FROM " & dbSchema & ".CardAccountTypes ORDER BY type_name"
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

' 에러 메시지 및 성공 메시지 초기화
Dim errorMsg, successMsg

' 카드 사용 등록 처리
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    Dim cardId, usageDate, amount, purpose, accountTypeId, storeName
    Dim approver1, approver2, approver3
    
    cardId = PreventSQLInjection(Request.Form("card_id"))
    usageDate = PreventSQLInjection(Request.Form("usage_date"))
    amount = PreventSQLInjection(Request.Form("amount"))
    purpose = PreventSQLInjection(Request.Form("purpose"))
    accountTypeId = PreventSQLInjection(Request.Form("account_type_id"))
    storeName = PreventSQLInjection(Request.Form("store_name"))
    approver1 = PreventSQLInjection(Request.Form("approver_step1"))
    approver2 = PreventSQLInjection(Request.Form("approver_step2"))
    approver3 = PreventSQLInjection(Request.Form("approver_step3"))
    
    ' 금액에서 콤마 제거
    amount = Replace(amount, ",", "")
    
    ' 입력값 검증
    If cardId = "" Or usageDate = "" Or amount = "" Or accountTypeId = "" Or storeName = "" Or approver1 = "" Then
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
        
        ' 트랜잭션 시작
        db.BeginTrans
        
        On Error Resume Next
        
        ' SQL 명령문 - 실제 DB 구조에 맞게 수정
        cmd.CommandText = "INSERT INTO " & dbSchema & ".CardUsage (user_id, card_id, usage_date, amount, expense_category_id, department_id, store_name, purpose, approval_status) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)"
        
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
        
        cmd.Execute
        
        If Err.Number = 0 Then
            ' 방금 삽입된 CardUsage의 ID 가져오기
            Dim usageId
            Set cmd = Server.CreateObject("ADODB.Command")
            cmd.ActiveConnection = db
            cmd.CommandText = "SELECT IDENT_CURRENT('CardUsage') AS usage_id"
            Dim rs
            Set rs = cmd.Execute()
            usageId = rs("usage_id")
            
            ' ApprovalLogs 테이블에 결재선 데이터 삽입
            Dim approverIds(2)
            approverIds(0) = approver1
            approverIds(1) = approver2
            approverIds(2) = approver3
            
            Dim i
            For i = 0 To 2
                If approverIds(i) <> "" Then
                    Dim approvalSQL
                    approvalSQL = "INSERT INTO " & dbSchema & ".ApprovalLogs " & _
                                "(approver_id, target_table_name, target_id, approval_step, status, created_at) " & _
                                "VALUES (?, 'CardUsage', ?, ?, '대기', GETDATE())"
                    
                    Set cmd = Server.CreateObject("ADODB.Command")
                    cmd.ActiveConnection = db
                    cmd.CommandText = approvalSQL
                    cmd.Parameters.Append cmd.CreateParameter("@approver_id", 200, 1, 30, approverIds(i))
                    cmd.Parameters.Append cmd.CreateParameter("@target_id", 3, 1, , usageId)
                    cmd.Parameters.Append cmd.CreateParameter("@approval_step", 3, 1, , i + 1)
                    
                    cmd.Execute
                    
                    If Err.Number <> 0 Then
                        Exit For
                    End If
                End If
            Next
            
            If Err.Number = 0 Then
                db.CommitTrans
                successMsg = "카드 사용 내역이 등록되었습니다."
                Response.Redirect "card_usage.asp"
            Else
                db.RollbackTrans
                errorMsg = "결재선 등록 중 오류가 발생했습니다: " & Err.Description
            End If
        Else
            db.RollbackTrans
            errorMsg = "카드 사용 내역 등록 중 오류가 발생했습니다: " & Err.Description
        End If
        
        On Error GoTo 0
    End If
End If

On Error GoTo 0
%>

<!--#include file="../includes/header.asp"-->

<style>
.container { max-width: 900px; }
.card {
    border: none;
    box-shadow: 0 4px 6px rgba(0,0,0,0.1);
    border-radius: 12px;
    margin-top: 2rem;
}
.card-header {
    background-color: #f8f9fa;
    border-bottom: 1px solid #eee;
    padding: 1.5rem;
    border-radius: 12px 12px 0 0 !important;
}
.card-body { padding: 2rem; }
.form-group { margin-bottom: 1.5rem; }
.form-control {
    border-radius: 6px;
    border: 1px solid #ced4da;
    padding: 0.75rem 1rem;
    font-size: 1rem;
}
.form-control:focus {
    border-color: #80bdff;
    box-shadow: 0 0 0 0.2rem rgba(0,123,255,.25);
}
label { 
    font-weight: 600;
    margin-bottom: 0.5rem;
    color: #495057;
}
.btn {
    padding: 0.75rem 1.5rem;
    font-weight: 600;
    border-radius: 6px;
    transition: all 0.2s;
}
.btn-primary {
    background-color: #007bff;
    border-color: #007bff;
}
.btn-primary:hover {
    background-color: #0069d9;
    border-color: #0062cc;
    transform: translateY(-1px);
}
.btn-secondary {
    background-color: #6c757d;
    border-color: #6c757d;
}
.btn-secondary:hover {
    background-color: #5a6268;
    border-color: #545b62;
    transform: translateY(-1px);
}
.approver-section {
    background-color: #f8f9fa;
    padding: 1.5rem;
    border-radius: 8px;
    margin-bottom: 2rem;
}
.approver-flow {
    display: flex;
    align-items: center;
    gap: 1rem;
    margin-bottom: 1rem;
}
.approver-box {
    flex: 1;
    background: white;
    border: 1px solid #dee2e6;
    border-radius: 6px;
    padding: 1rem;
    text-align: center;
    cursor: pointer;
    transition: all 0.2s;
}
.approver-box:hover {
    border-color: #007bff;
    box-shadow: 0 2px 4px rgba(0,0,0,0.1);
}
.approver-box .job-grade {
    font-size: 0.9rem;
    color: #6c757d;
    margin-bottom: 0.25rem;
}
.approver-box .name {
    font-weight: 600;
    color: #212529;
}
.approver-arrow {
    color: #adb5bd;
    font-size: 1.5rem;
    font-weight: bold;
}
.required-mark {
    color: #dc3545;
    margin-left: 2px;
}
</style>

<div class="container">
    <div class="row justify-content-center">
        <div class="col-md-10">
            <div class="card">
                <div class="card-header">
                    <h2 class="text-center mb-0">카드 사용 내역 등록</h2>
                </div>
                <div class="card-body">
                    <% If errorMsg <> "" Then %>
                    <div class="alert alert-danger" role="alert">
                        <%= errorMsg %>
                    </div>
                    <% End If %>
                    
                    <% If successMsg <> "" Then %>
                    <div class="alert alert-success" role="alert">
                        <%= successMsg %>
                    </div>
                    <% End If %>
                    
                    <form method="post" action="card_usage_add.asp">
                        <!-- 결재선 지정 섹션 -->
                        <div class="approver-section">
                            <div class="form-group">
                                <label>결재선 지정<span class="required-mark">*</span></label>
                                <div class="approver-flow">
                                    <div class="approver-box" id="approver_box1">
                                        <div class="job-grade">1차 결재자</div>
                                        <div class="name">선택해주세요</div>
                                        <input type="hidden" id="approver_step1" name="approver_step1" required>
                                    </div>
                                    <div class="approver-arrow">→</div>
                                    <div class="approver-box" id="approver_box2">
                                        <div class="job-grade">2차 결재자</div>
                                        <div class="name">선택해주세요</div>
                                        <input type="hidden" id="approver_step2" name="approver_step2">
                                    </div>
                                    <div class="approver-arrow">→</div>
                                    <div class="approver-box" id="approver_box3">
                                        <div class="job-grade">3차 결재자</div>
                                        <div class="name">선택해주세요</div>
                                        <input type="hidden" id="approver_step3" name="approver_step3">
                                    </div>
                                </div>
                                <button type="button" class="btn btn-secondary mt-3" onclick="openApprovalLinePopup()">결재자 선택</button>
                            </div>
                        </div>

                        <div class="form-group">
                            <label for="card_id">카드 선택<span class="required-mark">*</span></label>
                            <select class="form-control" id="card_id" name="card_id" required>
                                <option value="">선택해주세요</option>
                                <% 
                                If Not cardRS.EOF Then
                                    Do While Not cardRS.EOF 
                                %>
                                    <option value="<%= cardRS("card_id") %>"><%= cardRS("account_name") %></option>
                                <% 
                                    cardRS.MoveNext
                                    Loop
                                End If
                                %>
                            </select>
                        </div>
                        
                        <div class="form-group">
                            <label for="store_name">사용처<span class="required-mark">*</span></label>
                            <input type="text" class="form-control" id="store_name" name="store_name" required>
                        </div>
                        
                        <div class="form-group">
                            <label for="account_type_id">계정과목<span class="required-mark">*</span></label>
                            <select class="form-control" id="account_type_id" name="account_type_id" required>
                                <option value="">선택해주세요</option>
                                <% 
                                If Not accountTypeRS.EOF Then
                                    Do While Not accountTypeRS.EOF 
                                %>
                                    <option value="<%= accountTypeRS("account_type_id") %>"><%= accountTypeRS("type_name") %></option>
                                <% 
                                    accountTypeRS.MoveNext
                                    Loop
                                End If
                                %>
                            </select>
                        </div>
                        
                        <div class="form-group">
                            <label for="amount">금액<span class="required-mark">*</span></label>
                            <input type="number" class="form-control" id="amount" name="amount" required>
                        </div>
                        
                        <div class="form-group">
                            <label for="purpose">사용 목적<span class="required-mark">*</span></label>
                            <textarea class="form-control" id="purpose" name="purpose" rows="3" required></textarea>
                        </div>
                        
                        <div class="form-group">
                            <label for="usage_date">사용일자<span class="required-mark">*</span></label>
                            <input type="date" class="form-control" id="usage_date" name="usage_date" required>
                        </div>
                        
                        <div class="form-group text-center mt-4">
                            <button type="submit" class="btn btn-primary">등록</button>
                            <a href="card_usage.asp" class="btn btn-secondary ml-2">취소</a>
                        </div>
                    </form>
                </div>
            </div>
        </div>
    </div>
</div>

<script>
function openApprovalLinePopup() {
    var width = 800;
    var height = 600;
    var left = (screen.width - width) / 2;
    var top = (screen.height - height) / 2;
    
    window.open('approval_line_popup.asp', 'ApprovalLinePopup',
        'width=' + width + ',height=' + height + ',left=' + left + ',top=' + top + 
        ',location=no,status=no,scrollbars=yes');
}

function setApprover(step, userId, userName, job_grade) {
    var approverBox = document.getElementById('approver_box' + step);
    if (approverBox) {
        approverBox.querySelector('.job-grade').textContent = job_grade || (step + '차 결재자');
        approverBox.querySelector('.name').textContent = userName;
        document.getElementById('approver_step' + step).value = userId;
    }
}
</script>

<!--#include file="../includes/footer.asp"--> 