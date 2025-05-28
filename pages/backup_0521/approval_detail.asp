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
    RedirectTo("/contents/card_car_used/index.asp")
End If

' 파라미터 검증
Dim usageId
usageId = Request.QueryString("id")
If usageId = "" Then
    Response.Write "<script>alert('잘못된 접근입니다.'); history.back();</script>"
    Response.End
End If

' 카드 사용 내역 조회
Dim usageRS, usageSQL
usageSQL = "SELECT cu.*, ca.account_name, u.name AS user_name, u.department_id, " & _
          "d.name AS department_name, u.job_grade, j.name AS job_grade_name, " & _
          "u.user_id AS requester_id, u.name AS requester_name, " & _
          "d.department_id AS requester_dept_id, " & _
          "cu.created_at, cu.approval_status " & _
          "FROM " & dbSchema & ".CardUsage cu " & _
          "JOIN " & dbSchema & ".CardAccount ca ON cu.card_id = ca.card_id " & _
          "JOIN " & dbSchema & ".Users u ON cu.user_id = u.user_id " & _
          "LEFT JOIN " & dbSchema & ".Department d ON u.department_id = d.department_id " & _
          "LEFT JOIN " & dbSchema & ".Job_Grade j ON u.job_grade = j.job_grade_id " & _
          "WHERE cu.usage_id = ? "

Dim cmd
Set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = db
cmd.CommandText = usageSQL
cmd.Parameters.Append cmd.CreateParameter("@usage_id", 3, 1, , usageId)

Set usageRS = cmd.Execute()

If usageRS.EOF Then
    Response.Write "<script>alert('존재하지 않는 카드 사용 내역입니다.'); history.back();</script>"
    Response.End
End If

' 결재 정보 조회
Dim approvalRS, approvalSQL
approvalSQL = "SELECT al.*, u.name AS approver_name, u.department_id, " & _
             "d.name AS department_name, u.job_grade, j.name AS job_grade_name " & _
             "FROM " & dbSchema & ".ApprovalLogs al " & _
             "JOIN " & dbSchema & ".Users u ON al.approver_id = u.user_id " & _
             "LEFT JOIN " & dbSchema & ".Department d ON u.department_id = d.department_id " & _
             "LEFT JOIN " & dbSchema & ".Job_Grade j ON u.job_grade = j.job_grade_id " & _
             "WHERE al.target_table_name = 'CardUsage' AND al.target_id = ? " & _
             "ORDER BY al.approval_step"

Set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = db
cmd.CommandText = approvalSQL
cmd.Parameters.Append cmd.CreateParameter("@target_id", 3, 1, , usageId)

Set approvalRS = cmd.Execute()

' 현재 사용자의 결재 권한 확인 및 이전 단계 결재 상태 확인
Dim canApprove, myApprovalStep, myApprovalStatus, prevApprovalComplete
canApprove = False
myApprovalStep = 0
myApprovalStatus = ""
prevApprovalComplete = True

If Not approvalRS.EOF Then
    approvalRS.MoveFirst
    
    ' 먼저 현재 사용자의 결재 단계 찾기
    Do While Not approvalRS.EOF
        If approvalRS("approver_id") = Session("user_id") Then
            canApprove = True
            myApprovalStep = approvalRS("approval_step")
            myApprovalStatus = approvalRS("status")
            Exit Do
        End If
        approvalRS.MoveNext
    Loop
    
    ' 이전 단계들의 결재 상태 확인
    If myApprovalStep > 1 Then
        approvalRS.MoveFirst
        Do While Not approvalRS.EOF
            If approvalRS("approval_step") < myApprovalStep Then
                If approvalRS("status") <> "승인" Then
                    prevApprovalComplete = False
                    Exit Do
                End If
            End If
            approvalRS.MoveNext
        Loop
    End If
    approvalRS.MoveFirst
End If

' POST 요청 처리 (결재 처리)
Dim errorMsg, successMsg
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    Dim action, comments
    action = Request.Form("action")
    comments = PreventSQLInjection(Request.Form("comments"))
    
    If action <> "" And canApprove And (myApprovalStatus = "대기" or myApprovalStatus = "반려") Then
        If Not prevApprovalComplete Then
            errorMsg = "이전 단계의 결재가 완료되지 않았습니다."
        Else
            ' 결재 처리
            db.BeginTrans
            
            On Error Resume Next
            
            If action = "반려" Then
                ' 반려 처리 시 모든 결재자의 상태 업데이트
                Dim updateAllSQL
                updateAllSQL = "UPDATE " & dbSchema & ".ApprovalLogs SET " & _
                              "status = CASE " & _
                              "   WHEN approval_step = 1 THEN '반려' " & _
                              "   ELSE '대기' " & _
                              "END, " & _
                              "comments = CASE " & _
                              "   WHEN approver_id = ? THEN ? " & _
                              "   ELSE NULL " & _
                              "END, " & _
                              "approved_at = CASE " & _
                              "   WHEN approver_id = ? THEN GETDATE() " & _
                              "   ELSE NULL " & _
                              "END " & _
                              "WHERE target_table_name = 'CardUsage' AND target_id = ?"

                Set cmd = Server.CreateObject("ADODB.Command")
                cmd.ActiveConnection = db
                cmd.CommandText = updateAllSQL
                cmd.Parameters.Append cmd.CreateParameter("@approver_id", 200, 1, 30, Session("user_id"))
                cmd.Parameters.Append cmd.CreateParameter("@comments", 200, 1, 500, comments)
                cmd.Parameters.Append cmd.CreateParameter("@approver_id2", 200, 1, 30, Session("user_id"))
                cmd.Parameters.Append cmd.CreateParameter("@target_id", 3, 1, , usageId)
                cmd.Execute

                ' CardUsage 상태 업데이트
                Dim updateUsageSQL
                updateUsageSQL = "UPDATE " & dbSchema & ".CardUsage SET " & _
                                "approval_status = '반려' " & _
                                "WHERE usage_id = ?"
                
                Set cmd = Server.CreateObject("ADODB.Command")
                cmd.ActiveConnection = db
                cmd.CommandText = updateUsageSQL
                cmd.Parameters.Append cmd.CreateParameter("@usage_id", 3, 1, , usageId)
                cmd.Execute
            Else
                ' 승인 처리
                Dim updateSQL
                updateSQL = "UPDATE " & dbSchema & ".ApprovalLogs SET " & _
                           "status = ?, comments = ?, approved_at = GETDATE() " & _
                           "WHERE target_table_name = 'CardUsage' AND target_id = ? AND approver_id = ?"
                
                Set cmd = Server.CreateObject("ADODB.Command")
                cmd.ActiveConnection = db
                cmd.CommandText = updateSQL
                cmd.Parameters.Append cmd.CreateParameter("@status", 200, 1, 20, action)
                cmd.Parameters.Append cmd.CreateParameter("@comments", 200, 1, 500, comments)
                cmd.Parameters.Append cmd.CreateParameter("@target_id", 3, 1, , usageId)
                cmd.Parameters.Append cmd.CreateParameter("@approver_id", 200, 1, 30, Session("user_id"))
                cmd.Execute
                
                ' 최종 결재자이고 승인인 경우 문서 상태 업데이트
                Dim isLastApprover, totalApprovers, rs
                Set cmd = Server.CreateObject("ADODB.Command")
                cmd.ActiveConnection = db
                cmd.CommandText = "SELECT COUNT(*) AS total FROM " & dbSchema & ".ApprovalLogs WHERE target_table_name = 'CardUsage' AND target_id = ?"
                cmd.Parameters.Append cmd.CreateParameter("@target_id", 3, 1, , usageId)
                Set rs = cmd.Execute()
                totalApprovers = rs("total")
                
                isLastApprover = (myApprovalStep = totalApprovers)
                
                If isLastApprover Then
                    updateSQL = "UPDATE " & dbSchema & ".CardUsage SET " & _
                              "approval_status = '완료' " & _
                              "WHERE usage_id = ?"
                    
                    Set cmd = Server.CreateObject("ADODB.Command")
                    cmd.ActiveConnection = db
                    cmd.CommandText = updateSQL
                    cmd.Parameters.Append cmd.CreateParameter("@usage_id", 3, 1, , usageId)
                    cmd.Execute
                End If
            End If
            
            If Err.Number = 0 Then
                db.CommitTrans
                successMsg = "결재가 처리되었습니다."
                Response.Redirect Request.ServerVariables("URL") & "?id=" & usageId
            Else
                db.RollbackTrans
                errorMsg = "결재 처리 중 오류가 발생했습니다: " & Err.Description
            End If
            On Error GoTo 0
        End If
    Else
        errorMsg = "이전 단계의 결재가 완료되지 않았습니다."
    End If
End If
%>

<!--#include file="../includes/header.asp"-->

<style>
.container {
    max-width: 1200px;
    margin: 0 auto;
    padding: 2rem 1rem;
}

.card {
    border: none;
    box-shadow: 0 0 20px rgba(0,0,0,0.05);
    border-radius: 16px;
    margin-bottom: 2rem;
    background: #fff;
    overflow: hidden;
}

.card-header {
    background: linear-gradient(to right, #4A90E2, #5A9EEA);
    border-bottom: none;
    padding: 1.5rem;
}

.card-header h5 {
    color: #fff;
    font-weight: 600;
    margin: 0;
    font-size: 1.25rem;
}

.card-body {
    padding: 2rem;
}

.table {
    margin-bottom: 0;
}

.table th {
    background-color: #F8FAFC !important;
    color: #2C3E50;
    font-weight: 600;
    border-bottom: 2px solid #E9ECEF;
    padding: 1rem;
    font-size: 0.95rem;
}

.table td {
    padding: 1rem;
    vertical-align: middle;
    border-bottom: 1px solid #E9ECEF;
    color: #2C3E50;
}

.form-control {
    border-radius: 8px;
    border: 2px solid #E9ECEF;
    padding: 0.875rem 1rem;
    font-size: 1rem;
    transition: all 0.2s ease;
}

.form-control:focus {
    border-color: #4A90E2;
    box-shadow: 0 0 0 4px rgba(74,144,226,0.1);
}

.form-select {
    border-radius: 8px;
    border: 2px solid #E9ECEF;
    padding: 0.875rem 1rem;
    font-size: 1rem;
}

.input-group-text {
    background-color: #F8FAFC;
    border: 2px solid #E9ECEF;
    border-left: none;
    color: #2C3E50;
    font-weight: 500;
}

.required-mark {
    color: #E74C3C;
    margin-left: 4px;
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

.btn-success {
    background: linear-gradient(to right, #2ECC71, #27AE60);
    border: none;
    color: white;
}

.btn-success:hover {
    transform: translateY(-2px);
    box-shadow: 0 4px 12px rgba(46,204,113,0.2);
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

.bg-secondary {
    background: #F1F5F9 !important;
    color: #475569;
}

.approval-line {
    background: #F8FAFC;
    border-radius: 12px;
    padding: 1.5rem;
    margin-bottom: 2rem;
}

.approval-steps {
    display: flex;
    gap: 1rem;
    margin-bottom: 1rem;
}

.approval-step {
    flex: 1;
    background: white;
    border: 2px solid #E9ECEF;
    border-radius: 10px;
    padding: 1.25rem;
    transition: all 0.2s ease;
}

.approval-step:hover {
    border-color: #4A90E2;
    box-shadow: 0 4px 12px rgba(74,144,226,0.1);
}

.step-label {
    display: inline-block;
    font-weight: 600;
    color: #2C3E50;
    margin-bottom: 1rem;
    font-size: 0.95rem;
    padding: 0.5rem 1rem;
    background: #E9ECEF;
    border-radius: 6px;
}

.approver-info {
    margin-top: 0.5rem;
}

.approver-name {
    font-weight: 600;
    color: #2C3E50;
    margin-bottom: 0.25rem;
}

.approver-dept {
    font-size: 0.9rem;
    color: #64748B;
}

.approval-status {
    display: flex;
    align-items: center;
    justify-content: space-between;
    margin-top: 1rem;
    padding-top: 1rem;
    border-top: 1px solid #E9ECEF;
}

.badge {
    padding: 0.5rem 1rem;
    font-weight: 500;
    border-radius: 6px;
    font-size: 0.875rem;
    margin: 0 0.5rem;
}

.approval-date {
    font-size: 0.875rem;
    color: #64748B;
}

.approval-comment {
    margin-top: 0.75rem;
    padding: 0.75rem;
    background: #F1F5F9;
    border-radius: 6px;
    font-size: 0.875rem;
    color: #475569;
}

.approval-comment i {
    margin-right: 0.5rem;
    color: #64748B;
}

.comments-section {
    margin-top: 2rem;
    padding: 1.5rem;
    background: #F8FAFC;
    border-radius: 12px;
}

.comments-section textarea {
    border: 2px solid #E9ECEF;
    border-radius: 8px;
    padding: 1rem;
    width: 100%;
    min-height: 100px;
    margin-bottom: 1rem;
}

.text-danger {
    color: #E74C3C !important;
}

.alert {
    border: none;
    border-radius: 12px;
    padding: 1.25rem 1.5rem;
    margin-bottom: 2rem;
    font-weight: 500;
}

.alert-danger {
    background: #FDF1F1;
    color: #E74C3C;
}

.alert-success {
    background: #EDF9F0;
    color: #2ECC71;
}

.page-header {
    display: flex;
    justify-content: space-between;
    align-items: center;
    margin-bottom: 2rem;
    padding: 1rem;
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

.btn-group-nav {
    display: flex;
    gap: 0.5rem;
}

.btn-nav {
    padding: 0.625rem 1.25rem;
    font-size: 0.9rem;
}
</style>

<div class="container">
    <div class="page-header">
        <h2 class="page-title">카드 사용 내역 상세</h2>
        <div class="btn-group-nav">
            <a href="card_usage.asp" class="btn btn-secondary btn-nav">
                <i class="fas fa-list me-1"></i> 목록으로
            </a>
            <a href="dashboard.asp" class="btn btn-secondary btn-nav">
                <i class="fas fa-home me-1"></i> 대시보드
            </a>
        </div>
    </div>

    <% If errorMsg <> "" Then %>
        <div class="alert alert-danger" role="alert">
            <i class="fas fa-exclamation-circle me-2"></i><%= errorMsg %>
        </div>
    <% End If %>
    
    <% If successMsg <> "" Then %>
        <div class="alert alert-success" role="alert">
            <i class="fas fa-check-circle me-2"></i><%= successMsg %>
        </div>
    <% End If %>

    <div class="card">
        <div class="card-header">
            <h5 class="card-title mb-0">기본 정보</h5>
        </div>
        <div class="card-body">
            <% If usageRS("user_id") = Session("user_id") And (usageRS("approval_status") = "대기" Or usageRS("approval_status") = "반려") Then %>
                <form method="post" action="card_usage_update.asp">
                    <input type="hidden" name="usage_id" value="<%= usageId %>">
                    <table class="table table-bordered">
                        <tr>
                            <th style="width: 15%;">신청자</th>
                            <td style="width: 35%;">
                                <%= usageRS("user_name") %>
                                (<%= usageRS("department_name") %> / <%= usageRS("job_grade_name") %>)
                            </td>
                            <th style="width: 15%;">상태</th>
                            <td style="width: 35%;">
                                <% 
                                Dim statusClass
                                Select Case usageRS("approval_status")
                                    Case "승인"
                                        statusClass = "bg-success"
                                    Case "반려"
                                        statusClass = "bg-danger"
                                    Case "대기"
                                        statusClass = "bg-secondary"
                                    Case "완료"
                                        statusClass = "bg-primary"
                                End Select
                                %>
                                <span class="badge <%= statusClass %>">
                                    <%= usageRS("approval_status") %>
                                </span>
                            </td>
                        </tr>
                        <tr>
                            <th>카드 선택<span class="required-mark">*</span></th>
                            <td>
                                <select class="form-select" name="card_id" required>
                                    <option value="">선택해주세요</option>
                                    <% 
                                    Dim cardSQL, cardRS
                                    cardSQL = "SELECT card_id, account_name FROM " & dbSchema & ".CardAccount ORDER BY account_name"
                                    Set cardRS = db.Execute(cardSQL)
                                    Do While Not cardRS.EOF
                                    %>
                                    <option value="<%= cardRS("card_id") %>" <%= IIf(CStr(cardRS("card_id")) = CStr(usageRS("card_id")), "selected", "") %>>
                                        <%= cardRS("account_name") %>
                                    </option>
                                    <%
                                        cardRS.MoveNext
                                    Loop
                                    %>
                                </select>
                            </td>
                            <th>사용일자<span class="required-mark">*</span></th>
                            <td>
                                <input type="date" name="usage_date" class="form-control" 
                                       value="<%= FormatDateTime(usageRS("usage_date"), 2) %>" required>
                            </td>
                        </tr>
                        <tr>
                            <th>사용처<span class="required-mark">*</span></th>
                            <td>
                                <input type="text" name="store_name" class="form-control" 
                                       value="<%= usageRS("store_name") %>" required>
                            </td>
                            <th>금액<span class="required-mark">*</span></th>
                            <td>
                                <div class="input-group">
                                    <input type="text" name="amount" class="form-control text-end" 
                                           value="<%= FormatNumber(usageRS("amount")) %>" required>
                                    <span class="input-group-text">원</span>
                                </div>
                            </td>
                        </tr>
                        <tr>
                            <th>사용 목적<span class="required-mark">*</span></th>
                            <td colspan="3">
                                <textarea name="purpose" class="form-control" rows="3" required><%= usageRS("purpose") %></textarea>
                            </td>
                        </tr>
                    </table>
                    <div class="text-center mt-4">
                        <button type="submit" class="btn btn-primary me-2">
                            <i class="fas fa-save me-1"></i> 수정사항 저장
                        </button>
                        <a href="dashboard.asp" class="btn btn-secondary ms-2">
                            <i class="fas fa-times me-1"></i> 취소
                        </a>
                    </div>
                </form>
            <% Else %>
                <table class="table table-bordered">
                    <tr>
                        <th style="width: 15%;">신청자</th>
                        <td style="width: 35%;">
                            <%= usageRS("user_name") %>
                            (<%= usageRS("department_name") %> / <%= usageRS("job_grade_name") %>)
                        </td>
                        <th style="width: 15%;">상태</th>
                        <td style="width: 35%;">
                            <% 
                            Select Case usageRS("approval_status")
                                Case "승인"
                                    statusClass = "bg-success"
                                Case "반려"
                                    statusClass = "bg-danger"
                                Case "대기"
                                    statusClass = "bg-secondary"
                                Case "완료"
                                    statusClass = "bg-primary"
                            End Select
                            %>
                            <span class="badge <%= statusClass %>">
                                <%= usageRS("approval_status") %>
                            </span>
                        </td>
                    </tr>
                    <tr>
                        <th>카드</th>
                        <td><%= usageRS("account_name") %></td>
                        <th>사용일자</th>
                        <td><%= FormatDateTime(usageRS("usage_date"), 2) %></td>
                    </tr>
                    <tr>
                        <th>사용처</th>
                        <td><%= usageRS("store_name") %></td>
                        <th>금액</th>
                        <td><%= FormatNumber(usageRS("amount")) %>원</td>
                    </tr>
                    <tr>
                        <th>사용 목적</th>
                        <td colspan="3"><%= usageRS("purpose") %></td>
                    </tr>
                </table>
            <% End If %>
        </div>
    </div>

    <!-- 결재선 정보 -->
    <div class="card mb-4">
        <div class="card-header">
            <h5 class="card-title mb-0">결재선</h5>
        </div>
        <div class="card-body">
            <div class="approval-line">
                <div class="approval-steps">
                    <% 
                    If Not approvalRS.EOF Then
                        approvalRS.MoveFirst
                        Do While Not approvalRS.EOF 
                    %>
                        <div class="approval-step">
                            <div class="step-label"><%= approvalRS("approval_step") %>차 결재</div>
                            <div class="approver-info">
                                <div class="approver-name"><%= approvalRS("approver_name") %></div>
                                <div class="approver-dept"><%= approvalRS("department_name") %> / <%= approvalRS("job_grade_name") %></div>
                                <div class="approval-status">
                                    <% 
                                    Select Case approvalRS("status")
                                        Case "승인"
                                            statusClass = "bg-success"
                                        Case "반려"
                                            statusClass = "bg-danger"
                                        Case "대기"
                                            statusClass = "bg-secondary"
                                    End Select
                                    %>
                                    <span class="badge <%= statusClass %>"><%= approvalRS("status") %></span>
                                    <% If Not IsNull(approvalRS("approved_at")) Then %>
                                        <span class="approval-date"><%= FormatDateTime(approvalRS("approved_at"), 2) %></span>
                                    <% End If %>
                                </div>
                                <% If approvalRS("comments") <> "" Then %>
                                    <div class="approval-comment">
                                        <i class="fas fa-comment"></i><%= approvalRS("comments") %>
                                    </div>
                                <% End If %>
                            </div>
                        </div>
                    <%
                            approvalRS.MoveNext
                        Loop
                        approvalRS.MoveFirst
                    End If
                    %>
                </div>

                <% If canApprove And myApprovalStatus = "대기" Then %>
                    <div class="comments-section">
                        <form method="post">
                            <div class="form-group">
                                <label for="comments" class="form-label">결재 의견</label>
                                <textarea class="form-control" id="comments" name="comments" rows="3" 
                                        placeholder="결재 의견을 입력해주세요..."></textarea>
                            </div>
                            
                            <div class="text-center mt-4">
                                <button type="submit" name="action" value="승인" class="btn btn-success">
                                    <i class="fas fa-check me-2"></i> 승인
                                </button>
                                <button type="submit" name="action" value="반려" class="btn btn-danger">
                                    <i class="fas fa-times me-2"></i> 반려
                                </button>
                                <a href="dashboard.asp" class="btn btn-secondary">
                                    <i class="fas fa-arrow-left me-2"></i> 취소
                                </a>
                            </div>
                        </form>
                    </div>
                <% End If %>
            </div>
        </div>
    </div>
</div>

<script>
// 금액 입력 필드 포맷팅
function formatAmount(input) {
    let value = input.value.replace(/[^\d]/g, '');
    if (value) {
        input.value = new Intl.NumberFormat('ko-KR').format(value);
    }
}

// 폼 제출 시 금액 콤마 제거
document.querySelector('form')?.addEventListener('submit', function(e) {
    const amountInput = document.querySelector('input[name="amount"]');
    if (amountInput) {
        amountInput.value = amountInput.value.replace(/,/g, '');
    }
});

// 금액 입력 필드 이벤트 리스너
document.querySelector('input[name="amount"]')?.addEventListener('input', function(e) {
    formatAmount(this);
});
</script>

<!--#include file="../includes/footer.asp"--> 