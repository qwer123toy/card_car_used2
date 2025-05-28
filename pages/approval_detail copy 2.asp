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
Dim docId, docType
docId = Request.QueryString("id")
docType = Request.QueryString("type")

If docId = "" Then
    Response.Write "<script>alert('잘못된 접근입니다.'); history.back();</script>"
    Response.End
End If

' 문서 타입이 지정되지 않은 경우 기본값으로 CardUsage 설정
If docType = "" Then
    docType = "CardUsage"
End If

Dim usageRS, usageSQL
Dim isCardUsage, isVehicleRequest

' 문서 타입에 따라 조회 SQL 지정
isCardUsage = (docType = "CardUsage")
isVehicleRequest = (docType = "VehicleRequests")

If isCardUsage Then
    ' 카드 사용 내역 조회
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

ElseIf isVehicleRequest Then
    ' 차량이용 신청 조회
    usageSQL = "SELECT vr.*, u.name AS user_name, u.department_id, " & _
            "d.name AS department_name, u.job_grade, j.name AS job_grade_name, " & _
            "u.user_id AS requester_id, u.name AS requester_name, " & _
            "d.department_id AS requester_dept_id, " & _
            "vr.created_at, vr.approval_status, " & _
            "fr.rate AS fuel_rate " & _
            "FROM " & dbSchema & ".VehicleRequests vr " & _
            "JOIN " & dbSchema & ".Users u ON vr.user_id = u.user_id " & _
            "LEFT JOIN " & dbSchema & ".Department d ON u.department_id = d.department_id " & _
            "LEFT JOIN " & dbSchema & ".Job_Grade j ON u.job_grade = j.job_grade_id " & _
            "LEFT JOIN " & dbSchema & ".FuelRate fr ON fr.date <= vr.start_date " & _
            "WHERE vr.request_id = ? " & _
            "ORDER BY fr.date DESC"
Else
    Response.Write "<script>alert('잘못된 문서 유형입니다.'); history.back();</script>"
    Response.End
End If

Dim cmd
Set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = db
cmd.CommandText = usageSQL
cmd.Parameters.Append cmd.CreateParameter("@doc_id", 3, 1, , docId)

Set usageRS = cmd.Execute()

If usageRS.EOF Then
    Response.Write "<script>alert('존재하지 않는 문서입니다.'); history.back();</script>"
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
             "WHERE al.target_table_name = ? AND al.target_id = ? " & _
             "ORDER BY al.approval_step"

Set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = db
cmd.CommandText = approvalSQL
cmd.Parameters.Append cmd.CreateParameter("@target_table_name", 200, 1, 30, docType)
cmd.Parameters.Append cmd.CreateParameter("@target_id", 3, 1, , docId)

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
                              "WHERE target_table_name = ? AND target_id = ?"

                Set cmd = Server.CreateObject("ADODB.Command")
                cmd.ActiveConnection = db
                cmd.CommandText = updateAllSQL
                cmd.Parameters.Append cmd.CreateParameter("@approver_id", 200, 1, 30, Session("user_id"))
                cmd.Parameters.Append cmd.CreateParameter("@comments", 200, 1, 500, comments)
                cmd.Parameters.Append cmd.CreateParameter("@approver_id2", 200, 1, 30, Session("user_id"))
                cmd.Parameters.Append cmd.CreateParameter("@target_table_name", 200, 1, 30, docType)
                cmd.Parameters.Append cmd.CreateParameter("@target_id", 3, 1, , docId)
                cmd.Execute

                ' 대상 테이블 상태 업데이트
                Dim updateDocSQL
                If isCardUsage Then
                    updateDocSQL = "UPDATE " & dbSchema & ".CardUsage SET " & _
                                  "approval_status = '반려' " & _
                                  "WHERE usage_id = ?"
                ElseIf isVehicleRequest Then
                    updateDocSQL = "UPDATE " & dbSchema & ".VehicleRequests SET " & _
                                  "approval_status = '반려' " & _
                                  "WHERE request_id = ?"
                End If
                
                Set cmd = Server.CreateObject("ADODB.Command")
                cmd.ActiveConnection = db
                cmd.CommandText = updateDocSQL
                cmd.Parameters.Append cmd.CreateParameter("@doc_id", 3, 1, , docId)
                cmd.Execute
            Else
                ' 승인 처리
                Dim updateSQL
                updateSQL = "UPDATE " & dbSchema & ".ApprovalLogs SET " & _
                           "status = ?, comments = ?, approved_at = GETDATE() " & _
                           "WHERE target_table_name = ? AND target_id = ? AND approver_id = ?"
                
                Set cmd = Server.CreateObject("ADODB.Command")
                cmd.ActiveConnection = db
                cmd.CommandText = updateSQL
                cmd.Parameters.Append cmd.CreateParameter("@status", 200, 1, 20, action)
                cmd.Parameters.Append cmd.CreateParameter("@comments", 200, 1, 500, comments)
                cmd.Parameters.Append cmd.CreateParameter("@target_table_name", 200, 1, 30, docType)
                cmd.Parameters.Append cmd.CreateParameter("@target_id", 3, 1, , docId)
                cmd.Parameters.Append cmd.CreateParameter("@approver_id", 200, 1, 30, Session("user_id"))
                cmd.Execute
                
                ' 최종 결재자이고 승인인 경우 문서 상태 업데이트
                Dim isLastApprover, totalApprovers, rs
                Set cmd = Server.CreateObject("ADODB.Command")
                cmd.ActiveConnection = db
                cmd.CommandText = "SELECT COUNT(*) AS total FROM " & dbSchema & ".ApprovalLogs WHERE target_table_name = ? AND target_id = ?"
                cmd.Parameters.Append cmd.CreateParameter("@target_table_name", 200, 1, 30, docType)
                cmd.Parameters.Append cmd.CreateParameter("@target_id", 3, 1, , docId)
                Set rs = cmd.Execute()
                totalApprovers = rs("total")
                
                isLastApprover = (myApprovalStep = totalApprovers)
                
                If isLastApprover Then
                    If isCardUsage Then
                        updateSQL = "UPDATE " & dbSchema & ".CardUsage SET " & _
                                "approval_status = '완료' " & _
                                "WHERE usage_id = ?"
                    ElseIf isVehicleRequest Then
                        updateSQL = "UPDATE " & dbSchema & ".VehicleRequests SET " & _
                                "approval_status = '완료' " & _
                                "WHERE request_id = ?"
                    End If
                    
                    Set cmd = Server.CreateObject("ADODB.Command")
                    cmd.ActiveConnection = db
                    cmd.CommandText = updateSQL
                    cmd.Parameters.Append cmd.CreateParameter("@doc_id", 3, 1, , docId)
                    cmd.Execute
                End If
            End If
            
            If Err.Number = 0 Then
                db.CommitTrans
                successMsg = "결재가 처리되었습니다."
                Response.Redirect Request.ServerVariables("URL") & "?id=" & docId & "&type=" & docType
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
}

/* 결재선 표 스타일 */
.approval-line-table-container {
    border: 2px solid #E9ECEF;
    border-radius: 12px;
    padding: 1.5rem;
    margin-bottom: 1.75rem;
    background-color: #fff;
}

.approval-line-table {
    width: 100%;
    border-collapse: collapse;
    margin-bottom: 0;
    table-layout: fixed;
}

.approval-cell {
    border: 2px solid #2C3E50;
    padding: 1rem;
    text-align: center;
    vertical-align: middle;
    background: #fff;
    position: relative;
    min-height: 80px;
    width: 20%; /* 5개 셀 동일 크기 */
    overflow: hidden;
    word-wrap: break-word;
}

.approval-cell .form-control {
    max-width: 100%;
    box-sizing: border-box;
}

.approval-cell .input-group {
    max-width: 100%;
}

.approval-cell .input-group .form-control {
    min-width: 0;
    flex: 1;
}

/* 첫 번째 행 (직급) 스타일 */
.position-row .approval-cell {
    height: 50px;
    font-weight: 600;
    color: #2C3E50;
    font-size: 1rem;
    background: #F8FAFC;
}

/* 두 번째 행 (이름과 순서) 스타일 */
.name-row .approval-cell {
    height: 120px;
    position: relative;
    padding: 1.5rem 1rem;
}

.step-number {
    position: absolute;
    top: 8px;
    left: 8px;
    background: #4A90E2;
    color: white;
    width: 20px;
    height: 20px;
    border-radius: 50%;
    display: flex;
    align-items: center;
    justify-content: center;
    font-size: 0.8rem;
    font-weight: 600;
}

.name-cell .approver-name {
    font-weight: 600;
    color: #2C3E50;
    font-size: 1rem;
    margin-top: 10px;
    line-height: 1.2;
}

.approval-status-info {
    margin-top: 10px;
}

.approval-status-info .badge {
    margin-bottom: 5px;
    display: inline-block;
}

.approval-date {
    font-size: 0.8rem;
    color: #64748B;
    margin-top: 5px;
}

.approval-comment {
    font-size: 0.85rem;
    color: #475569;
    margin-top: 8px;
    padding: 5px 8px;
    background: #F1F5F9;
    border-radius: 4px;
    border-left: 3px solid #4A90E2;
}

.badge-success {
    background: #DCFCE7 !important;
    color: #166534 !important;
    border: 1px solid #BBF7D0;
}

.badge-danger {
    background: #FEE2E2 !important;
    color: #DC2626 !important;
    border: 1px solid #FECACA;
}

.badge-secondary {
    background: #F1F5F9 !important;
    color: #475569 !important;
    border: 1px solid #E2E8F0;
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
        <h2 class="page-title">
            <% If isCardUsage Then %>
            카드 사용 내역 상세
            <% ElseIf isVehicleRequest Then %>
            차량 이용 신청 상세
            <% End If %>
        </h2>
        <div class="btn-group-nav">
            <% If isCardUsage Then %>
            <a href="card_usage.asp" class="btn btn-secondary btn-nav">
                <i class="fas fa-list me-1"></i> 목록으로
            </a>
            <% ElseIf isVehicleRequest Then %>
            <a href="vehicle_request.asp" class="btn btn-secondary btn-nav">
                <i class="fas fa-list me-1"></i> 목록으로
            </a>
            <% End If %>
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
    
        <!-- 결재 처리 섹션 -->
        <% If canApprove And (myApprovalStatus = "대기" Or myApprovalStatus = "반려") Then %>
        <div class="card mb-4">
            <div class="card-header">
                <h5 class="card-title mb-0">결재 처리</h5>
            </div>
            <div class="card-body">
                <div class="approval-line-table-container">
                    <form method="post">
                        <input type="hidden" name="doc_type" value="<%= docType %>">
                        <table class="approval-line-table">
                            <tbody>
                                <tr>
                                    <td class="approval-cell" style="background: #F8FAFC; font-weight: 600; width: 20%;">결재 의견</td>
                                    <td class="approval-cell" colspan="4" style="text-align: left; padding: 1rem;">
                                        <textarea class="form-control" id="comments" name="comments" rows="3" 
                                                placeholder="결재 의견을 입력해주세요..." style="border: 1px solid #E9ECEF; width: 100%;"></textarea>
                                    </td>
                                </tr>
                                <tr>
                                    <td class="approval-cell" style="background: #F8FAFC; font-weight: 600;">결재 처리</td>
                                    <td class="approval-cell" colspan="4" style="text-align: center; padding: 1.5rem;">
                                        <button type="submit" name="action" value="승인" class="btn btn-success me-2">
                                            <i class="fas fa-check me-2"></i> 승인
                                        </button>
                                        <button type="submit" name="action" value="반려" class="btn btn-danger me-2">
                                            <i class="fas fa-times me-2"></i> 반려
                                        </button>
                                        <a href="dashboard.asp" class="btn btn-secondary">
                                            <i class="fas fa-arrow-left me-2"></i> 취소
                                        </a>
                                    </td>
                                </tr>
                            </tbody>
                        </table>
                    </form>
                </div>
            </div>
        </div>
        <% End If %>

        <!-- 결재선 -->
        <div class="card mb-2">

            <div class="card-body" style="padding: 1rem;">
                
                    <table class="approval-line-table">
                        <tbody>
                            <!-- 첫 번째 행: 직급 (5개 고정) -->
                            <tr class="position-row">
                                <% 
                                ' 필드 존재 여부 확인 함수
                                Function FieldExists(rs, fieldName)
                                    Dim f
                                    FieldExists = False
                                    For Each f in rs.Fields
                                        If LCase(f.Name) = LCase(fieldName) Then
                                            FieldExists = True
                                            Exit Function
                                        End If
                                    Next
                                End Function
                                
                                ' 안전한 필드 접근 함수
                                Function SafeField(rs, fieldName)
                                    If FieldExists(rs, fieldName) And Not IsNull(rs(fieldName)) Then
                                        SafeField = rs(fieldName)
                                    Else
                                        SafeField = ""
                                    End If
                                End Function
                                
                                ' 결재자 정보를 배열로 저장
                                Dim approvers(5)
                                Dim approverCount
                                approverCount = 0
                                
                                If Not approvalRS.EOF Then
                                    approvalRS.MoveFirst
                                    Do While Not approvalRS.EOF And approverCount < 5
                                        Set approvers(approverCount) = CreateObject("Scripting.Dictionary")
                                        approvers(approverCount).Add "step", SafeField(approvalRS, "approval_step")
                                        approvers(approverCount).Add "name", SafeField(approvalRS, "approver_name")
                                        approvers(approverCount).Add "job_grade", SafeField(approvalRS, "job_grade_name")
                                        approvers(approverCount).Add "status", SafeField(approvalRS, "status")
                                        approvers(approverCount).Add "approved_at", SafeField(approvalRS, "approved_at")
                                        approvers(approverCount).Add "comments", SafeField(approvalRS, "comments")
                                        approverCount = approverCount + 1
                                        approvalRS.MoveNext
                                    Loop
                                    approvalRS.MoveFirst
                                End If
                                
                                ' 5개 셀 고정으로 출력
                                For i = 0 To 4
                                    Dim jobGradeName
                                    If i < approverCount And IsObject(approvers(i)) Then
                                        jobGradeName = approvers(i).Item("job_grade")
                                        If jobGradeName = "" Then jobGradeName = "직급 정보 없음"
                                    Else
                                        jobGradeName = ""
                                    End If
                                %>
                                    <td class="approval-cell">
                                        <%= jobGradeName %>
                                    </td>
                                <%
                                Next

                                %>
                            </tr>
                            <!-- 두 번째 행: 이름과 순서 (5개 고정) -->
                            <tr class="name-row">
                                <% 
                                For i = 0 To 4
                                    If i < approverCount And IsObject(approvers(i)) Then
                                        approverName = approvers(i).Item("name")
                                        approvalStatus = approvers(i).Item("status")
                                        approvedDate = approvers(i).Item("approved_at")
                                        comments = approvers(i).Item("comments")
                                
                                        Select Case approvalStatus
                                            Case "승인"
                                                stepStatusClass = "badge-success"
                                            Case "반려"
                                                stepStatusClass = "badge-danger"
                                            Case "대기"
                                                stepStatusClass = "badge-secondary"
                                            Case Else
                                                stepStatusClass = "badge-secondary"
                                        End Select
                                %>
                                    <td class="approval-cell name-cell">
                                        <span class="step-number"><%= approvers(i).Item("step") %></span>
                                        <div class="approver-name"><%= approverName %></div>
                                        <div class="approval-status-info">
                                            <span class="badge <%= stepStatusClass %>"><%= approvalStatus %></span>
                                            <% If approvedDate <> "" Then %>
                                                <div class="approval-date"><%= FormatDateTime(approvedDate, 2) %></div>
                                            <% End If %>
                                            <% If comments <> "" Then %>
                                                <div class="approval-comment">
                                                    <i class="fas fa-comment me-1"></i><%= comments %>
                                                </div>
                                            <% End If %>
                                        </div>
                                    </td>
                                <%
                                    Else
                                %>
                                    <td class="approval-cell name-cell">
                                        <!-- 빈 셀 -->
                                    </td>
                                <%
                                    End If
                                Next
                                %>
                            </tr>
                        </tbody>
                    </table>
                    
            
        

                    <% If isCardUsage Then %>
                                                <!-- 카드 사용 내역 기본정보 -->
                        <form method="post" action="card_usage_update.asp" id="updateForm" >
                            <input type="hidden" name="usage_id" value="<%= docId %>">
                            <table class="approval-line-table" >
                                <tbody >
                                    <tr>
                                        <td class="approval-cell" style="background: #F8FAFC; font-weight: 600; width: 20%; border-top: none;">제목</td>
                                        <td class="approval-cell" style="width: 80%; border-top: none;" colspan="4" >
                                            <% If usageRS("user_id") = Session("user_id") And (usageRS("approval_status") = "대기" Or usageRS("approval_status") = "반려") Then %>
                                                <input type="text" name="title" class="form-control" style="width: 96%; border-radius: 0;" value="<%= CardSafeField(usageRS, "title") %>" required>
                                            <% Else %>
                                                <%= CardSafeField(usageRS, "title") %>
                                            <% End If %>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="approval-cell" style="background: #F8FAFC; font-weight: 600; width: 20%;">신청자</td>
                                        <td class="approval-cell" style="width: 80%;" colspan="4">
                                            <% 
                                            ' 필드 존재 여부 확인 함수
                                            Function CardFieldExists(rs, fieldName)
                                                Dim f
                                                CardFieldExists = False
                                                For Each f in rs.Fields
                                                    If LCase(f.Name) = LCase(fieldName) Then
                                                        CardFieldExists = True
                                                        Exit Function
                                                    End If
                                                Next
                                            End Function
                                            
                                            ' 안전한 필드 접근 함수
                                            Function CardSafeField(rs, fieldName)
                                                If CardFieldExists(rs, fieldName) And Not IsNull(rs(fieldName)) Then
                                                    CardSafeField = rs(fieldName)
                                                Else
                                                    CardSafeField = ""
                                                End If
                                            End Function
                                            
                                            Dim userName, deptName, jobName
                                            userName = CardSafeField(usageRS, "user_name")
                                            deptName = CardSafeField(usageRS, "department_name")
                                            jobName = CardSafeField(usageRS, "job_grade_name")
                                            
                                            Response.Write userName
                                            
                                            If deptName <> "" Or jobName <> "" Then
                                                Response.Write " ("
                                                
                                                If deptName <> "" Then
                                                    Response.Write deptName
                                                    If jobName <> "" Then
                                                        Response.Write " / "
                                                    End If
                                                End If
                                                
                                                If jobName <> "" Then
                                                    Response.Write jobName
                                                End If
                                                
                                                Response.Write ")"
                                            End If
                                            %>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="approval-cell" style="background: #F8FAFC; font-weight: 600;">상태</td>
                                        <td class="approval-cell" colspan="4">
                                            <% 
                                            Dim cardStatusClass, cardStatus
                                            cardStatus = CardSafeField(usageRS, "approval_status")
                                            
                                            Select Case cardStatus
                                                Case "승인"
                                                    cardStatusClass = "bg-success"
                                                Case "반려"
                                                    cardStatusClass = "bg-danger"
                                                Case "대기"
                                                    cardStatusClass = "bg-secondary"
                                                Case "완료"
                                                    cardStatusClass = "bg-primary"
                                                Case Else
                                                    cardStatusClass = "bg-secondary"
                                            End Select
                                            %>
                                            <span class="badge <%= cardStatusClass %>">
                                                <%= cardStatus %>
                                            </span>
                                        </td>
                                    </tr>
                                    
                                    <tr>
                                        <td class="approval-cell" style="background: #F8FAFC; font-weight: 600;">카드</td>
                                        <td class="approval-cell" colspan="4">
                                            <% If usageRS("user_id") = Session("user_id") And (usageRS("approval_status") = "대기" Or usageRS("approval_status") = "반려") Then %>
                                                <select class="form-select" name="card_id" required>
                                                    <option value="">선택해주세요</option>
                                                    <% 
                                                    Dim cardSQL, cardRS
                                                    cardSQL = "SELECT card_id, account_name, issuer FROM " & dbSchema & ".CardAccount ORDER BY account_name"
                                                    Set cardRS = db.Execute(cardSQL)
                                                    
                                                    Dim selectedCardId
                                                    selectedCardId = CardSafeField(usageRS, "card_id")
                                                    
                                                    Do While Not cardRS.EOF
                                                    %>
                                                    <option value="<%= cardRS("card_id") %>" <%= IIf(CStr(cardRS("card_id")) = CStr(selectedCardId), "selected", "") %>>
                                                        <%= cardRS("account_name") %> (<%= cardRS("issuer") %>)
                                                    </option>
                                                    <%
                                                        cardRS.MoveNext
                                                    Loop
                                                    %>
                                                </select>
                                            <% Else %>
                                                <%= CardSafeField(usageRS, "account_name") %>
                                            <% End If %>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="approval-cell" style="background: #F8FAFC; font-weight: 600;">사용일자</td>
                                        <td class="approval-cell" colspan="4">
                                            <% If usageRS("user_id") = Session("user_id") And (usageRS("approval_status") = "대기" Or usageRS("approval_status") = "반려") Then %>
                                                <% 
                                                Dim usageDate
                                                usageDate = CardSafeField(usageRS, "usage_date")
                                                Dim usageDateValue
                                                
                                                If usageDate <> "" Then
                                                    usageDateValue = FormatDateTime(usageDate, 2)
                                                Else
                                                    usageDateValue = ""
                                                End If
                                                %>
                                                <input type="date" name="usage_date" class="form-control" style="width: 96%; border-radius: 0;" value="<%= usageDateValue %>" required>
                                            <% Else %>
                                                <%= FormatDateTime(CardSafeField(usageRS, "usage_date"), 2) %>
                                            <% End If %>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="approval-cell" style="background: #F8FAFC; font-weight: 600;">사용처</td>
                                        <td class="approval-cell" colspan="4">
                                            <% If usageRS("user_id") = Session("user_id") And (usageRS("approval_status") = "대기" Or usageRS("approval_status") = "반려") Then %>
                                                <input type="text" name="store_name" class="form-control"  style="width: 98%; border-radius: 0;" value="<%= CardSafeField(usageRS, "store_name") %>" required>
                                            <% Else %>
                                                <%= CardSafeField(usageRS, "store_name") %>
                                            <% End If %>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="approval-cell" style="background: #F8FAFC; font-weight: 600;">금액</td>
                                        <td class="approval-cell" colspan="4">
                                            <% If usageRS("user_id") = Session("user_id") And (usageRS("approval_status") = "대기" Or usageRS("approval_status") = "반려") Then %>
                                                <div class="input-group">
                                                    <% 
                                                    Dim amount
                                                    amount = CardSafeField(usageRS, "amount")
                                                    Dim amountValue
                                                    
                                                    If amount <> "" Then
                                                        amountValue = FormatNumber(amount)
                                                    Else
                                                        amountValue = ""
                                                    End If
                                                    %>
                                                    <input type="text" name="amount" class="form-control text-end"  style=" width: 90%; border-radius: 0;" value="<%= amountValue %>" required>
                                                    <span class="input-group-text">원</span>
                                                </div>
                                            <% Else %>
                                                <%= FormatNumber(CardSafeField(usageRS, "amount")) %>원
                                            <% End If %>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="approval-cell" style="background: #F8FAFC; font-weight: 600;">사용 목적</td>
                                        <td class="approval-cell" colspan="4">
                                            <% If usageRS("user_id") = Session("user_id") And (usageRS("approval_status") = "대기" Or usageRS("approval_status") = "반려") Then %>
                                                <textarea name="purpose" class="form-control" rows="5"  style="width: 98%; border-radius: 0;" required><%= CardSafeField(usageRS, "purpose") %></textarea>
                                            <% Else %>
                                                <%= CardSafeField(usageRS, "purpose") %>
                                            <% End If %>
                                        </td>
                                    </tr>
                                </tbody>
                            </table>
                    <% Else %>
                        <!-- 차량 이용 신청 기본정보 -->
                        <form method="post" action="vehicle_request_update.asp" id="updateForm">
                            <input type="hidden" name="request_id" value="<%= docId %>">
                            <table class="approval-line-table">
                                <tbody>
                                    <% 
                                        ' 필드 존재 여부 확인 함수
                                        Function VehicleFieldExists(rs, fieldName)
                                            Dim f
                                            VehicleFieldExists = False
                                            For Each f in rs.Fields
                                                If LCase(f.Name) = LCase(fieldName) Then
                                                    VehicleFieldExists = True
                                                    Exit Function
                                                End If
                                            Next
                                        End Function
                                        
                                        ' 안전한 필드 접근 함수
                                        Function VehicleSafeField(rs, fieldName)
                                            If VehicleFieldExists(rs, fieldName) And Not IsNull(rs(fieldName)) Then
                                                VehicleSafeField = rs(fieldName)
                                            Else
                                                VehicleSafeField = ""
                                            End If
                                        End Function
                                    
                                        ' 운행 거리와 단가로 유류비 계산
                                        Dim distance, fuelRate, tollFee, parkingFee, total_cost
                                        
                                        ' 안전하게 필드값 가져오기
                                        distance = 0
                                        If VehicleFieldExists(usageRS, "distance") And Not IsNull(usageRS("distance")) Then
                                            distance = CDbl(usageRS("distance"))
                                        End If
                                        
                                        fuelRate = 2000 ' 기본값
                                        If VehicleFieldExists(usageRS, "fuel_rate") And Not IsNull(usageRS("fuel_rate")) Then
                                            fuelRate = CDbl(usageRS("fuel_rate"))
                                        End If
                                        
                                        tollFee = 0
                                        If VehicleFieldExists(usageRS, "toll_fee") And Not IsNull(usageRS("toll_fee")) Then
                                            tollFee = CDbl(usageRS("toll_fee"))
                                        End If
                                        
                                        parkingFee = 0
                                        If VehicleFieldExists(usageRS, "parking_fee") And Not IsNull(usageRS("parking_fee")) Then
                                            parkingFee = CDbl(usageRS("parking_fee"))
                                        End If
                                        
                                        total_cost = (distance * fuelRate) + tollFee + parkingFee
                                    %>
                                    <tr>
                                        <td class="approval-cell" style="background: #F8FAFC; font-weight: 600; width: 20%; border-top: none;">제목</td>
                                        <td class="approval-cell" style="width: 80%; border-top: none;" colspan="3">
                                            <% If usageRS("user_id") = Session("user_id") And (usageRS("approval_status") = "대기" Or usageRS("approval_status") = "반려") Then %>
                                                <input type="text" name="title" class="form-control"  style="width: 96%; border-radius: 0;" value="<%= VehicleSafeField(usageRS, "title") %>" required>
                                            <% Else %>
                                                <%= VehicleSafeField(usageRS, "title") %>
                                            <% End If %>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="approval-cell" style="background: #F8FAFC; font-weight: 600; width: 20%;">신청자</td>
                                        <td class="approval-cell" style="width: 80%;" colspan="3">
                                            <%= VehicleSafeField(usageRS, "user_name") %>
                                            (<%= VehicleSafeField(usageRS, "department_name") %> / <%= VehicleSafeField(usageRS, "job_grade_name") %>)
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="approval-cell" style="background: #F8FAFC; font-weight: 600;">상태</td>
                                        <td class="approval-cell" colspan="3">
                                            <% 
                                            Dim statusClass3, approvalStatus
                                            approvalStatus = VehicleSafeField(usageRS, "approval_status")
                                            
                                            Select Case approvalStatus
                                                Case "승인"
                                                    statusClass3 = "bg-success"
                                                Case "반려"
                                                    statusClass3 = "bg-danger"
                                                Case "대기"
                                                    statusClass3 = "bg-secondary"
                                                Case "완료"
                                                    statusClass3 = "bg-primary"
                                                Case Else
                                                    statusClass3 = "bg-secondary"
                                            End Select
                                            %>
                                            <span class="badge <%= statusClass3 %>">
                                                <%= approvalStatus %>
                                            </span>
                                        </td>
                                    </tr>
                                    
                                    <tr>
                                        <td class="approval-cell" style="background: #F8FAFC; font-weight: 600;">시작일자</td>
                                        <td class="approval-cell" style="width: 20%;">
                                            <% If usageRS("user_id") = Session("user_id") And (usageRS("approval_status") = "대기" Or usageRS("approval_status") = "반려") Then %>
                                                <% 
                                                Dim startDate
                                                startDate = VehicleSafeField(usageRS, "start_date")
                                                Dim startDateValue
                                                If startDate <> "" Then
                                                    startDateValue = FormatDateTime(startDate, 2)
                                                Else
                                                    startDateValue = ""
                                                End If
                                                %>
                                                <input type="date" name="start_date" class="form-control"  style="width: 96%; border-radius: 0;" value="<%= startDateValue %>" required>
                                            <% Else %>
                                                <% 
                                                startDate = VehicleSafeField(usageRS, "start_date")
                                                If startDate <> "" Then
                                                    Response.Write FormatDateTime(startDate, 2)
                                                End If
                                                %>
                                            <% End If %>
                                        </td>
                                        <td class="approval-cell" style="background: #F8FAFC; font-weight: 600;">종료일자</td>
                                        <td class="approval-cell" style="width: 20%;">
                                            <% If usageRS("user_id") = Session("user_id") And (usageRS("approval_status") = "대기" Or usageRS("approval_status") = "반려") Then %>
                                                <% 
                                                Dim endDate
                                                endDate = VehicleSafeField(usageRS, "end_date")
                                                Dim endDateValue
                                                If endDate <> "" Then
                                                    endDateValue = FormatDateTime(endDate, 2)
                                                Else
                                                    endDateValue = ""
                                                End If
                                                %>
                                                <input type="date" name="end_date" class="form-control"  style="width: 96%; border-radius: 0;" value="<%= endDateValue %>" required>
                                            <% Else %>
                                                <% 
                                                endDate = VehicleSafeField(usageRS, "end_date")
                                                If endDate <> "" Then
                                                    Response.Write FormatDateTime(endDate, 2)
                                                End If
                                                %>
                                            <% End If %>
                                        </td>
                                        <td class="approval-cell" style="width: 20%;"></td>
                                    </tr>
                                    <tr>
                                        <td class="approval-cell" style="background: #F8FAFC; font-weight: 600;">출발지</td>
                                        <td class="approval-cell">
                                            <% If usageRS("user_id") = Session("user_id") And (usageRS("approval_status") = "대기" Or usageRS("approval_status") = "반려") Then %>
                                                <input type="text" name="start_location" class="form-control"  style="width: 96%; border-radius: 0;" value="<%= VehicleSafeField(usageRS, "start_location") %>" required>
                                            <% Else %>
                                                <%= VehicleSafeField(usageRS, "start_location") %>
                                            <% End If %>
                                        </td>
                                        <td class="approval-cell" style="background: #F8FAFC; font-weight: 600;">목적지</td>
                                        <td class="approval-cell">
                                            <% If usageRS("user_id") = Session("user_id") And (usageRS("approval_status") = "대기" Or usageRS("approval_status") = "반려") Then %>
                                                <input type="text" name="destination" class="form-control"  style="width: 96%; border-radius: 0;" value="<%= VehicleSafeField(usageRS, "destination") %>" required>
                                            <% Else %>
                                                <%= VehicleSafeField(usageRS, "destination") %>
                                            <% End If %>
                                        </td>
                                        <td class="approval-cell"></td>
                                    </tr>
                                    <tr>
                                        <td class="approval-cell" style="background: #F8FAFC; font-weight: 600;">운행거리</td>
                                        <td class="approval-cell">
                                            <% If usageRS("user_id") = Session("user_id") And (usageRS("approval_status") = "대기" Or usageRS("approval_status") = "반려") Then %>
                                                <div class="input-group">
                                                    <input type="text" name="distance" class="form-control text-end"  style="width: 80%; border-radius: 0;" value="<%= distance %>" required>
                                                    <span class="input-group-text">km</span>
                                                </div>
                                            <% Else %>
                                                <%= FormatNumber(distance) %> km
                                            <% End If %>
                                        </td>
                                        <td class="approval-cell" style="background: #F8FAFC; font-weight: 600;">유류비 단가</td>
                                        <td class="approval-cell">
                                            <%= FormatNumber(fuelRate) %> 원
                                        </td>
                                        <td class="approval-cell"></td>
                                    </tr>
                                    <tr>
                                        <td class="approval-cell" style="background: #F8FAFC; font-weight: 600;">통행료</td>
                                        <td class="approval-cell">
                                            <% If usageRS("user_id") = Session("user_id") And (usageRS("approval_status") = "대기" Or usageRS("approval_status") = "반려") Then %>
                                                <div class="input-group">
                                                    <input type="text" name="toll_fee" class="form-control text-end"  style="width: 80%; border-radius: 0;" value="<%= FormatNumber(tollFee) %>">
                                                    <span class="input-group-text">원</span>
                                                </div>
                                            <% Else %>
                                                <%= FormatNumber(tollFee) %> 원
                                            <% End If %>
                                        </td>
                                        <td class="approval-cell" style="background: #F8FAFC; font-weight: 600;">주차비</td>
                                        <td class="approval-cell">
                                            <% If usageRS("user_id") = Session("user_id") And (usageRS("approval_status") = "대기" Or usageRS("approval_status") = "반려") Then %>
                                                <div class="input-group">
                                                    <input type="text" name="parking_fee" class="form-control text-end" style="width: 80%; border-radius: 0;"  value="<%= FormatNumber(parkingFee) %>">
                                                    <span class="input-group-text">원</span>
                                                </div>
                                            <% Else %>
                                                <%= FormatNumber(parkingFee) %> 원
                                            <% End If %>
                                        </td>
                                        <td class="approval-cell"></td>
                                    </tr>
                                    <tr>
                                        <td class="approval-cell" style="background: #F8FAFC; font-weight: 600;">총 예상 비용</td>
                                        <td class="approval-cell" colspan="3">
                                            <%= FormatNumber(total_cost) %> 원
                                            <small style="color: #64748B;">(유류비: <%= FormatNumber(distance * fuelRate) %>원, 통행료: <%= FormatNumber(tollFee) %>원, 주차비: <%= FormatNumber(parkingFee) %>원)</small>
                                            <input type="hidden" name="total_cost" value="<%= total_cost %>">
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="approval-cell" style="background: #F8FAFC; font-weight: 600;">업무 목적</td>
                                        <td class="approval-cell" colspan="3">
                                            <% If usageRS("user_id") = Session("user_id") And (usageRS("approval_status") = "대기" Or usageRS("approval_status") = "반려") Then %>
                                                <textarea name="purpose" class="form-control" rows="5"  style="width: 96%; border-radius: 0;" required><%= VehicleSafeField(usageRS, "purpose") %></textarea>
                                            <% Else %>
                                                <%= VehicleSafeField(usageRS, "purpose") %>
                                            <% End If %>
                                        </td>
                                    </tr>
                                </tbody>
                            </table>
                    <% End If %>
                
                
                <!-- 수정 버튼 -->
                <% If usageRS("user_id") = Session("user_id") And (usageRS("approval_status") = "대기" Or usageRS("approval_status") = "반려") Then %>
                    <div class="text-center mt-4 pt-3" style="border-top: 1px solid #E9ECEF;">
                        <button type="submit" class="btn btn-primary me-2">
                            <i class="fas fa-save me-1"></i> 수정
                        </button>
                        <a href="dashboard.asp" class="btn btn-secondary ms-2">
                            <i class="fas fa-times me-1"></i> 취소
                        </a>
                    </div>
                </form>
                <% End If %>
            </div>
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
document.querySelector('#updateForm')?.addEventListener('submit', function(e) {
    // 카드 사용 내역의 금액 필드
    const amountInput = document.querySelector('input[name="amount"]');
    if (amountInput) {
        amountInput.value = amountInput.value.replace(/,/g, '');
    }
    
    // 차량 이용 신청의 금액 필드들
    const tollFeeInput = document.querySelector('input[name="toll_fee"]');
    if (tollFeeInput) {
        tollFeeInput.value = tollFeeInput.value.replace(/,/g, '');
    }
    
    const parkingFeeInput = document.querySelector('input[name="parking_fee"]');
    if (parkingFeeInput) {
        parkingFeeInput.value = parkingFeeInput.value.replace(/,/g, '');
    }
});

// 금액 입력 필드 이벤트 리스너
document.querySelector('input[name="amount"]')?.addEventListener('input', function(e) {
    formatAmount(this);
});

document.querySelector('input[name="toll_fee"]')?.addEventListener('input', function(e) {
    formatAmount(this);
});

document.querySelector('input[name="parking_fee"]')?.addEventListener('input', function(e) {
    formatAmount(this);
});
</script>

<!--#include file="../includes/footer.asp"--> 