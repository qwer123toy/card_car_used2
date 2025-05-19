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
          "d.name AS department_name, u.job_grade, j.name AS job_grade_name " & _
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

' 현재 사용자의 결재 권한 확인
Dim canApprove, myApprovalStep, myApprovalStatus
canApprove = False
myApprovalStep = 0
myApprovalStatus = ""

If Not approvalRS.EOF Then
    approvalRS.MoveFirst
    Do While Not approvalRS.EOF
        If approvalRS("approver_id") = Session("user_id") Then
            canApprove = True
            myApprovalStep = approvalRS("approval_step")
            myApprovalStatus = approvalRS("status")
            Exit Do
        End If
        approvalRS.MoveNext
    Loop
    approvalRS.MoveFirst
End If

' POST 요청 처리 (결재 처리)
Dim errorMsg, successMsg
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    Dim action, comments
    action = Request.Form("action")
    comments = PreventSQLInjection(Request.Form("comments"))
    
    If action <> "" And canApprove And myApprovalStatus = "대기" Then
        ' 이전 단계 결재가 모두 승인되었는지 확인
        Dim canProceed
        canProceed = True
        
        If myApprovalStep > 1 Then
            approvalRS.MoveFirst
            Do While Not approvalRS.EOF
                If approvalRS("approval_step") < myApprovalStep And approvalRS("status") <> "승인" Then
                    canProceed = False
                    Exit Do
                End If
                approvalRS.MoveNext
            Loop
        End If
        
        If canProceed Then
            ' 결재 처리
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
            
            On Error Resume Next
            cmd.Execute
            
            If Err.Number = 0 Then
                ' 최종 결재자인지 확인
                Dim isLastApprover, totalApprovers, rs
                Set cmd = Server.CreateObject("ADODB.Command")
                cmd.ActiveConnection = db
                cmd.CommandText = "SELECT COUNT(*) AS total FROM " & dbSchema & ".ApprovalLogs WHERE target_table_name = 'CardUsage' AND target_id = ?"
                cmd.Parameters.Append cmd.CreateParameter("@target_id", 3, 1, , usageId)
                Set rs = cmd.Execute()
                totalApprovers = rs("total")
                
                isLastApprover = (myApprovalStep = totalApprovers)
                
                ' 최종 결재자이고 승인인 경우 문서 상태 업데이트
                If isLastApprover And action = "승인" Then
                    updateSQL = "UPDATE " & dbSchema & ".CardUsage SET " & _
                              "status = '완료' " & _
                              "WHERE usage_id = ?"
                    
                    Set cmd = Server.CreateObject("ADODB.Command")
                    cmd.ActiveConnection = db
                    cmd.CommandText = updateSQL
                    cmd.Parameters.Append cmd.CreateParameter("@usage_id", 3, 1, , usageId)
                    cmd.Execute
                End If
                
                successMsg = "결재가 처리되었습니다."
                
                ' 페이지 새로고침
                Response.Redirect Request.ServerVariables("URL") & "?id=" & usageId
            Else
                errorMsg = "결재 처리 중 오류가 발생했습니다: " & Err.Description
            End If
            On Error GoTo 0
        Else
            errorMsg = "이전 단계의 결재가 완료되지 않았습니다."
        End If
    End If
End If
%>

<!--#include file="../includes/header.asp"-->

<div class="container mt-5">
    <div class="row justify-content-center">
        <div class="col-md-10">
            <div class="card">
                <div class="card-header">
                    <h2 class="text-center">카드 사용 내역 결재</h2>
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
                    
                    <!-- 카드 사용 내역 정보 -->
                    <div class="card mb-4">
                        <div class="card-header">
                            <h5 class="mb-0">카드 사용 내역</h5>
                        </div>
                        <div class="card-body">
                            <table class="table">
                                <tr>
                                    <th style="width: 150px;">신청자</th>
                                    <td>
                                        <%= usageRS("user_name") %>
                                        (<%= usageRS("department_name") %> / <%= usageRS("job_grade_name") %>)
                                    </td>
                                </tr>
                                <tr>
                                    <th>카드</th>
                                    <td><%= usageRS("account_name") %></td>
                                </tr>
                                <tr>
                                    <th>사용일자</th>
                                    <td><%= FormatDate(usageRS("usage_date")) %></td>
                                </tr>
                                <tr>
                                    <th>금액</th>
                                    <td><%= FormatNumber(usageRS("amount")) %>원</td>
                                </tr>
                                <tr>
                                    <th>사용처</th>
                                    <td><%= usageRS("store_name") %></td>
                                </tr>
                                <tr>
                                    <th>사용 목적</th>
                                    <td><%= usageRS("purpose") %></td>
                                </tr>
                            </table>
                        </div>
                    </div>
                    
                    <!-- 결재선 정보 -->
                    <div class="card mb-4">
                        <div class="card-header">
                            <h5 class="mb-0">결재 이력</h5>
                        </div>
                        <div class="card-body">
                            <table class="table">
                                <thead>
                                    <tr>
                                        <th>순서</th>
                                        <th>결재자</th>
                                        <th>부서/직급</th>
                                        <th>상태</th>
                                        <th>결재일시</th>
                                        <th>의견</th>
                                    </tr>
                                </thead>
                                <tbody>
                                    <% 
                                    If Not approvalRS.EOF Then
                                        Do While Not approvalRS.EOF 
                                    %>
                                    <tr>
                                        <td><%= approvalRS("approval_step") %>차</td>
                                        <td><%= approvalRS("approver_name") %></td>
                                        <td><%= approvalRS("department_name") %> / <%= approvalRS("job_grade_name") %></td>
                                        <td>
                                            <% 
                                            Dim statusClass
                                            Select Case approvalRS("status")
                                                Case "승인"
                                                    statusClass = "badge bg-success"
                                                Case "반려"
                                                    statusClass = "badge bg-danger"
                                                Case "대기"
                                                    statusClass = "badge bg-secondary"
                                            End Select
                                            %>
                                            <span class="<%= statusClass %>"><%= approvalRS("status") %></span>
                                        </td>
                                        <td>
                                            <% If Not IsNull(approvalRS("approved_at")) Then %>
                                                <%= FormatDateTime(approvalRS("approved_at"), 2) %>
                                            <% Else %>
                                                -
                                            <% End If %>
                                        </td>
                                        <td>
                                            <% If Not IsNull(approvalRS("comments")) Then %>
                                                <%= approvalRS("comments") %>
                                            <% Else %>
                                                -
                                            <% End If %>
                                        </td>
                                    </tr>
                                    <% 
                                        approvalRS.MoveNext
                                        Loop
                                    End If 
                                    %>
                                </tbody>
                            </table>
                        </div>
                    </div>
                    
                    <!-- 결재 처리 폼 -->
                    <% If canApprove And myApprovalStatus = "대기" Then %>
                    <div class="card">
                        <div class="card-header">
                            <h5 class="mb-0">결재 처리</h5>
                        </div>
                        <div class="card-body">
                            <form method="post">
                                <div class="form-group">
                                    <label for="comments">의견</label>
                                    <textarea class="form-control" id="comments" name="comments" rows="3"></textarea>
                                </div>
                                
                                <div class="text-center mt-4">
                                    <button type="submit" name="action" value="승인" class="btn btn-success">승인</button>
                                    <button type="submit" name="action" value="반려" class="btn btn-danger">반려</button>
                                    <a href="dashboard.asp" class="btn btn-secondary">취소</a>
                                </div>
                            </form>
                        </div>
                    </div>
                    <% End If %>
                </div>
            </div>
        </div>
    </div>
</div>

<!--#include file="../includes/footer.asp"--> 