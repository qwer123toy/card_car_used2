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

' 사용자 정보 조회
Dim SQL, rs, userName, userDeptId, userDeptName
On Error Resume Next

' 세션에서 사용자 이름 불러오기 (DB 오류 시 대비)
If Session("name") <> "" Then
    userName = Session("name")
Else
    userName = Session("user_id")
End If

' 세션에서 부서 ID 불러오기
If Session("department_id") <> "" Then
    userDeptId = Session("department_id")
Else
    userDeptId = 0
End If

' 부서명 기본값 설정
userDeptName = "부서 정보 없음"

' DB에서 사용자 정보 조회 시도
SQL = "SELECT u.name, u.department_id, d.name AS department_name FROM Users u " & _
      "LEFT JOIN Department d ON u.department_id = d.department_id " & _
      "WHERE u.user_id = '" & Session("user_id") & "'"
      
' 파라미터화된 쿼리 사용을 위해 명령 객체 생성
Dim cmd
Set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = db
cmd.CommandText = "SELECT u.name, u.department_id, d.name AS department_name FROM Users u " & _
                 "LEFT JOIN Department d ON u.department_id = d.department_id " & _
                 "WHERE u.user_id = ?"
cmd.Parameters.Append cmd.CreateParameter("@user_id", 200, 1, 30, Session("user_id"))

' 명령 실행
Set rs = cmd.Execute()

' DB 조회에 성공했고 사용자 정보가 있는 경우
If Err.Number = 0 And Not rs.EOF Then
    userName = rs("name")
    userDeptId = rs("department_id")
    
    ' NULL 값 처리
    If Not IsNull(rs("department_name")) Then
    userDeptName = rs("department_name")
End If
    
rs.Close
End If

' 최근 카드 사용 내역 조회 시도
Dim recentCardRS
Set recentCardRS = Server.CreateObject("ADODB.Recordset")

' 빈 레코드셋 초기화 (오류 발생 시 이 레코드셋 사용)
recentCardRS.Fields.Append "usage_id", 3 ' adInteger
recentCardRS.Fields.Append "usage_date", 7 ' adDate
recentCardRS.Fields.Append "amount", 6 ' adCurrency
recentCardRS.Fields.Append "account_name", 200, 100 ' adVarChar
recentCardRS.Open

' 실제 데이터 조회 시도
Dim recentCardCmd
Set recentCardCmd = Server.CreateObject("ADODB.Command")
recentCardCmd.ActiveConnection = db
recentCardCmd.CommandText = "SELECT TOP 5 c.usage_id, c.usage_date, c.amount, ca.account_name " & _
                "FROM CardUsage c " & _
                "JOIN CardAccount ca ON c.card_id = ca.card_id " & _
                          "WHERE c.user_id = ? " & _
                "ORDER BY c.usage_date DESC"
recentCardCmd.Parameters.Append recentCardCmd.CreateParameter("@user_id", 200, 1, 30, Session("user_id"))

On Error Resume Next
Dim tempRS
Set tempRS = recentCardCmd.Execute()

' 데이터 조회 성공 시 임시 레코드셋에서 실제 레코드셋으로 데이터 복사
If Err.Number = 0 And Not tempRS.EOF Then
    ' 기존 빈 레코드셋 닫고 정상 레코드셋으로 대체
    recentCardRS.Close
    Set recentCardRS = tempRS
End If
On Error GoTo 0

' 최근 차량 사용 내역 조회 시도
Dim recentVehicleRS
Set recentVehicleRS = Server.CreateObject("ADODB.Recordset")

' 빈 레코드셋 초기화 (오류 발생 시 이 레코드셋 사용)
recentVehicleRS.Fields.Append "request_id", 3 ' adInteger
recentVehicleRS.Fields.Append "request_date", 7 ' adDate
recentVehicleRS.Fields.Append "purpose", 200, 100 ' adVarChar
recentVehicleRS.Fields.Append "approval_status", 200, 20 ' adVarChar
recentVehicleRS.Fields.Append "distance", 5 ' adDouble
recentVehicleRS.Fields.Append "fuel_rate", 5 ' adDouble
recentVehicleRS.Open

' 실제 데이터 조회 시도
Dim recentVehicleCmd
Set recentVehicleCmd = Server.CreateObject("ADODB.Command")
recentVehicleCmd.ActiveConnection = db
recentVehicleCmd.CommandText = "SELECT TOP 5 v.request_id, v.request_date, v.purpose, v.approval_status, " & _
                   "v.distance, ISNULL(fr.rate, 2000) as fuel_rate " & _
                   "FROM VehicleRequests v " & _
                   "LEFT JOIN (SELECT TOP 1 * FROM FuelRate ORDER BY date DESC) fr ON 1=1 " & _
                   "WHERE v.user_id = ? AND v.is_deleted = 0 " & _
                   "ORDER BY v.request_date DESC"
recentVehicleCmd.Parameters.Append recentVehicleCmd.CreateParameter("@user_id", 200, 1, 30, Session("user_id"))

On Error Resume Next
Dim tempVehicleRS
Set tempVehicleRS = recentVehicleCmd.Execute()

' 데이터 조회 성공 시 임시 레코드셋에서 실제 레코드셋으로 데이터 복사
If Err.Number = 0 And Not tempVehicleRS.EOF Then
    ' 기존 빈 레코드셋 닫고 정상 레코드셋으로 대체
    recentVehicleRS.Close
    Set recentVehicleRS = tempVehicleRS
End If

' 결재 대기 문서 조회
Dim approvalPendingRS
Set approvalPendingRS = Server.CreateObject("ADODB.Recordset")

' 빈 레코드셋 초기화 (오류 발생 시 이 레코드셋 사용)
approvalPendingRS.Fields.Append "target_table_name", 200, 30 ' adVarChar
approvalPendingRS.Fields.Append "doc_id", 3 ' adInteger
approvalPendingRS.Fields.Append "doc_date", 7 ' adDate
approvalPendingRS.Fields.Append "title", 200, 100 ' adVarChar
approvalPendingRS.Fields.Append "amount", 6 ' adCurrency
approvalPendingRS.Fields.Append "requester_name", 200, 50 ' adVarChar
approvalPendingRS.Fields.Append "department_name", 200, 50 ' adVarChar
approvalPendingRS.Fields.Append "status", 200, 20 ' adVarChar
approvalPendingRS.Open

' 실제 데이터 조회 시도
Dim approvalPendingCmd
Set approvalPendingCmd = Server.CreateObject("ADODB.Command")
approvalPendingCmd.ActiveConnection = db
approvalPendingCmd.CommandText = "SELECT al.target_table_name, cu.usage_id as doc_id, cu.usage_date as doc_date, " & _
                               "ISNULL(cu.title, cu.store_name) as title, cu.amount, " & _
                               "u.name AS requester_name, d.name AS department_name, al.status " & _
                               "FROM dbo.ApprovalLogs al " & _
                               "JOIN dbo.CardUsage cu ON al.target_id = cu.usage_id AND al.target_table_name = 'CardUsage' " & _
                               "JOIN dbo.Users u ON cu.user_id = u.user_id " & _
                               "LEFT JOIN dbo.Department d ON u.department_id = d.department_id " & _
                               "WHERE al.approver_id = ? " & _
                               "AND al.status IN ('대기', '반려') " & _
                               
                               "UNION ALL " & _
                               
                               "SELECT al.target_table_name, vr.request_id as doc_id, vr.start_date as doc_date, " & _
                               "ISNULL(vr.title, vr.purpose) as title, (vr.distance * 2000) as amount, " & _
                               "u.name AS requester_name, d.name AS department_name, al.status " & _
                               "FROM dbo.ApprovalLogs al " & _
                               "JOIN dbo.VehicleRequests vr ON al.target_id = vr.request_id AND al.target_table_name = 'VehicleRequests' " & _
                               "JOIN dbo.Users u ON vr.user_id = u.user_id " & _
                               "LEFT JOIN dbo.Department d ON u.department_id = d.department_id " & _
                               "WHERE al.approver_id = ? " & _
                               "AND al.status IN ('대기', '반려') " & _
                               
                               "ORDER BY doc_date DESC"
approvalPendingCmd.Parameters.Append approvalPendingCmd.CreateParameter("@approver_id1", 200, 1, 30, Session("user_id"))
approvalPendingCmd.Parameters.Append approvalPendingCmd.CreateParameter("@approver_id2", 200, 1, 30, Session("user_id"))

On Error Resume Next
Dim tempApprovalPendingRS
Set tempApprovalPendingRS = approvalPendingCmd.Execute()

' 데이터 조회 성공 시 임시 레코드셋에서 실제 레코드셋으로 데이터 복사
If Err.Number = 0 And Not tempApprovalPendingRS.EOF Then
    ' 기존 빈 레코드셋 닫고 정상 레코드셋으로 대체
    approvalPendingRS.Close
    Set approvalPendingRS = tempApprovalPendingRS
End If
On Error GoTo 0

' 결재 완료 문서 조회
Dim approvalCompletedRS
Set approvalCompletedRS = Server.CreateObject("ADODB.Recordset")

' 빈 레코드셋 초기화 (오류 발생 시 이 레코드셋 사용)
approvalCompletedRS.Fields.Append "target_table_name", 200, 30 ' adVarChar
approvalCompletedRS.Fields.Append "doc_id", 3 ' adInteger
approvalCompletedRS.Fields.Append "doc_date", 7 ' adDate
approvalCompletedRS.Fields.Append "title", 200, 100 ' adVarChar
approvalCompletedRS.Fields.Append "amount", 6 ' adCurrency
approvalCompletedRS.Fields.Append "requester_name", 200, 50 ' adVarChar
approvalCompletedRS.Fields.Append "department_name", 200, 50 ' adVarChar
approvalCompletedRS.Fields.Append "status", 200, 20 ' adVarChar
approvalCompletedRS.Fields.Append "approved_at", 7 ' adDate
approvalCompletedRS.Open

' 실제 데이터 조회 시도
Dim approvalCompletedCmd
Set approvalCompletedCmd = Server.CreateObject("ADODB.Command")
approvalCompletedCmd.ActiveConnection = db
approvalCompletedCmd.CommandText = "SELECT al.target_table_name, cu.usage_id as doc_id, cu.usage_date as doc_date, " & _
                                 "cu.store_name as title, cu.amount, " & _
                                 "u.name AS requester_name, d.name AS department_name, al.status, " & _
                                 "al.approved_at " & _
                                 "FROM dbo.ApprovalLogs al " & _
                                 "JOIN dbo.CardUsage cu ON al.target_id = cu.usage_id AND al.target_table_name = 'CardUsage' " & _
                                 "JOIN dbo.Users u ON cu.user_id = u.user_id " & _
                                 "LEFT JOIN dbo.Department d ON u.department_id = d.department_id " & _
                                 "WHERE al.approver_id = ? " & _
                                 "AND al.status IN ('승인') " & _
                                 
                                 "UNION ALL " & _
                                 
                                 "SELECT al.target_table_name, vr.request_id as doc_id, vr.start_date as doc_date, " & _
                                 "ISNULL(vr.title, vr.purpose) as title, (vr.distance * 2000) as amount, " & _
                                 "u.name AS requester_name, d.name AS department_name, al.status, " & _
                                 "al.approved_at " & _
                                 "FROM dbo.ApprovalLogs al " & _
                                 "JOIN dbo.VehicleRequests vr ON al.target_id = vr.request_id AND al.target_table_name = 'VehicleRequests' " & _
                                 "JOIN dbo.Users u ON vr.user_id = u.user_id " & _
                                 "LEFT JOIN dbo.Department d ON u.department_id = d.department_id " & _
                                 "WHERE al.approver_id = ? " & _
                                 "AND al.status IN ('승인') " & _
                                 
                                 "ORDER BY approved_at DESC"
approvalCompletedCmd.Parameters.Append approvalCompletedCmd.CreateParameter("@approver_id1", 200, 1, 30, Session("user_id"))
approvalCompletedCmd.Parameters.Append approvalCompletedCmd.CreateParameter("@approver_id2", 200, 1, 30, Session("user_id"))

On Error Resume Next
Dim tempApprovalCompletedRS
Set tempApprovalCompletedRS = approvalCompletedCmd.Execute()

' 데이터 조회 성공 시 임시 레코드셋에서 실제 레코드셋으로 데이터 복사
If Err.Number = 0 And Not tempApprovalCompletedRS.EOF Then
    ' 기존 빈 레코드셋 닫고 정상 레코드셋으로 대체
    approvalCompletedRS.Close
    Set approvalCompletedRS = tempApprovalCompletedRS
End If
On Error GoTo 0
%>
<!--#include file="../includes/header.asp"-->

<div class="dashboard-container">
    <!-- 환영 메시지 섹션 -->
    <div class="welcome-section">
        <div class="welcome-content">
            <h1>환영합니다, <%= userName %> 님</h1>
            <p class="department"><%= userDeptName %> 소속</p>
            <p class="welcome-text">카드 지출 결의 및 개인차량 이용 내력 관리 시스템</p>
        </div>
        <div class="welcome-actions">
            <a href="/contents/card_car_used/pages/my_profile.asp" class="btn btn-outline-light">
                <i class="fas fa-user"></i> 내 정보 보기
            </a>
        </div>
    </div>

    <!-- 결재 섹션 -->
    <div class="section-container">
        <div class="section-row">
            <!-- 결재 대기 문서 -->
            <div class="section-card">
                <div class="card-header">
                    <h2>결재 대기 문서</h2>
                </div>
                <div class="card-body">
                    <%
                    ' 결재 대기 문서 조회 (카드 사용 내역)
                    Dim pendingSQL
                    pendingSQL = "SELECT al.target_table_name, cu.usage_id as doc_id, cu.usage_date as doc_date, " & _
                               "cu.store_name as title, cu.amount, " & _
                               "u.name AS requester_name, d.name AS department_name, al.status " & _
                               "FROM dbo.ApprovalLogs al " & _
                               "JOIN dbo.CardUsage cu ON al.target_id = cu.usage_id AND al.target_table_name = 'CardUsage' " & _
                               "JOIN dbo.Users u ON cu.user_id = u.user_id " & _
                               "LEFT JOIN dbo.Department d ON u.department_id = d.department_id " & _
                               "WHERE al.approver_id = '" & Session("user_id") & "' " & _
                               "AND al.status IN ('대기', '반려') " & _
                               
                               "UNION ALL " & _
                               
                               "SELECT al.target_table_name, vr.request_id as doc_id, vr.start_date as doc_date, " & _
                               "ISNULL(vr.title, vr.purpose) as title, (vr.distance * 2000) as amount, " & _
                               "u.name AS requester_name, d.name AS department_name, al.status " & _
                               "FROM dbo.ApprovalLogs al " & _
                               "JOIN dbo.VehicleRequests vr ON al.target_id = vr.request_id AND al.target_table_name = 'VehicleRequests' " & _
                               "JOIN dbo.Users u ON vr.user_id = u.user_id " & _
                               "LEFT JOIN dbo.Department d ON u.department_id = d.department_id " & _
                               "WHERE al.approver_id = '" & Session("user_id") & "' " & _
                               "AND al.status IN ('대기', '반려') " & _
                               
                               "ORDER BY doc_date DESC"
                    
                    If Not approvalPendingRS.EOF Then
                        Do While Not approvalPendingRS.EOF
                    %>
                        <div class="approval-item">
                            <div class="approval-content">
                                <div class="approval-header">
                                    <span class="store-name"><%= approvalPendingRS("title") %></span>
                                    <span class="amount"><%= FormatNumber(approvalPendingRS("amount")) %>원</span>
                                </div>
                                <div class="approval-info">
                                    <span class="requester"><%= approvalPendingRS("requester_name") %> (<%= approvalPendingRS("department_name") %>)</span>
                                    <span class="date"><%= FormatDateTime(approvalPendingRS("doc_date"), 2) %></span>
                                </div>
                                <% 
                                Dim statusClass
                                Select Case approvalPendingRS("status")
                                    Case "승인"
                                        statusClass = "status-approved"
                                    Case "반려"
                                        statusClass = "status-rejected"
                                    Case "대기"
                                        statusClass = "status-pending"
                                    Case Else
                                        statusClass = "status-other"
                                End Select
                                %>
                                <span class="status-badge <%= statusClass %>"><%= approvalPendingRS("status") %></span>
                                <span class="doc-type-badge"><%= IIf(approvalPendingRS("target_table_name")="CardUsage", "카드", "차량") %></span>
                            </div>
                            <a href="approval_detail.asp?id=<%= approvalPendingRS("doc_id") %>&type=<%= approvalPendingRS("target_table_name") %>" class="btn btn-sm btn-outline-primary">상세보기</a>
                        </div>
                    <%
                            approvalPendingRS.MoveNext
                        Loop
                    Else
                    %>
                        <div class="no-data">결재 대기 중인 문서가 없습니다.</div>
                    <%
                    End If
                    %>
                </div>
                <div class="card-footer">
                    <a href="pending_approvals.asp" class="btn btn-outline-primary">모든 내역 보기</a>
                </div>
            </div>

            <!-- 결재 완료 문서 -->
            <div class="section-card">
                <div class="card-header">
                    <h2>결재 완료 문서</h2>
                </div>
                <div class="card-body">
                    <%
                    ' 결재 완료 문서 조회
                    Dim completedSQL
                    completedSQL = "SELECT al.target_table_name, cu.usage_id as doc_id, cu.usage_date as doc_date, " & _
                                 "cu.store_name as title, cu.amount, " & _
                                 "u.name AS requester_name, d.name AS department_name, al.status, " & _
                                 "al.approved_at " & _
                                 "FROM dbo.ApprovalLogs al " & _
                                 "JOIN dbo.CardUsage cu ON al.target_id = cu.usage_id AND al.target_table_name = 'CardUsage' " & _
                                 "JOIN dbo.Users u ON cu.user_id = u.user_id " & _
                                 "LEFT JOIN dbo.Department d ON u.department_id = d.department_id " & _
                                 "WHERE al.approver_id = '" & Session("user_id") & "' " & _
                                 "AND al.status IN ('승인') " & _
                                 
                                 "UNION ALL " & _
                                 
                                 "SELECT al.target_table_name, vr.request_id as doc_id, vr.start_date as doc_date, " & _
                                 "ISNULL(vr.title, vr.purpose) as title, (vr.distance * 2000) as amount, " & _
                                 "u.name AS requester_name, d.name AS department_name, al.status, " & _
                                 "al.approved_at " & _
                                 "FROM dbo.ApprovalLogs al " & _
                                 "JOIN dbo.VehicleRequests vr ON al.target_id = vr.request_id AND al.target_table_name = 'VehicleRequests' " & _
                                 "JOIN dbo.Users u ON vr.user_id = u.user_id " & _
                                 "LEFT JOIN dbo.Department d ON u.department_id = d.department_id " & _
                                 "WHERE al.approver_id = '" & Session("user_id") & "' " & _
                                 "AND al.status IN ('승인') " & _
                                 
                                 "ORDER BY approved_at DESC"
                    
                    If Not approvalCompletedRS.EOF Then
                        Do While Not approvalCompletedRS.EOF
                    %>
                        <div class="approval-item">
                            <div class="approval-content">
                                <div class="approval-header">
                                    <span class="store-name"><%= approvalCompletedRS("title") %></span>
                                    <span class="amount"><%= FormatNumber(approvalCompletedRS("amount")) %>원</span>
                                </div>
                                <div class="approval-info">
                                    <span class="requester"><%= approvalCompletedRS("requester_name") %> (<%= approvalCompletedRS("department_name") %>)</span>
                                    <span class="date"><%= FormatDateTime(approvalCompletedRS("approved_at"), 2) %></span>
                                </div>
                                <span class="status-badge <%= IIf(approvalCompletedRS("status")="승인", "status-approved", "status-rejected") %>">
                                    <%= approvalCompletedRS("status") %>
                                </span>
                                <span class="doc-type-badge"><%= IIf(approvalCompletedRS("target_table_name")="CardUsage", "카드", "차량") %></span>
                            </div>
                            <a href="approval_detail.asp?id=<%= approvalCompletedRS("doc_id") %>&type=<%= approvalCompletedRS("target_table_name") %>" class="btn btn-sm btn-outline-primary">상세보기</a>
                        </div>
                    <%
                            approvalCompletedRS.MoveNext
                        Loop
                    Else
                    %>
                        <div class="no-data">결재 완료한 문서가 없습니다.</div>
                    <%
                    End If
                    %>
                </div>
                <div class="card-footer">
                    <a href="completed_approvals.asp" class="btn btn-outline-primary">모든 내역 보기</a>
                </div>
            </div>
        </div>

        <div class="section-row">
            <!-- 카드 사용 내역 -->
            <div class="section-card">
                <div class="card-header">
                    <h2>최근 카드 사용 내역</h2>
                </div>
                <div class="card-body">
                    <% If recentCardRS.EOF Then %>
                        <div class="no-data">최근 카드 사용 내역이 없습니다.</div>
                    <% Else %>
                        <% Do While Not recentCardRS.EOF %>
                            <div class="usage-item">
                                <div class="usage-content">
                                    <div class="usage-header">
                                        <span class="card-name"><%= recentCardRS("account_name") %></span>
                                        <span class="amount"><%= FormatNumber(recentCardRS("amount")) %>원</span>
                                    </div>
                                    <div class="usage-date">
                                        <%= FormatDateTime(recentCardRS("usage_date"), 2) %>
                                    </div>
                                </div>
                            </div>
                        <% 
                            recentCardRS.MoveNext
                            Loop 
                        %>
                    <% End If %>
                </div>
                <div class="card-footer">
                    <a href="/contents/card_car_used/pages/card_usage.asp" class="btn btn-outline-primary">모든 내역 보기</a>
                    <a href="/contents/card_car_used/pages/card_usage_add.asp" class="btn btn-primary">새 내역 등록</a>
                </div>
            </div>

            <!-- 차량 사용 내역 -->
            <div class="section-card">
                <div class="card-header">
                    <h2>최근 차량 사용 신청</h2>
                </div>
                <div class="card-body">
                    <% If recentVehicleRS.EOF Then %>
                        <div class="no-data">최근 차량 사용 신청 내역이 없습니다.</div>
                    <% Else %>
                        <% Do While Not recentVehicleRS.EOF %>
                            <div class="usage-item">
                                <div class="usage-content">
                                    <div class="usage-header">
                                        <span class="purpose"><%= recentVehicleRS("purpose") %></span>
                                        <span class="amount"><%= FormatNumber(CDbl(recentVehicleRS("distance")) * CDbl(recentVehicleRS("fuel_rate"))) %>원</span>
                                    </div>
                                    <div class="usage-subheader">
                                        <% 
                                        Select Case recentVehicleRS("approval_status")
                                            Case "승인"
                                                Response.Write "<span class='status-badge status-approved'>"
                                            Case "반려"
                                                Response.Write "<span class='status-badge status-rejected'>"
                                            Case "대기"
                                                Response.Write "<span class='status-badge status-pending'>"
                                            Case Else
                                                Response.Write "<span class='status-badge status-other'>"
                                        End Select
                                        Response.Write recentVehicleRS("approval_status") & "</span>"
                                        %>
                                        <span class="distance-info"><%= FormatNumber(recentVehicleRS("distance")) %>km</span>
                                    </div>
                                    <div class="usage-date">
                                        <%= FormatDateTime(recentVehicleRS("request_date"), 2) %>
                                    </div>
                                </div>
                            </div>
                        <% 
                            recentVehicleRS.MoveNext
                            Loop 
                        %>
                    <% End If %>
                </div>
                <div class="card-footer">
                    <a href="/contents/card_car_used/pages/vehicle_request.asp" class="btn btn-outline-primary">모든 내역 보기</a>
                    <a href="/contents/card_car_used/pages/vehicle_request_add.asp" class="btn btn-primary">새 신청서 작성</a>
                </div>
            </div>
        </div>
    </div>
</div>

<style>
.dashboard-container {
    padding: 20px;
    max-width: 1400px;
    margin: 0 auto;
}

.welcome-section {
    background: linear-gradient(135deg, #4A90E2 0%, #2C3E50 100%);
    color: white;
    padding: 30px;
    border-radius: 10px;
    margin-bottom: 30px;
    display: flex;
    justify-content: space-between;
    align-items: center;
}

.welcome-content h1 {
    font-size: 24px;
    margin: 0;
    font-weight: 600;
}

.welcome-content .department {
    font-size: 16px;
    opacity: 0.9;
    margin: 5px 0;
}

.welcome-content .welcome-text {
    font-size: 14px;
    opacity: 0.8;
    margin: 5px 0;
}

.section-container {
    display: flex;
    flex-direction: column;
    gap: 20px;
}

.section-row {
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(400px, 1fr));
    gap: 20px;
}

.section-card {
    background: white;
    border-radius: 10px;
    box-shadow: 0 2px 4px rgba(0,0,0,0.1);
    overflow: hidden;
}

.card-header {
    padding: 15px 20px;
    border-bottom: 1px solid #eee;
}

.card-header h2 {
    margin: 0;
    font-size: 18px;
    font-weight: 600;
    color: #333;
}

.card-body {
    padding: 20px;
    max-height: 400px;
    overflow-y: auto;
}

.card-footer {
    padding: 15px 20px;
    border-top: 1px solid #eee;
    display: flex;
    justify-content: space-between;
    gap: 10px;
}

.approval-item, .usage-item {
    padding: 15px;
    border-bottom: 1px solid #eee;
    display: flex;
    justify-content: space-between;
    align-items: center;
    gap: 15px;
}

.approval-item:last-child, .usage-item:last-child {
    border-bottom: none;
}

.approval-content, .usage-content {
    flex: 1;
}

.approval-header, .usage-header {
    display: flex;
    justify-content: space-between;
    margin-bottom: 5px;
}

.usage-subheader {
    display: flex;
    justify-content: space-between;
    margin-bottom: 5px;
    align-items: center;
}

.store-name, .card-name, .purpose {
    font-weight: 500;
    color: #333;
}

.amount {
    font-weight: 600;
    color: #2C3E50;
}

.distance-info {
    font-size: 12px;
    color: #666;
}

.approval-info, .usage-date {
    font-size: 13px;
    color: #666;
}

.status-badge {
    display: inline-block;
    padding: 4px 8px;
    border-radius: 4px;
    font-size: 12px;
    font-weight: 500;
}

.doc-type-badge {
    display: inline-block;
    padding: 4px 8px;
    border-radius: 4px;
    font-size: 12px;
    font-weight: 500;
    background-color: #E3F0FF;
    color: #1B73E8;
    margin-left: 4px;
}

.status-approved {
    background-color: #E3F9E5;
    color: #1B873F;
}

.status-rejected {
    background-color: #FFE9E9;
    color: #DA3633;
}

.status-pending {
    background-color: #FFF8E6;
    color: #D4A72C;
}

.status-other {
    background-color: #F6F8FA;
    color: #57606A;
}

.no-data {
    text-align: center;
    color: #666;
    padding: 30px;
    font-size: 14px;
}

.btn {
    padding: 8px 16px;
    border-radius: 6px;
    font-size: 14px;
    font-weight: 500;
    transition: all 0.2s;
}

.btn-sm {
    padding: 4px 12px;
    font-size: 13px;
}

.btn-primary {
    background-color: #4A90E2;
    border-color: #4A90E2;
    color: white;
}

.btn-primary:hover {
    background-color: #357ABD;
    border-color: #357ABD;
}

.btn-outline-primary {
    color: #4A90E2;
    border-color: #4A90E2;
    background-color: transparent;
}

.btn-outline-primary:hover {
    background-color: #4A90E2;
    color: white;
}

.btn-outline-light {
    color: white;
    border-color: white;
    background-color: transparent;
}

.btn-outline-light:hover {
    background-color: rgba(255,255,255,0.1);
}
</style>

<!--#include file="../includes/footer.asp"--> 