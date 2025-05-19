<%@ Language="VBScript" CodePage="65001" %>
<%
Response.CodePage = 65001
Response.CharSet = "utf-8"

' 오류 처리 초기화
On Error Resume Next
%>

<!--#include file="../db.asp"-->
<!--#include file="../includes/functions.asp"-->

<%
' 로그인 상태 확인
If Not IsAuthenticated() Then
    Response.Write "<script>alert('로그인이 필요합니다.'); location.href='../index.asp';</script>"
    Response.End
End If

' 현재 로그인한 사용자 정보
Dim currentUserId
currentUserId = Session("user_id")

' 1. 내가 결재해야 할 문서 조회
Dim pendingSQL, pendingRS
pendingSQL = "SELECT cu.usage_id, cu.card_name, cu.store_name, cu.amount, cu.usage_date, " & _
            "cu.purpose, cu.approval_status, u.name as requester_name, " & _
            "al.approval_step, al.status as approval_status, " & _
            "d.name as dept_name, j.name as job_grade_name " & _
            "FROM " & dbSchema & ".ApprovalLogs al " & _
            "INNER JOIN " & dbSchema & ".CardUsage cu ON al.target_id = cu.usage_id " & _
            "INNER JOIN " & dbSchema & ".Users u ON cu.user_id = u.user_id " & _
            "LEFT JOIN " & dbSchema & ".Department d ON u.department_id = d.department_id " & _
            "LEFT JOIN " & dbSchema & ".job_grade j ON u.job_grade = j.job_grade_id " & _
            "WHERE al.target_table_name = 'CardUsage' " & _
            "AND al.approver_id = '" & PreventSQLInjection(currentUserId) & "' " & _
            "AND al.status = '대기' " & _
            "AND (al.approval_step = 1 OR EXISTS (" & _
            "    SELECT 1 FROM " & dbSchema & ".ApprovalLogs prev " & _
            "    WHERE prev.target_table_name = 'CardUsage' " & _
            "    AND prev.target_id = al.target_id " & _
            "    AND prev.approval_step = al.approval_step - 1 " & _
            "    AND prev.status = '승인'" & _
            ")) " & _
            "ORDER BY cu.created_at DESC"

Set pendingRS = db.Execute(pendingSQL)

' 2. 내가 신청한 문서 조회
Dim myRequestSQL, myRequestRS
myRequestSQL = "SELECT cu.*, u.name as requester_name, " & _
              "d.name as dept_name, " & _
              "STUFF((" & _
              "    SELECT ', ' + approver.name " & _
              "    FROM " & dbSchema & ".ApprovalLogs al " & _
              "    INNER JOIN " & dbSchema & ".Users approver ON al.approver_id = approver.user_id " & _
              "    WHERE al.target_table_name = 'CardUsage' " & _
              "    AND al.target_id = cu.usage_id " & _
              "    ORDER BY al.approval_step " & _
              "    FOR XML PATH('')" & _
              "), 1, 2, '') as approver_names " & _
              "FROM " & dbSchema & ".CardUsage cu " & _
              "INNER JOIN " & dbSchema & ".Users u ON cu.user_id = u.user_id " & _
              "LEFT JOIN " & dbSchema & ".Department d ON u.department_id = d.department_id " & _
              "WHERE cu.user_id = '" & PreventSQLInjection(currentUserId) & "' " & _
              "ORDER BY cu.created_at DESC"

Set myRequestRS = db.Execute(myRequestSQL)

' 3. 최근 처리한 결재 내역 조회
Dim processedSQL, processedRS
processedSQL = "SELECT TOP 5 cu.usage_id, cu.card_name, cu.store_name, cu.amount, " & _
              "u.name as requester_name, d.name as dept_name, " & _
              "al.status as approval_status, al.processed_at, al.comment " & _
              "FROM " & dbSchema & ".ApprovalLogs al " & _
              "INNER JOIN " & dbSchema & ".CardUsage cu ON al.target_id = cu.usage_id " & _
              "INNER JOIN " & dbSchema & ".Users u ON cu.user_id = u.user_id " & _
              "LEFT JOIN " & dbSchema & ".Department d ON u.department_id = d.department_id " & _
              "WHERE al.approver_id = '" & PreventSQLInjection(currentUserId) & "' " & _
              "AND al.status IN ('승인', '반려') " & _
              "ORDER BY al.processed_at DESC"

Set processedRS = db.Execute(processedSQL)
%>

<!--#include file="../includes/header.asp"-->

<div class="container mt-4">
    <h2 class="mb-4">대시보드</h2>
    
    <!-- 1. 내가 결재해야 할 문서 -->
    <div class="card mb-4">
        <div class="card-header bg-primary text-white">
            <h5 class="card-title mb-0">
                <i class="fas fa-clock"></i> 내가 결재해야 할 문서
            </h5>
        </div>
        <div class="card-body">
            <% If pendingRS.EOF Then %>
            <div class="alert alert-info">
                결재 대기 중인 문서가 없습니다.
            </div>
            <% Else %>
            <div class="table-responsive">
                <table class="table table-hover">
                    <thead>
                        <tr>
                            <th>신청자</th>
                            <th>부서</th>
                            <th>카드명</th>
                            <th>사용처</th>
                            <th>금액</th>
                            <th>결재단계</th>
                            <th>처리</th>
                        </tr>
                    </thead>
                    <tbody>
                        <% Do Until pendingRS.EOF %>
                        <tr>
                            <td>
                                <%= pendingRS("requester_name") %>
                                (<%= pendingRS("job_grade_name") %>)
                            </td>
                            <td><%= pendingRS("dept_name") %></td>
                            <td><%= pendingRS("card_name") %></td>
                            <td><%= pendingRS("store_name") %></td>
                            <td class="text-right"><%= FormatNumber(pendingRS("amount"), 0) %>원</td>
                            <td><%= pendingRS("approval_step") %>차 결재</td>
                            <td>
                                <a href="approval_detail.asp?id=<%= pendingRS("usage_id") %>" class="btn btn-sm btn-primary">
                                    결재하기
                                </a>
                            </td>
                        </tr>
                        <%
                            pendingRS.MoveNext
                            Loop
                        %>
                    </tbody>
                </table>
            </div>
            <% End If %>
        </div>
    </div>
    
    <!-- 2. 내가 신청한 문서 -->
    <div class="card mb-4">
        <div class="card-header bg-success text-white">
            <h5 class="card-title mb-0">
                <i class="fas fa-file-alt"></i> 내가 신청한 문서
            </h5>
        </div>
        <div class="card-body">
            <% If myRequestRS.EOF Then %>
            <div class="alert alert-info">
                신청한 문서가 없습니다.
            </div>
            <% Else %>
            <div class="table-responsive">
                <table class="table table-hover">
                    <thead>
                        <tr>
                            <th>카드명</th>
                            <th>사용처</th>
                            <th>금액</th>
                            <th>사용일자</th>
                            <th>결재자</th>
                            <th>상태</th>
                            <th>상세</th>
                        </tr>
                    </thead>
                    <tbody>
                        <% Do Until myRequestRS.EOF %>
                        <tr>
                            <td><%= myRequestRS("card_name") %></td>
                            <td><%= myRequestRS("store_name") %></td>
                            <td class="text-right"><%= FormatNumber(myRequestRS("amount"), 0) %>원</td>
                            <td><%= FormatDateTime(myRequestRS("usage_date"), 2) %></td>
                            <td><%= IIf(IsNull(myRequestRS("approver_names")), "-", myRequestRS("approver_names")) %></td>
                            <td>
                                <span class="badge badge-<%= GetStatusClass(myRequestRS("approval_status")) %>">
                                    <%= myRequestRS("approval_status") %>
                                </span>
                            </td>
                            <td>
                                <a href="approval_detail.asp?id=<%= myRequestRS("usage_id") %>" class="btn btn-sm btn-info">
                                    상세보기
                                </a>
                            </td>
                        </tr>
                        <%
                            myRequestRS.MoveNext
                            Loop
                        %>
                    </tbody>
                </table>
            </div>
            <% End If %>
        </div>
    </div>
    
    <!-- 3. 최근 처리한 결재 내역 -->
    <div class="card mb-4">
        <div class="card-header bg-info text-white">
            <h5 class="card-title mb-0">
                <i class="fas fa-history"></i> 최근 처리한 결재 내역
            </h5>
        </div>
        <div class="card-body">
            <% If processedRS.EOF Then %>
            <div class="alert alert-info">
                최근 처리한 결재 내역이 없습니다.
            </div>
            <% Else %>
            <div class="table-responsive">
                <table class="table table-hover">
                    <thead>
                        <tr>
                            <th>처리일시</th>
                            <th>신청자</th>
                            <th>카드명</th>
                            <th>금액</th>
                            <th>처리결과</th>
                            <th>의견</th>
                        </tr>
                    </thead>
                    <tbody>
                        <% Do Until processedRS.EOF %>
                        <tr>
                            <td><%= FormatDateTime(processedRS("processed_at"), 2) %></td>
                            <td>
                                <%= processedRS("requester_name") %>
                                (<%= processedRS("dept_name") %>)
                            </td>
                            <td><%= processedRS("card_name") %></td>
                            <td class="text-right"><%= FormatNumber(processedRS("amount"), 0) %>원</td>
                            <td>
                                <span class="badge badge-<%= GetStatusClass(processedRS("approval_status")) %>">
                                    <%= processedRS("approval_status") %>
                                </span>
                            </td>
                            <td>
                                <% If Not IsNull(processedRS("comment")) And processedRS("comment") <> "" Then %>
                                <span class="text-muted"><%= processedRS("comment") %></span>
                                <% End If %>
                            </td>
                        </tr>
                        <%
                            processedRS.MoveNext
                            Loop
                        %>
                    </tbody>
                </table>
            </div>
            <% End If %>
        </div>
    </div>
</div>

<%
Function GetStatusClass(status)
    Select Case status
        Case "대기"
            GetStatusClass = "warning"
        Case "승인"
            GetStatusClass = "success"
        Case "반려"
            GetStatusClass = "danger"
        Case "진행중"
            GetStatusClass = "info"
        Case Else
            GetStatusClass = "secondary"
    End Select
End Function
%>

<style>
.badge {
    font-size: 0.9rem;
    padding: 0.5em 1em;
}
.badge-warning {
    background-color: #ffc107;
    color: #000;
}
.badge-success {
    background-color: #28a745;
    color: #fff;
}
.badge-danger {
    background-color: #dc3545;
    color: #fff;
}
.badge-info {
    background-color: #17a2b8;
    color: #fff;
}
.badge-secondary {
    background-color: #6c757d;
    color: #fff;
}
.table td {
    vertical-align: middle;
}
</style>

<!--#include file="../includes/footer.asp"--> 