<%@ Language="VBScript" CodePage="65001" %>
<%
Response.CodePage = 65001
Response.CharSet = "utf-8"
%>

<!--#include virtual="/contents/card_car_used/db.asp"-->
<!--#include virtual="/contents/card_car_used/includes/functions.asp"-->
<%
' 로그인 체크
If Not IsAuthenticated() Then
    RedirectTo("/contents/card_car_used/index.asp")
End If

' 페이징 처리
Dim pageSize, currentPage
pageSize = 10 ' 페이지당 표시할 항목 수
currentPage = Request.QueryString("page")
If currentPage = "" Then
    currentPage = 1
Else
    currentPage = CInt(currentPage)
End If

' 필터링
Dim status
status = Request.QueryString("status")
If status = "" Then
    status = "all"
End If

' 결재 완료 문서 조회
Dim totalCount, totalPages
Dim countSQL, completedSQL, rs
Dim statusCondition

' 상태 조건 설정
Select Case status
    Case "approved"
        statusCondition = "al.status = '승인'"
    Case "rejected"
        statusCondition = "al.status = '반려'"
    Case Else
        statusCondition = "al.status IN ('승인', '반려')"
End Select

' 전체 건수 조회
countSQL = "SELECT COUNT(*) AS cnt FROM dbo.CardUsage cu " & _
           "JOIN dbo.Users u ON cu.user_id = u.user_id " & _
           "LEFT JOIN dbo.Department d ON u.department_id = d.department_id " & _
           "JOIN dbo.ApprovalLogs al ON cu.usage_id = al.target_id " & _
           "WHERE al.target_table_name = 'CardUsage' " & _
           "AND al.approver_id = '" & Session("user_id") & "' " & _
           "AND " & statusCondition

Set rs = db.Execute(countSQL)
totalCount = rs("cnt")
totalPages = (totalCount + pageSize - 1) \ pageSize

Dim offsetVal
offsetVal = (currentPage - 1) * pageSize

completedSQL = "SELECT TOP " & pageSize & " cu.usage_id, cu.usage_date, cu.store_name, cu.amount, cu.purpose, " & _
              "u.name AS requester_name, d.name AS department_name, " & _
              "al.status, al.comments, al.approved_at " & _
              "FROM dbo.CardUsage cu " & _
              "JOIN dbo.Users u ON cu.user_id = u.user_id " & _
              "LEFT JOIN dbo.Department d ON u.department_id = d.department_id " & _
              "JOIN dbo.ApprovalLogs al ON cu.usage_id = al.target_id " & _
              "WHERE al.target_table_name = 'CardUsage' " & _
              "AND al.approver_id = '" & Session("user_id") & "' " & _
              "AND " & statusCondition & " " & _
              "AND cu.usage_id NOT IN (" & _
              "    SELECT TOP " & offsetVal & " cu2.usage_id " & _
              "    FROM dbo.CardUsage cu2 " & _
              "    JOIN dbo.Users u2 ON cu2.user_id = u2.user_id " & _
              "    LEFT JOIN dbo.Department d2 ON u2.department_id = d2.department_id " & _
              "    JOIN dbo.ApprovalLogs al2 ON cu2.usage_id = al2.target_id " & _
              "    WHERE al2.target_table_name = 'CardUsage' " & _
              "    AND al2.approver_id = '" & Session("user_id") & "' " & _
              "    AND " & statusCondition & " " & _
              "    ORDER BY al2.approved_at DESC" & _
              ") " & _
              "ORDER BY al.approved_at DESC"


Set rs = db99.Execute(completedSQL)
%>

<!--#include virtual="/contents/card_car_used/includes/header.asp"-->

<div class="container mt-4">
    <div class="d-flex justify-content-between align-items-center mb-4">
        <h2>결재 완료 문서 목록</h2>
        <a href="dashboard.asp" class="btn btn-outline-secondary">
            <i class="fas fa-arrow-left"></i> 대시보드로 돌아가기
        </a>
    </div>

    <div class="card">
        <div class="card-header">
            <div class="d-flex justify-content-between align-items-center">
                <div class="btn-group">
                    <a href="?status=all" class="btn btn-outline-primary <%= IIf(status="all" Or status="", "active", "") %>">전체</a>
                    <a href="?status=approved" class="btn btn-outline-primary <%= IIf(status="approved", "active", "") %>">승인</a>
                    <a href="?status=rejected" class="btn btn-outline-primary <%= IIf(status="rejected", "active", "") %>">반려</a>
                </div>
            </div>
        </div>
        <div class="card-body">
            <% If rs.EOF Then %>
                <div class="text-center py-5">
                    <p class="text-muted">결재 완료된 문서가 없습니다.</p>
                </div>
            <% Else %>
                <div class="table-responsive">
                    <table class="table table-hover">
                        <thead>
                            <tr>
                                <th>처리일</th>
                                <th>신청자</th>
                                <th>부서</th>
                                <th>사용처</th>
                                <th>금액</th>
                                <th>용도</th>
                                <th>상태</th>
                                <th>의견</th>
                                <th>상세</th>
                            </tr>
                        </thead>
                        <tbody>
                            <% Do While Not rs.EOF %>
                                <tr>
                                    <td><%= FormatDateTime(rs("approved_at"), 2) %></td>
                                    <td><%= rs("requester_name") %></td>
                                    <td><%= rs("department_name") %></td>
                                    <td><%= rs("store_name") %></td>
                                    <td class="text-right">
                                        <% 
                                        If Not IsNull(rs("amount")) Then
                                            Response.Write FormatNumber(rs("amount"), 0) & "원"
                                        Else
                                            Response.Write "0원"
                                        End If
                                        %>
                                    </td>
                                    <td><%= Left(rs("purpose"), 20) & IIf(Len(rs("purpose")) > 20, "...", "") %></td>
                                    <td>
                                        <span class="badge badge-<%= IIf(rs("status")="승인", "success", "danger") %>">
                                            <%= rs("status") %>
                                        </span>
                                    </td>
                                    <td>
                                        <% If Not IsNull(rs("comments")) And rs("comments") <> "" Then %>
                                            <span class="text-muted" title="<%= rs("comments") %>">
                                                <%= Left(rs("comments"), 10) & IIf(Len(rs("comments")) > 10, "...", "") %>
                                            </span>
                                        <% End If %>
                                    </td>
                                    <td>
                                        <a href="approval_detail.asp?id=<%= rs("usage_id") %>" class="btn btn-sm btn-outline-primary">상세보기</a>
                                    </td>
                                </tr>
                            <%
                                rs.MoveNext
                                Loop
                            %>
                        </tbody>
                    </table>
                </div>

                <!-- 페이징 -->
                <% If totalPages > 1 Then %>
                    <div class="d-flex justify-content-center mt-4">
                        <nav aria-label="Page navigation">
                            <ul class="pagination">
                                <% If currentPage > 1 Then %>
                                    <li class="page-item">
                                        <a class="page-link" href="?page=<%= currentPage - 1 %>&status=<%= status %>">&laquo; 이전</a>
                                    </li>
                                <% End If %>

                                <% 
                                Dim startPage, endPage
                                startPage = ((currentPage - 1) \ 5) * 5 + 1
                                endPage = Min(startPage + 4, totalPages)

                                For i = startPage To endPage
                                %>
                                    <li class="page-item <%= IIf(i = currentPage, "active", "") %>">
                                        <a class="page-link" href="?page=<%= i %>&status=<%= status %>"><%= i %></a>
                                    </li>
                                <% Next %>

                                <% If currentPage < totalPages Then %>
                                    <li class="page-item">
                                        <a class="page-link" href="?page=<%= currentPage + 1 %>&status=<%= status %>">다음 &raquo;</a>
                                    </li>
                                <% End If %>
                            </ul>
                        </nav>
                    </div>
                <% End If %>
            <% End If %>
        </div>
    </div>
</div>

<style>
.badge {
    padding: 6px 10px;
    font-weight: 500;
}
.badge-success {
    background-color: #E3F9E5;
    color: #1B873F;
}
.badge-danger {
    background-color: #FFE9E9;
    color: #DA3633;
}
.table td {
    vertical-align: middle;
}
.btn-group .btn.active {
    background-color: #4A90E2;
    color: white;
}
</style>

<!--#include virtual="/contents/card_car_used/includes/footer.asp"--> 