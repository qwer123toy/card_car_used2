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

' 결재 대기 문서 조회
Dim totalCount, totalPages
Dim countSQL, pendingSQL, rs

' 전체 건수 조회
countSQL = "SELECT COUNT(*) AS cnt FROM dbo.CardUsage cu " & _
           "JOIN dbo.Users u ON cu.user_id = u.user_id " & _
           "LEFT JOIN dbo.Department d ON u.department_id = d.department_id " & _
           "JOIN dbo.ApprovalLogs al ON cu.usage_id = al.target_id " & _
           "WHERE al.target_table_name = 'CardUsage' " & _
           "AND al.approver_id = '" & Session("user_id") & "' " & _
           "AND al.status = '대기' " & _
           "AND (al.approval_step = 1 OR EXISTS (" & _
           "    SELECT 1 FROM dbo.ApprovalLogs prev " & _
           "    WHERE prev.target_table_name = 'CardUsage' " & _
           "    AND prev.target_id = al.target_id " & _
           "    AND prev.approval_step = al.approval_step - 1 " & _
           "    AND prev.status = '승인'" & _
           "))"

Set rs = db99.Execute(countSQL)
totalCount = rs("cnt")
totalPages = (totalCount + pageSize - 1) \ pageSize

Dim offsetVal
offsetVal = (currentPage - 1) * pageSize

pendingSQL = "SELECT TOP " & pageSize & " cu.usage_id, cu.usage_date, cu.store_name, cu.amount, cu.purpose, " & _
    "u.name AS requester_name, d.name AS department_name, al.status, " & _
    "al.created_at AS requested_at, al.approval_step " & _
    "FROM dbo.CardUsage cu " & _
    "JOIN dbo.Users u ON cu.user_id = u.user_id " & _
    "LEFT JOIN dbo.Department d ON u.department_id = d.department_id " & _
    "JOIN dbo.ApprovalLogs al ON cu.usage_id = al.target_id " & _
    "WHERE al.target_table_name = 'CardUsage' " & _
    "AND al.approver_id = '" & Session("user_id") & "' " & _
    "AND al.status = '대기' " & _
    "AND (al.approval_step = 1 OR EXISTS (" & _
        "SELECT 1 FROM dbo.ApprovalLogs prev " & _
        "WHERE prev.target_table_name = 'CardUsage' " & _
        "AND prev.target_id = al.target_id " & _
        "AND prev.approval_step = al.approval_step - 1 " & _
        "AND prev.status = '승인')) " & _
    "AND cu.usage_id NOT IN (" & _
        "SELECT TOP " & offsetVal & " cu2.usage_id " & _
        "FROM dbo.CardUsage cu2 " & _
        "JOIN dbo.Users u2 ON cu2.user_id = u2.user_id " & _
        "LEFT JOIN dbo.Department d2 ON u2.department_id = d2.department_id " & _
        "JOIN dbo.ApprovalLogs al2 ON cu2.usage_id = al2.target_id " & _
        "WHERE al2.target_table_name = 'CardUsage' " & _
        "AND al2.approver_id = '" & Session("user_id") & "' " & _
        "AND al2.status = '대기' " & _
        "AND (al2.approval_step = 1 OR EXISTS (" & _
            "SELECT 1 FROM dbo.ApprovalLogs prev2 " & _
            "WHERE prev2.target_table_name = 'CardUsage' " & _
            "AND prev2.target_id = al2.target_id " & _
            "AND prev2.approval_step = al2.approval_step - 1 " & _
            "AND prev2.status = '승인')) " & _
        "ORDER BY al2.created_at DESC" & _
    ") " & _
    "ORDER BY al.created_at DESC"


Set rs = db99.Execute(pendingSQL)
%>

<!--#include virtual="/contents/card_car_used/includes/header.asp"-->

<div class="container mt-4">
    <div class="d-flex justify-content-between align-items-center mb-4">
        <h2 class="page-title">결재 대기 문서 목록</h2>
        <a href="dashboard.asp" class="btn btn-secondary">
            <i class="fas fa-arrow-left me-2"></i>대시보드로 돌아가기
        </a>
    </div>

    <div class="card">
        <div class="card-body">
            <% If rs.EOF Then %>
                <div class="text-center py-5">
                    <p class="text-muted">결재 대기 중인 문서가 없습니다.</p>
                </div>
            <% Else %>
                <div class="table-responsive">
                    <table class="table table-hover align-middle">
                        <thead class="table-light">
                            <tr>
                                <th>신청일</th>
                                <th>신청자</th>
                                <th>부서</th>
                                <th>사용처</th>
                                <th class="text-end">금액</th>
                                <th>용도</th>
                                <th>상태</th>
                                <th>처리</th>
                            </tr>
                        </thead>
                        <tbody>
                            <% Do While Not rs.EOF %>
                                <tr>
                                    <td><%= FormatDateTime(rs("requested_at"), 2) %></td>
                                    <td><%= rs("requester_name") %></td>
                                    <td><%= rs("department_name") %></td>
                                    <td><%= rs("store_name") %></td>
                                    <td class="text-end">
                                        <% 
                                        If Not IsNull(rs("amount")) Then
                                            Response.Write FormatCurrency(rs("amount"), 0)
                                        Else
                                            Response.Write FormatCurrency(0, 0)
                                        End If
                                        %>
                                    </td>
                                    <td><%= Left(rs("purpose"), 20) & IIf(Len(rs("purpose")) > 20, "...", "") %></td>
                                    <td>
                                        <span class="badge rounded-pill bg-warning-subtle text-warning px-3 py-2">
                                            <%= rs("approval_step") %>차 결재 대기
                                        </span>
                                    </td>
                                    <td>
                                        <a href="approval_detail.asp?id=<%= rs("usage_id") %>" class="btn btn-sm btn-primary">
                                            결재하기
                                        </a>
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
                                        <a class="page-link" href="?page=<%= currentPage - 1 %>">&laquo;</a>
                                    </li>
                                <% End If %>

                                <% 
                                Dim startPage, endPage
                                startPage = ((currentPage - 1) \ 5) * 5 + 1
                                endPage = Min(startPage + 4, totalPages)

                                For i = startPage To endPage
                                %>
                                    <li class="page-item <%= IIf(i = currentPage, "active", "") %>">
                                        <a class="page-link" href="?page=<%= i %>"><%= i %></a>
                                    </li>
                                <% Next %>

                                <% If currentPage < totalPages Then %>
                                    <li class="page-item">
                                        <a class="page-link" href="?page=<%= currentPage + 1 %>">&raquo;</a>
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
.page-title {
    font-size: 1.75rem;
    font-weight: 600;
    color: #2C3E50;
    margin: 0;
}

.table {
    font-size: 0.95rem;
    margin-top: 1rem;
}

.table th {
    font-weight: 600;
    color: #2C3E50;
    background-color: #F8F9FA;
    border-bottom: 2px solid #E9ECEF;
}

.table td {
    padding: 1rem 0.75rem;
    vertical-align: middle;
    border-bottom: 1px solid #E9ECEF;
}

.badge {
    font-weight: 500;
    font-size: 0.85rem;
    padding: 0.5rem 1rem;
    border-radius: 2rem;
}

.bg-warning-subtle {
    background-color: #FFF8E6;
}

.text-warning {
    color: #D4A72C !important;
}

.btn-secondary {
    background-color: #6C757D;
    border-color: #6C757D;
    color: white;
    padding: 0.5rem 1rem;
}

.btn-secondary:hover {
    background-color: #5A6268;
    border-color: #545B62;
    color: white;
}

.pagination {
    margin: 2rem 0 1rem;
}

.page-link {
    color: #4A90E2;
    padding: 0.5rem 1rem;
    border-radius: 0.25rem;
    margin: 0 0.25rem;
}

.page-item.active .page-link {
    background-color: #4A90E2;
    border-color: #4A90E2;
}

.page-item:first-child .page-link,
.page-item:last-child .page-link {
    margin: 0;
}

.card {
    box-shadow: 0 0.125rem 0.25rem rgba(0, 0, 0, 0.075);
    border: none;
    border-radius: 0.75rem;
    margin-bottom: 2rem;
}

.card-header {
    background-color: white;
    border-bottom: 1px solid #E9ECEF;
    padding: 1.5rem;
    border-radius: 0.75rem 0.75rem 0 0 !important;
}

.card-body {
    padding: 1.5rem;
}

.table-responsive {
    margin: -1.5rem;
    padding: 1.5rem;
    border-radius: 0.75rem;
}

.btn-primary {
    background-color: #4A90E2;
    border-color: #4A90E2;
}

.btn-primary:hover {
    background-color: #357ABD;
    border-color: #357ABD;
}

.btn-sm {
    padding: 0.4rem 0.8rem;
    font-size: 0.875rem;
}
</style>

<!--#include virtual="/contents/card_car_used/includes/footer.asp"--> 