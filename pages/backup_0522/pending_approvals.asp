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

' 전체 건수 조회 - 카드 사용 내역과 차량 이용 신청 모두 포함
countSQL = "SELECT " & _
           "(SELECT COUNT(*) FROM dbo.CardUsage cu " & _
           "JOIN dbo.ApprovalLogs al ON cu.usage_id = al.target_id AND al.target_table_name = 'CardUsage' " & _
           "WHERE al.approver_id = '" & Session("user_id") & "' " & _
           "AND al.status = '대기' " & _
           "AND (al.approval_step = 1 OR EXISTS (" & _
           "    SELECT 1 FROM dbo.ApprovalLogs prev " & _
           "    WHERE prev.target_table_name = 'CardUsage' " & _
           "    AND prev.target_id = al.target_id " & _
           "    AND prev.approval_step = al.approval_step - 1 " & _
           "    AND prev.status = '승인'" & _
           "))) + " & _
           "(SELECT COUNT(*) FROM dbo.VehicleRequests vr " & _
           "JOIN dbo.ApprovalLogs al ON vr.request_id = al.target_id AND al.target_table_name = 'VehicleRequests' " & _
           "WHERE al.approver_id = '" & Session("user_id") & "' " & _
           "AND al.status = '대기' " & _
           "AND (al.approval_step = 1 OR EXISTS (" & _
           "    SELECT 1 FROM dbo.ApprovalLogs prev " & _
           "    WHERE prev.target_table_name = 'VehicleRequests' " & _
           "    AND prev.target_id = al.target_id " & _
           "    AND prev.approval_step = al.approval_step - 1 " & _
           "    AND prev.status = '승인'" & _
           "))) AS cnt"

Set rs = db99.Execute(countSQL)
totalCount = rs("cnt")
totalPages = (totalCount + pageSize - 1) \ pageSize

Dim offsetVal
offsetVal = (currentPage - 1) * pageSize

' 카드 사용 내역과 차량 이용 신청 모두 포함하는 쿼리
pendingSQL = "SELECT * FROM (" & _
    "SELECT TOP " & totalCount & " " & _
    "'카드' AS doc_type, 'CardUsage' AS target_table_name, cu.usage_id AS doc_id, cu.usage_date AS doc_date, " & _
    "cu.store_name AS title, cu.amount, cu.purpose, " & _
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
        
    "UNION ALL " & _
    
    "SELECT TOP " & totalCount & " " & _
    "'차량' AS doc_type, 'VehicleRequests' AS target_table_name, vr.request_id AS doc_id, vr.start_date AS doc_date, " & _
    "vr.purpose AS title, " & _
    "CASE WHEN vr.distance IS NOT NULL AND fr.rate IS NOT NULL " & _
    "THEN (vr.distance * fr.rate) + ISNULL(vr.toll_fee, 0) + ISNULL(vr.parking_fee, 0) " & _
    "ELSE 0 END AS amount, " & _
    "vr.purpose, " & _
    "u.name AS requester_name, d.name AS department_name, al.status, " & _
    "al.created_at AS requested_at, al.approval_step " & _
    "FROM dbo.VehicleRequests vr " & _
    "JOIN dbo.Users u ON vr.user_id = u.user_id " & _
    "LEFT JOIN dbo.Department d ON u.department_id = d.department_id " & _
    "JOIN dbo.ApprovalLogs al ON vr.request_id = al.target_id " & _
    "LEFT JOIN (SELECT TOP 1 * FROM dbo.FuelRate ORDER BY date DESC) fr ON 1=1 " & _
    "WHERE al.target_table_name = 'VehicleRequests' " & _
    "AND al.approver_id = '" & Session("user_id") & "' " & _
    "AND al.status = '대기' " & _
    "AND (al.approval_step = 1 OR EXISTS (" & _
        "SELECT 1 FROM dbo.ApprovalLogs prev " & _
        "WHERE prev.target_table_name = 'VehicleRequests' " & _
        "AND prev.target_id = al.target_id " & _
        "AND prev.approval_step = al.approval_step - 1 " & _
        "AND prev.status = '승인')) " & _
    ") AS combined_data " & _
    "ORDER BY requested_at DESC " & _
    "OFFSET " & offsetVal & " ROWS " & _
    "FETCH NEXT " & pageSize & " ROWS ONLY"

Set rs = db99.Execute(pendingSQL)
%>

<!--#include virtual="/contents/card_car_used/includes/header.asp"-->

<style>
.container {
    max-width: 1200px;
    margin: 0 auto;
    padding: 2rem 1rem;
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

.card {
    background: #fff;
    border: none;
    border-radius: 16px;
    box-shadow: 0 0 20px rgba(0,0,0,0.05);
    margin-bottom: 2rem;
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
    padding: 1.5rem;
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

.badge {
    padding: 0.5rem 1rem;
    font-weight: 500;
    border-radius: 6px;
    font-size: 0.875rem;
}

.badge-waiting {
    background: #FFF8E6;
    color: #D4A72C;
}

.badge-approved {
    background: #E3F9E5;
    color: #1B873F;
}

.badge-rejected {
    background: #FFE9E9;
    color: #DA3633;
}

.btn {
    padding: 0.875rem 1.5rem;
    font-weight: 600;
    border-radius: 8px;
    transition: all 0.2s ease;
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

.btn-secondary {
    background: #F8FAFC;
    border: 2px solid #E9ECEF;
    color: #2C3E50;
}

.btn-secondary:hover {
    background: #E9ECEF;
    transform: translateY(-2px);
}

.btn-sm {
    padding: 0.5rem 1rem;
    font-size: 0.875rem;
}

.pagination {
    display: flex;
    justify-content: center;
    gap: 0.5rem;
    margin-top: 2rem;
}

.page-link {
    padding: 0.5rem 1rem;
    border-radius: 6px;
    color: #2C3E50;
    background: #F8FAFC;
    border: 2px solid #E9ECEF;
    font-weight: 500;
    transition: all 0.2s ease;
}

.page-link:hover {
    background: #E9ECEF;
    transform: translateY(-2px);
}

.page-item.active .page-link {
    background: #4A90E2;
    border-color: #4A90E2;
    color: white;
}

.empty-state {
    text-align: center;
    padding: 3rem 1.5rem;
}

.empty-state i {
    font-size: 3rem;
    color: #E9ECEF;
    margin-bottom: 1rem;
}

.empty-state p {
    color: #64748B;
    font-size: 1.1rem;
    margin: 0;
}
</style>

<div class="container">
    <div class="page-header">
        <h2 class="page-title">결재 대기 문서</h2>
        <div class="d-flex gap-2">
            <a href="dashboard.asp" class="btn btn-secondary">
                <i class="fas fa-home me-2"></i>대시보드
            </a>
        </div>
    </div>

    <div class="card">
        <div class="card-header">
            <h5 class="card-title">결재 대기 목록</h5>
        </div>
        <div class="card-body">
            <% If rs.EOF Then %>
                <div class="empty-state">
                    <i class="fas fa-inbox"></i>
                    <p>결재 대기 중인 문서가 없습니다.</p>
                </div>
            <% Else %>
                <div class="table-responsive">
                    <table class="table">
                        <thead>
                            <tr>
                                <th>신청일</th>
                                <th>신청자</th>
                                <th>부서</th>
                                <th>사용처</th>
                                <th class="text-end">금액</th>
                                <th>용도</th>
                                <th>상태</th>
                                <th class="text-center">처리</th>
                            </tr>
                        </thead>
                        <tbody>
                            <% 
                            ' 필드 존재 여부 확인 함수
                            Function SafeField(rs, fieldName)
                                On Error Resume Next
                                Dim value
                                value = rs(fieldName)
                                
                                If Err.Number <> 0 Or IsNull(value) Then
                                    SafeField = ""
                                Else
                                    SafeField = value
                                End If
                                On Error GoTo 0
                            End Function
                            
                            Function SafeDateTime(rs, fieldName)
                                Dim value
                                value = SafeField(rs, fieldName)
                                
                                If value = "" Then
                                    SafeDateTime = ""
                                Else
                                    SafeDateTime = FormatDateTime(value, 2)
                                End If
                            End Function
                            
                            Function SafeAmount(rs, fieldName)
                                Dim value
                                value = SafeField(rs, fieldName)
                                
                                If IsNull(value) Or Not IsNumeric(value) Then
                                    SafeAmount = "0원"
                                Else
                                    SafeAmount = FormatNumber(CDbl(value), 0) & "원"
                                End If
                            End Function

                            Do While Not rs.EOF %>
                                
                                <tr>
                                    <td><%= SafeDateTime(rs, "requested_at") %></td>
                                    <td><%= SafeField(rs, "requester_name") %></td>
                                    <td><%= SafeField(rs, "department_name") %></td>
                                    <td><%= SafeField(rs, "title") %></td>
                                    <td class="text-end"><%= SafeAmount(rs, "amount") %></td>
                                    <% 
                                    Dim purpose
                                    purpose = SafeField(rs, "purpose")
                                    
                                    If Len(purpose) > 20 Then
                                        purpose = Left(purpose, 20) & "..."
                                    End If
                                    %>
                                    <td><%= purpose %></td>
                                    <td>
                                        <span class="badge badge-waiting">
                                            <%= SafeField(rs, "status") %>
                                        </span>
                                    </td>
                                    <td class="text-center">
                                        <a href="approval_detail.asp?id=<%= SafeField(rs, "doc_id") %>" 
                                           class="btn btn-primary btn-sm">
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

                <!-- 페이지네이션 -->
                <% If totalPages > 1 Then %>
                    <nav class="pagination">
                        <% If currentPage > 1 Then %>
                            <li class="page-item">
                                <a class="page-link" href="?page=<%= currentPage - 1 %>">
                                    <i class="fas fa-chevron-left"></i>
                                </a>
                            </li>
                        <% End If %>

                        <% 
                        Dim startPage, endPage
                        startPage = currentPage - 2
                        If startPage < 1 Then startPage = 1
                        endPage = startPage + 4
                        If endPage > totalPages Then 
                            endPage = totalPages
                            startPage = endPage - 4
                            If startPage < 1 Then startPage = 1
                        End If

                        For i = startPage To endPage
                        %>
                            <li class="page-item <%= IIf(i = currentPage, "active", "") %>">
                                <a class="page-link" href="?page=<%= i %>"><%= i %></a>
                            </li>
                        <% Next %>

                        <% If currentPage < totalPages Then %>
                            <li class="page-item">
                                <a class="page-link" href="?page=<%= currentPage + 1 %>">
                                    <i class="fas fa-chevron-right"></i>
                                </a>
                            </li>
                        <% End If %>
                    </nav>
                <% End If %>
            <% End If %>
        </div>
    </div>
</div>

<!--#include virtual="/contents/card_car_used/includes/footer.asp"--> 