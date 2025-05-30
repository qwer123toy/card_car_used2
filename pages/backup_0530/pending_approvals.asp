<%@ Language="VBScript" CodePage="65001" %>
<%
Response.CodePage = 65001
Response.CharSet = "utf-8"
%>

<!--#include virtual="/db.asp"-->
<!--#include virtual="/includes/functions.asp"-->
<%
' 로그인 체크
If Not IsAuthenticated() Then
    RedirectTo("/index.asp")
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

' TOP IN 방식으로 페이징 처리 변경
' 페이지별 시작 번호 계산
Dim startRow
startRow = (currentPage - 1) * pageSize

' 카드 사용 내역과 차량 이용 신청 모두 포함하는 쿼리
pendingSQL = "SELECT * FROM (" & _
    "SELECT TOP " & pageSize & " " & _
    "'카드' AS doc_type, 'CardUsage' AS target_table_name, cu.usage_id AS doc_id, cu.usage_date AS doc_date, " & _
    "ISNULL(cu.title, cu.store_name) AS title, cu.store_name, cu.amount, cu.purpose, " & _
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
        "SELECT TOP " & startRow & " cu2.usage_id " & _
        "FROM dbo.CardUsage cu2 " & _
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
        
    "UNION ALL " & _
    
    "SELECT TOP " & pageSize & " " & _
    "'차량' AS doc_type, 'VehicleRequests' AS target_table_name, vr.request_id AS doc_id, vr.start_date AS doc_date, " & _
    "ISNULL(vr.title, vr.purpose) AS title, vr.destination AS store_name, " & _
    "(vr.distance * 2000) AS amount, " & _
    "vr.purpose, " & _
    "u.name AS requester_name, d.name AS department_name, al.status, " & _
    "al.created_at AS requested_at, al.approval_step " & _
    "FROM dbo.VehicleRequests vr " & _
    "JOIN dbo.Users u ON vr.user_id = u.user_id " & _
    "LEFT JOIN dbo.Department d ON u.department_id = d.department_id " & _
    "JOIN dbo.ApprovalLogs al ON vr.request_id = al.target_id " & _
    "WHERE al.target_table_name = 'VehicleRequests' " & _
    "AND al.approver_id = '" & Session("user_id") & "' " & _
    "AND al.status = '대기' " & _
    "AND (al.approval_step = 1 OR EXISTS (" & _
        "SELECT 1 FROM dbo.ApprovalLogs prev " & _
        "WHERE prev.target_table_name = 'VehicleRequests' " & _
        "AND prev.target_id = al.target_id " & _
        "AND prev.approval_step = al.approval_step - 1 " & _
        "AND prev.status = '승인')) " & _
    "AND vr.request_id NOT IN (" & _
        "SELECT TOP " & startRow & " vr2.request_id " & _
        "FROM dbo.VehicleRequests vr2 " & _
        "JOIN dbo.ApprovalLogs al2 ON vr2.request_id = al2.target_id " & _
        "WHERE al2.target_table_name = 'VehicleRequests' " & _
        "AND al2.approver_id = '" & Session("user_id") & "' " & _
        "AND al2.status = '대기' " & _
        "AND (al2.approval_step = 1 OR EXISTS (" & _
            "SELECT 1 FROM dbo.ApprovalLogs prev2 " & _
            "WHERE prev2.target_table_name = 'VehicleRequests' " & _
            "AND prev2.target_id = al2.target_id " & _
            "AND prev2.approval_step = al2.approval_step - 1 " & _
            "AND prev2.status = '승인')) " & _
        "ORDER BY al2.created_at DESC" & _
    ")" & _
    ") AS combined_data " & _
    "ORDER BY requested_at DESC"

Set rs = db99.Execute(pendingSQL)
%>

<!--#include virtual="/includes/header.asp"-->

<style>
.container {
    max-width: 1400px;
    margin: 0 auto;
    padding: 2rem 1rem;
}

.page-header {
    display: flex;
    justify-content: space-between;
    align-items: center;
    margin-bottom: 2rem;
    padding: 1.5rem;
    background: white;
    border-radius: 16px;
    box-shadow: 0 4px 20px rgba(0,0,0,0.08);
}

.page-title {
    font-size: 1.5rem;
    font-weight: 600;
    color: #2C3E50;
    margin: 0;
    display: flex;
    align-items: center;
}

.btn-group-nav {
    display: flex;
    gap: 0.5rem;
}

.btn-nav {
    padding: 0.875rem 1.5rem;
    font-size: 0.9rem;
    font-weight: 600;
    border-radius: 8px;
    transition: all 0.2s ease;
}

.card {
    border: none;
    box-shadow: 0 4px 20px rgba(0,0,0,0.08);
    border-radius: 16px;
    margin-bottom: 2rem;
    background: #fff;
    overflow: hidden;
}

.card-header {
    background: linear-gradient(135deg, #E8F2FF 0%, #F0F8FF 100%);
    border-bottom: 1px solid #E2E8F0;
    padding: 1.5rem;
}

.card-header h5 {
    color: #475569;
    font-weight: 600;
    margin: 0;
    font-size: 1.1rem;
}

.filter-buttons {
    display: flex;
    gap: 0.5rem;
    margin-top: 1rem;
}

.filter-btn {
    padding: 0.5rem 1rem;
    border: 1px solid #CBD5E1;
    background: #F8FAFC;
    color: #64748B;
    text-decoration: none;
    border-radius: 6px;
    font-weight: 500;
    transition: all 0.2s ease;
    font-size: 0.875rem;
}

.filter-btn:hover {
    background: #E2E8F0;
    border-color: #94A3B8;
    color: #475569;
    text-decoration: none;
    transform: translateY(-1px);
}

.filter-btn.active {
    background: #E0F2FE;
    border-color: #0EA5E9;
    color: #0369A1;
    box-shadow: 0 2px 8px rgba(14,165,233,0.15);
}

.badge {
    padding: 0.375rem 0.75rem;
    font-weight: 500;
    border-radius: 6px;
    font-size: 0.8rem;
}

.badge-success {
    background: #DCFCE7;
    color: #166534;
    border: 1px solid #BBF7D0;
}

.badge-danger {
    background: #FEE2E2;
    color: #DC2626;
    border: 1px solid #FECACA;
}

.badge-primary {
    background: #DBEAFE;
    color: #1D4ED8;
    border: 1px solid #BFDBFE;
}

.badge-info {
    background: #E0F2FE;
    color: #0369A1;
    border: 1px solid #BAE6FD;
}

.table {
    margin-bottom: 0;
}

.table th {
    background: linear-gradient(135deg, #F1F5F9 0%, #E2E8F0 100%);
    color: #475569;
    font-weight: 600;
    border: none;
    padding: 0.875rem;
    font-size: 0.9rem;
    white-space: nowrap;
}

.table td {
    padding: 1rem;
    vertical-align: middle;
    border-bottom: 1px solid #E9ECEF;
    color: #2C3E50;
}

.table tbody tr:hover {
    background-color: #F8FAFC;
    transition: background-color 0.2s ease;
}

.date-cell {
    font-size: 0.9rem;
    font-weight: 500;
    white-space: nowrap;
    min-width: 120px;
}

.amount-cell {
    font-weight: 600;
    color: #059669;
    text-align: right;
    white-space: nowrap;
}

.btn-sm {
    padding: 0.5rem 1rem;
    font-size: 0.875rem;
    border-radius: 6px;
    font-weight: 500;
}

.btn-outline-primary {
    border: 2px solid #4A90E2;
    color: #4A90E2;
    background: transparent;
}

.btn-outline-primary:hover {
    background: #4A90E2;
    color: white;
    transform: translateY(-2px);
    box-shadow: 0 4px 12px rgba(74,144,226,0.2);
}

.btn {
    padding: 0.875rem 1.5rem;
    font-weight: 600;
    border-radius: 8px;
    transition: all 0.2s ease;
}

.btn-primary {
    background: #4A90E2;
    border: none;
    color: white;
}

.btn-primary:hover {
    background: #357ABD;
    transform: translateY(-2px);
}

.btn-secondary {
    background: #6B7280;
    color: white;
    border: none;
}

.btn-secondary:hover {
    background: #4B5563;
    transform: translateY(-2px);
}

.pagination {
    margin-top: 2rem;
}

.page-link {
    border: none;
    padding: 1rem 1.25rem;
    margin: 0 0.25rem;
    border-radius: 8px;
    color: #2C3E50;
    background: #F8FAFC;
    transition: all 0.2s ease;
    font-weight: 500;
    min-height: 48px;
    display: flex;
    align-items: center;
    justify-content: center;
}

.page-link:hover {
    background: #E9ECEF;
    color: #2C3E50;
    transform: translateY(-2px);
}

.page-item.active .page-link {
    background: #4A90E2;
    color: white;
    box-shadow: 0 4px 12px rgba(74,144,226,0.2);
}

.empty-state {
    text-align: center;
    padding: 4rem 2rem;
    color: #64748B;
}

.empty-state i {
    font-size: 4rem;
    margin-bottom: 1rem;
    color: #CBD5E1;
}

.empty-state h5 {
    color: #64748B;
    margin-bottom: 0.5rem;
}

.empty-state p {
    color: #94A3B8;
}
</style>

<div class="container">
    <div class="page-header">
        <h2 class="page-title">
            <i class="fas fa-clock me-2"></i>결재 대기 문서
        </h2>
        <div class="btn-group-nav">
            <a href="dashboard.asp" class="btn btn-secondary btn-nav">
                <i class="fas fa-home me-1"></i> 대시보드
            </a>
        </div>
    </div>

    <div class="card">
        <div class="card-header">
            <h5><i class="fas fa-file-clock me-2"></i>결재 대기 목록</h5>
        </div>
        <div class="card-body">
            <% If rs.EOF Then %>
                <div class="empty-state">
                    <i class="fas fa-file-clock"></i>
                    <h5>결재 대기 중인 문서가 없습니다</h5>
                    <p>현재 처리할 결재 문서가 없습니다.</p>
                </div>
            <% Else %>
                <div class="table-responsive">
                    <table class="table table-hover">
                        <thead>
                            <tr>
                                <th style="text-align: center;">신청일</th>
                                <th style="text-align: center;">신청자</th>
                                <th style="text-align: center;">부서</th>
                                <th style="text-align: center;">종류</th>
                                <th style="text-align: center;">제목</th>
                                <th style="text-align: center;">사용처</th>
                                <th style="text-align: center;">금액</th>
                                <th style="text-align: center;">상태</th>
                                <th style="text-align: center;">처리</th>
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
                                    SafeAmount = FormatNumber(CDbl(value)) & "원"
                                End If
                            End Function

                            Do While Not rs.EOF %>
                                
                                <tr>
                                    <td style="text-align: center;" class="date-cell"><%= SafeDateTime(rs, "requested_at") %></td>
                                    <td style="text-align: center;"><%= SafeField(rs, "requester_name") %></td>
                                    <td style="text-align: center;"><%= SafeField(rs, "department_name") %></td>
                                    <td style="text-align: center;">
                                        <% 
                                        Dim docType
                                        docType = SafeField(rs, "doc_type")
                                        If docType = "카드" Then
                                        %>
                                            <span class="badge badge-primary">
                                                <i class="fas fa-credit-card me-1"></i>카드 사용
                                            </span>
                                        <% Else %>
                                            <span class="badge badge-info">
                                                <i class="fas fa-car me-1"></i>차량 이용
                                            </span>
                                        <% End If %>
                                    </td>
                                    <td style="text-align: center;"><%= SafeField(rs, "title") %></td>
                                    <td style="text-align: center;"><%= SafeField(rs, "store_name") %></td>
                                    <td style="text-align: center;" class="amount-cell"><%= SafeAmount(rs, "amount") %></td>
                                    <td style="text-align: center;">
                                        <span class="badge badge-info">
                                            <i class="fas fa-clock me-1"></i><%= SafeField(rs, "status") %>
                                        </span>
                                    </td>
                                    <td style="text-align: center;">
                                        <a href="approval_detail.asp?id=<%= SafeField(rs, "doc_id") %>&type=<%= SafeField(rs, "target_table_name") %>" 
                                           class="btn btn-sm btn-outline-primary">
                                            <i class="fas fa-check me-1"></i>결재하기
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
                                        <a class="page-link" href="?page=<%= currentPage - 1 %>">
                                            <i class="fas fa-chevron-left"></i> 이전
                                        </a>
                                    </li>
                                <% End If %>

                                <% 
                                Dim startPage, endPage
                                startPage = ((currentPage - 1) \ 5) * 5 + 1
                                endPage = startPage + 4
                                If endPage > totalPages Then endPage = totalPages

                                For i = startPage To endPage
                                %>
                                    <li class="page-item <%= IIf(i = currentPage, "active", "") %>">
                                        <a class="page-link" href="?page=<%= i %>"><%= i %></a>
                                    </li>
                                <% Next %>

                                <% If currentPage < totalPages Then %>
                                    <li class="page-item">
                                        <a class="page-link" href="?page=<%= currentPage + 1 %>">
                                            다음 <i class="fas fa-chevron-right"></i>
                                        </a>
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

<!--#include virtual="/includes/footer.asp"--> 