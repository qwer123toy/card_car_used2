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

' 세션에서 메시지 가져오기
If Session("success_msg") <> "" Then
    successMsg = Session("success_msg")
    Session("success_msg") = ""
End If

If Session("error_msg") <> "" Then
    errorMsg = Session("error_msg")
    Session("error_msg") = ""
End If

On Error Resume Next

' 페이징 처리
Dim pageSize, currentPage, startRow, totalRows, fuelRate
pageSize = 10 ' 페이지당 표시할 레코드 수
currentPage = Request.QueryString("page")
If currentPage = "" Or Not IsNumeric(currentPage) Then
    currentPage = 1
Else
    currentPage = CInt(currentPage)
End If
startRow = (currentPage - 1) * pageSize

' 검색 조건
Dim searchStartDate, searchEndDate, searchStatus, errorMsg, successMsg
searchStartDate = PreventSQLInjection(Request.QueryString("start_date"))
searchEndDate = PreventSQLInjection(Request.QueryString("end_date"))
searchStatus = PreventSQLInjection(Request.QueryString("status"))

' 검색 조건 SQL 생성
Dim searchCondition
searchCondition = " WHERE vr.user_id = '" & Session("user_id") & "' AND vr.is_deleted = 0 "

If searchStartDate <> "" Then
    searchCondition = searchCondition & " AND vr.start_date >= '" & searchStartDate & "'"
End If

If searchEndDate <> "" Then
    searchCondition = searchCondition & " AND vr.start_date <= '" & searchEndDate & "'"
End If

If searchStatus <> "" Then
    searchCondition = searchCondition & " AND vr.approval_status = '" & searchStatus & "'"
End If

' 총 레코드 수 조회
Dim countSQL, countRS
countSQL = "SELECT COUNT(*) AS total FROM VehicleRequests vr" & searchCondition
Set countRS = db99.Execute(countSQL)

If Err.Number <> 0 Then
    totalRows = 0
    Err.Clear
Else
totalRows = countRS("total")
countRS.Close
End If

' 전체 페이지 수 계산
Dim totalPages
totalPages = Ceil(totalRows / pageSize)
If totalPages < 1 Then totalPages = 1

' 유류비 단가 조회
Dim fuelRateSQL, fuelRateRS
fuelRateSQL = "SELECT TOP 1 rate FROM FuelRate ORDER BY date DESC"
Set fuelRateRS = db.Execute(fuelRateSQL)

If Err.Number <> 0 Or fuelRateRS.EOF Then
    fuelRate = 2000 ' 기본값 설정
Else
    fuelRate = fuelRateRS("rate")
End If

If Not fuelRateRS Is Nothing Then
    If fuelRateRS.State = 1 Then
        fuelRateRS.Close
    End If
    Set fuelRateRS = Nothing
End If

' 차량 이용 신청 내역 조회 - 구버전 SQL Server용 페이징 처리
Dim SQL, rs

' 기본 쿼리
SQL = "SELECT TOP " & pageSize & " vr.request_id, vr.start_date AS request_date, vr.purpose, vr.start_location, vr.destination, " & _
      "vr.distance, vr.approval_status, vr.title " & _
      "FROM VehicleRequests vr " & _
      searchCondition & " "

' 1페이지가 아닌 경우 ID 기준으로 건너뛰기
If currentPage > 1 Then
    SQL = "SELECT TOP " & pageSize & " vr.request_id, vr.start_date AS request_date, vr.purpose, vr.start_location, vr.destination, " & _
          "vr.distance, vr.approval_status, vr.title " & _
      "FROM VehicleRequests vr " & _
      searchCondition & " " & _
          "AND vr.request_id NOT IN (SELECT TOP " & startRow & " request_id FROM VehicleRequests vr " & _
          searchCondition & " ORDER BY vr.start_date DESC, vr.request_id DESC) " & _
          "ORDER BY vr.start_date DESC, vr.request_id DESC"
Else
    SQL = SQL & "ORDER BY vr.start_date DESC, vr.request_id DESC"
End If

Set rs = db99.Execute(SQL)

' 오류 발생 시 빈 레코드셋 생성
If Err.Number <> 0 Then
    Set rs = Server.CreateObject("ADODB.Recordset")
    rs.Fields.Append "request_id", 3 ' adInteger
    rs.Fields.Append "request_date", 7 ' adDate
    rs.Fields.Append "purpose", 200, 100 ' adVarChar
    rs.Fields.Append "start_location", 200, 100 ' adVarChar
    rs.Fields.Append "destination", 200, 100 ' adVarChar
    rs.Fields.Append "distance", 5 ' adDouble
    rs.Fields.Append "approval_status", 200, 20 ' adVarChar
    rs.Open
    Err.Clear
End If

' 페이지네이션 함수
Function Ceil(number)
    Ceil = Int(number)
    If Ceil <> number Then
        Ceil = Ceil + 1
    End If
End Function

On Error GoTo 0
%>
<!--#include file="../includes/header.asp"-->

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

.badge-secondary {
    background: #F1F5F9;
    color: #475569;
    border: 1px solid #E2E8F0;
}

.badge-outline {
    background: transparent;
    border: 1px solid #E5E7EB;
    color: #6B7280;
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

.btn-secondary {
    background: #6B7280;
    color: white;
    border: none;
}

.btn-secondary:hover {
    background: #4B5563;
    transform: translateY(-2px);
}

.btn-primary {
    background: #4A90E2;
    color: white;
    border: none;
}

.btn-primary:hover {
    background: #357ABD;
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

.form-group {
    margin-bottom: 1rem;
}

.form-group label {
    display: block;
    margin-bottom: 0.5rem;
    font-weight: 600;
    color: #2C3E50;
}

.form-group input,
.form-group select {
    width: 100%;
    padding: 0.75rem;
    border: 2px solid #E9ECEF;
    border-radius: 8px;
    font-size: 1rem;
    transition: border-color 0.2s ease;
}

.form-group input:focus,
.form-group select:focus {
    outline: none;
    border-color: #4A90E2;
    box-shadow: 0 0 0 3px rgba(74,144,226,0.1);
}

.btn {
    padding: 0.75rem 1.5rem;
    border-radius: 8px;
    font-weight: 600;
    text-decoration: none;
    display: inline-block;
    transition: all 0.2s ease;
    border: none;
    cursor: pointer;
}

.btn-outline {
    background: transparent;
    border: 2px solid #E9ECEF;
    color: #6B7280;
}

.btn-outline:hover {
    background: #F3F4F6;
    color: #374151;
}

.alert {
    padding: 1rem;
    border-radius: 8px;
    margin-bottom: 1rem;
}

.alert-error {
    background: #FEF2F2;
    border: 1px solid #FECACA;
    color: #B91C1C;
}

.alert-success {
    background: #F0FDF4;
    border: 1px solid #BBF7D0;
    color: #166534;
}
</style>

<div class="container">
    <div class="page-header">
        <h2 class="page-title">
            <i class="fas fa-car me-2"></i>개인차량 이용 신청 내역
        </h2>
        <div class="btn-group-nav">
            <a href="dashboard.asp" class="btn btn-secondary btn-nav">
                <i class="fas fa-home me-1"></i> 대시보드
            </a>
            <a href="vehicle_request_add.asp" class="btn btn-primary btn-nav">
                <i class="fas fa-plus me-1"></i> 새 신청서 작성
            </a>
        </div>
    </div>

    <div class="card">
        <div class="card-header">
            <h5><i class="fas fa-search me-2"></i>검색 조건</h5>
        </div>
        <div class="card-body">
        
            <% If errorMsg <> "" Then %>
            <div class="alert alert-error">
                <div>
                    <span><strong>오류:</strong> <%= errorMsg %></span>
                </div>
            </div>
            <% End If %>
            
            <% If successMsg <> "" Then %>
            <div class="alert alert-success">
                <div>
                    <span><strong>성공:</strong> <%= successMsg %></span>
                </div>
            </div>
            <% End If %>
            
            <!-- 검색 폼 -->
            <form id="searchForm" method="get" action="vehicle_request.asp">
                <div class="row justify-content-center">
                    <div class="col-lg-10">
                        <div class="row g-3 mb-3">
                            <div class="col-md-3">
                                <div class="form-group">
                                    <label for="start_date">시작일</label>
                                    <input type="date" id="start_date" name="start_date" value="<%= searchStartDate %>">
                                </div>
                            </div>
                            
                            <div class="col-md-3">
                                <div class="form-group">
                                    <label for="end_date">종료일</label>
                                    <input type="date" id="end_date" name="end_date" value="<%= searchEndDate %>">
                                </div>
                            </div>
                            
                            <div class="col-md-3">
                                <div class="form-group">
                                    <label for="status">상태</label>
                                    <select id="status" name="status">
                                        <option value="">전체</option>
                                        <option value="작성중" <% If searchStatus = "작성중" Then Response.Write("selected") End If %>>작성중</option>
                                        <option value="대기" <% If searchStatus = "대기" Then Response.Write("selected") End If %>>대기</option>
                                        <option value="승인" <% If searchStatus = "승인" Then Response.Write("selected") End If %>>승인</option>
                                        <option value="반려" <% If searchStatus = "반려" Then Response.Write("selected") End If %>>반려</option>
                                    </select>
                                </div>
                            </div>
                            
                            <div class="col-md-3 d-flex align-items-end">
                                <div class="form-group w-100">
                                    <button type="submit" class="btn btn-primary w-100 mb-2">
                                        <i class="fas fa-search me-1"></i>검색
                                    </button>
                                    <a href="vehicle_request.asp" class="btn btn-outline w-100">
                                        <i class="fas fa-refresh me-1"></i>초기화
                                    </a>
                                </div>
                            </div>
                        </div>
                    </div>
                </div>
            </form>
        </div>
    </div>
    
    <!-- 차량 이용 신청 내역 목록 -->
    <div class="card">
        <div class="card-header">
            <h5><i class="fas fa-list me-2"></i>차량 이용 신청 내역</h5>
        </div>
        <div class="card-body">
            <% If totalRows = 0 Then %>
                <div class="empty-state">
                    <i class="fas fa-car"></i>
                    <h5>등록된 개인차량 이용 신청 내역이 없습니다</h5>
                    <p>새로운 차량 이용 신청서를 작성해보세요.</p>
                </div>
            <% Else %>
                <div class="table-responsive">
                    <table class="table table-hover">
                        <thead>
                            <tr>
                                <th style="text-align: center;">신청일자</th>
                                <th style="text-align: center;">제목</th>
                                <th style="text-align: center;">업무 목적</th>
                                <th style="text-align: center;">출발지</th>
                                <th style="text-align: center;">목적지</th>
                                <th style="text-align: center;">거리(km)</th>
                                <th style="text-align: center;">금액</th>
                                <th style="text-align: center;">상태</th>
                                <th style="text-align: center;">관리</th>
                            </tr>
                        </thead>
                        <tbody>
                            <% Do While Not rs.EOF %>
                            <tr>
                                <td style="text-align: center;" class="date-cell"><%= FormatDate(rs("request_date")) %></td>
                                <td style="text-align: center;"><% 
                                    If IsNull(rs("title")) Or rs("title") = "" Then
                                        Response.Write(rs("purpose"))
                                    Else
                                        Response.Write(rs("title"))
                                    End If
                                %></td>
                                <td style="text-align: center;"><%= rs("purpose") %></td>
                                <td style="text-align: center;"><%= rs("start_location") %></td>
                                <td style="text-align: center;"><%= rs("destination") %></td>
                                <td style="text-align: center;"><%= rs("distance") %></td>
                                <td style="text-align: center;" class="amount-cell"><%= FormatNumber(CDbl(rs("distance")) * CDbl(fuelRate)) %>원</td>
                                <td style="text-align: center;">
                                    <% 
                                    Dim statusClass
                                    Select Case rs("approval_status")
                                        Case "승인"
                                            statusClass = "badge badge-success"
                                        Case "반려"
                                            statusClass = "badge badge-danger"
                                        Case "작성중"
                                            statusClass = "badge badge-secondary"
                                        Case "대기"
                                            statusClass = "badge badge-info"
                                        Case Else
                                            statusClass = "badge badge-outline"
                                    End Select
                                    %>
                                    <span class="<%= statusClass %>">
                                        <% If rs("approval_status") = "승인" Then %>
                                            <i class="fas fa-check me-1"></i>
                                        <% ElseIf rs("approval_status") = "반려" Then %>
                                            <i class="fas fa-times me-1"></i>
                                        <% ElseIf rs("approval_status") = "대기" Then %>
                                            <i class="fas fa-clock me-1"></i>
                                        <% Else %>
                                            <i class="fas fa-edit me-1"></i>
                                        <% End If %>
                                        <%= rs("approval_status") %>
                                    </span>
                                </td>
                                <td style="text-align: center;">
                                    <div style="display: flex; gap: 5px; justify-content: center;">
                                        <a href="vehicle_request_view.asp?id=<%= rs("request_id") %>" class="btn btn-sm btn-outline-primary">
                                            <i class="fas fa-eye me-1"></i>상세
                                        </a>
                                        <% If rs("approval_status") <> "완료" And rs("approval_status") <> "승인" Then %>
                                        <a href="vehicle_request_edit.asp?id=<%= rs("request_id") %>" class="btn btn-sm btn-secondary">
                                            <i class="fas fa-edit me-1"></i>수정
                                        </a>
                                        <% End If %>
                                        <% If rs("approval_status") = "승인" Then %>
                                        <a href="vehicle_request_print.asp?id=<%= rs("request_id") %>" class="btn btn-sm btn-primary">
                                            <i class="fas fa-print me-1"></i>출력
                                        </a>
                                        <% End If %>
                                    </div>
                                </td>
                            </tr>
                            <% 
                                rs.MoveNext
                                Loop 
                            %>
                        </tbody>
                    </table>
                </div>
            <% End If %>
            
            <!-- 페이징 -->
            <% If totalPages > 1 Then %>
                <div class="d-flex justify-content-center mt-4">
                    <nav aria-label="Page navigation">
                        <ul class="pagination">
                            <% If currentPage > 1 Then %>
                                <li class="page-item">
                                    <a class="page-link" href="vehicle_request.asp?page=<%= currentPage - 1 %>&start_date=<%= searchStartDate %>&end_date=<%= searchEndDate %>&status=<%= searchStatus %>">
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
                                    <a class="page-link" href="vehicle_request.asp?page=<%= i %>&start_date=<%= searchStartDate %>&end_date=<%= searchEndDate %>&status=<%= searchStatus %>"><%= i %></a>
                                </li>
                            <% Next %>

                            <% If currentPage < totalPages Then %>
                                <li class="page-item">
                                    <a class="page-link" href="vehicle_request.asp?page=<%= currentPage + 1 %>&start_date=<%= searchStartDate %>&end_date=<%= searchEndDate %>&status=<%= searchStatus %>">
                                        다음 <i class="fas fa-chevron-right"></i>
                                    </a>
                                </li>
                            <% End If %>
                        </ul>
                    </nav>
                </div>
            <% End If %>
        </div>
    </div>
</div>

<!--#include file="../includes/footer.asp"--> 