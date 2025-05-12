<!--#include file="../includes/connection.asp"-->
<!--#include file="../includes/functions.asp"-->
<%
' 로그인 체크
If Not IsAuthenticated() Then
    RedirectTo("../index.asp")
End If

' 페이징 처리
Dim pageSize, currentPage, startRow, totalRows
pageSize = 10 ' 페이지당 표시할 레코드 수
currentPage = Request.QueryString("page")
If currentPage = "" Or Not IsNumeric(currentPage) Then
    currentPage = 1
Else
    currentPage = CInt(currentPage)
End If
startRow = (currentPage - 1) * pageSize

' 검색 조건
Dim searchStartDate, searchEndDate, searchStatus
searchStartDate = PreventSQLInjection(Request.QueryString("start_date"))
searchEndDate = PreventSQLInjection(Request.QueryString("end_date"))
searchStatus = PreventSQLInjection(Request.QueryString("status"))

' 검색 조건 SQL 생성
Dim searchCondition
searchCondition = " WHERE vr.user_id = '" & Session("user_id") & "' AND vr.is_deleted = 0 "

If searchStartDate <> "" Then
    searchCondition = searchCondition & " AND vr.request_date >= '" & searchStartDate & "'"
End If

If searchEndDate <> "" Then
    searchCondition = searchCondition & " AND vr.request_date <= '" & searchEndDate & "'"
End If

If searchStatus <> "" Then
    searchCondition = searchCondition & " AND vr.approval_status = '" & searchStatus & "'"
End If

' 총 레코드 수 조회
Dim countSQL, countRS
countSQL = "SELECT COUNT(*) AS total FROM VehicleRequests vr" & searchCondition
Set countRS = dbConn.Execute(countSQL)
totalRows = countRS("total")
countRS.Close

' 전체 페이지 수 계산
Dim totalPages
totalPages = Ceil(totalRows / pageSize)
If totalPages < 1 Then totalPages = 1

' 차량 이용 신청 내역 조회
Dim SQL, rs
SQL = "SELECT vr.request_id, vr.request_date, vr.purpose, vr.start_location, vr.destination, " & _
      "vr.distance, vr.total_amount, vr.approval_status " & _
      "FROM VehicleRequests vr " & _
      searchCondition & " " & _
      "ORDER BY vr.request_date DESC, vr.request_id DESC " & _
      "OFFSET " & startRow & " ROWS FETCH NEXT " & pageSize & " ROWS ONLY"

Set rs = dbConn.Execute(SQL)

' 페이지네이션 함수
Function Ceil(number)
    Ceil = Int(number)
    If Ceil <> number Then
        Ceil = Ceil + 1
    End If
End Function
%>
<!--#include file="../includes/header.asp"-->

<div class="vehicle-request-container">
    <div class="shadcn-card" style="margin-bottom: 20px;">
        <div class="shadcn-card-header">
            <h2 class="shadcn-card-title">개인차량 이용 신청 내역</h2>
            <p class="shadcn-card-description">개인차량 이용 신청 내역을 조회합니다.</p>
        </div>
        
        <!-- 검색 폼 -->
        <div class="shadcn-card-content">
            <form id="searchForm" method="get" action="vehicle_request.asp">
                <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 10px; margin-bottom: 15px;">
                    <div class="form-group">
                        <label class="shadcn-input-label" for="start_date">시작일</label>
                        <input class="shadcn-input" type="date" id="start_date" name="start_date" value="<%= searchStartDate %>">
                    </div>
                    
                    <div class="form-group">
                        <label class="shadcn-input-label" for="end_date">종료일</label>
                        <input class="shadcn-input" type="date" id="end_date" name="end_date" value="<%= searchEndDate %>">
                    </div>
                    
                    <div class="form-group">
                        <label class="shadcn-input-label" for="status">상태</label>
                        <select class="shadcn-select" id="status" name="status">
                            <option value="">전체</option>
                            <option value="작성중" <%= If(searchStatus = "작성중", "selected", "") %>>작성중</option>
                            <option value="대기" <%= If(searchStatus = "대기", "selected", "") %>>대기</option>
                            <option value="승인" <%= If(searchStatus = "승인", "selected", "") %>>승인</option>
                            <option value="반려" <%= If(searchStatus = "반려", "selected", "") %>>반려</option>
                        </select>
                    </div>
                </div>
                
                <div style="display: flex; justify-content: flex-end; gap: 10px;">
                    <button type="submit" class="shadcn-btn shadcn-btn-primary">검색</button>
                    <a href="vehicle_request.asp" class="shadcn-btn shadcn-btn-outline">초기화</a>
                    <a href="vehicle_request_add.asp" class="shadcn-btn shadcn-btn-secondary">새 신청서 작성</a>
                </div>
            </form>
        </div>
    </div>
    
    <!-- 차량 이용 신청 내역 목록 -->
    <div class="shadcn-card">
        <div class="shadcn-card-content">
            <% If totalRows = 0 Then %>
                <div style="text-align: center; padding: 30px;">
                    <p>등록된 개인차량 이용 신청 내역이 없습니다.</p>
                </div>
            <% Else %>
                <table class="shadcn-table">
                    <thead class="shadcn-table-header">
                        <tr>
                            <th>신청일자</th>
                            <th>업무 목적</th>
                            <th>출발지</th>
                            <th>목적지</th>
                            <th>거리(km)</th>
                            <th>금액</th>
                            <th>상태</th>
                            <th>관리</th>
                        </tr>
                    </thead>
                    <tbody>
                        <% Do While Not rs.EOF %>
                        <tr>
                            <td><%= FormatDate(rs("request_date")) %></td>
                            <td><%= rs("purpose") %></td>
                            <td><%= rs("start_location") %></td>
                            <td><%= rs("destination") %></td>
                            <td><%= rs("distance") %></td>
                            <td><%= FormatNumber(rs("total_amount")) %></td>
                            <td>
                                <% 
                                Dim statusClass
                                Select Case rs("approval_status")
                                    Case "승인"
                                        statusClass = "shadcn-badge shadcn-badge-primary"
                                    Case "반려"
                                        statusClass = "shadcn-badge shadcn-badge-destructive"
                                    Case "작성중"
                                        statusClass = "shadcn-badge shadcn-badge-secondary"
                                    Case Else
                                        statusClass = "shadcn-badge shadcn-badge-outline"
                                End Select
                                %>
                                <span class="<%= statusClass %>"><%= rs("approval_status") %></span>
                            </td>
                            <td>
                                <div style="display: flex; gap: 5px;">
                                    <a href="vehicle_request_view.asp?id=<%= rs("request_id") %>" class="shadcn-btn shadcn-btn-outline" style="padding: 2px 8px; font-size: 0.75rem;">상세</a>
                                    <% If rs("approval_status") = "작성중" Then %>
                                    <a href="vehicle_request_edit.asp?id=<%= rs("request_id") %>" class="shadcn-btn shadcn-btn-secondary" style="padding: 2px 8px; font-size: 0.75rem;">수정</a>
                                    <% End If %>
                                    <% If rs("approval_status") = "승인" Then %>
                                    <a href="vehicle_request_print.asp?id=<%= rs("request_id") %>" class="shadcn-btn shadcn-btn-primary" style="padding: 2px 8px; font-size: 0.75rem;">출력</a>
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
            <% End If %>
            
            <!-- 페이징 -->
            <% If totalPages > 1 Then %>
            <div style="display: flex; justify-content: center; margin-top: 20px;">
                <div class="pagination">
                    <% 
                    Dim pageStart, pageEnd, i
                    
                    ' 표시할 페이지 범위 설정
                    pageStart = currentPage - 5
                    If pageStart < 1 Then pageStart = 1
                    
                    pageEnd = pageStart + 9
                    If pageEnd > totalPages Then pageEnd = totalPages
                    
                    ' 페이지 링크 생성
                    If currentPage > 1 Then
                    %>
                        <a href="vehicle_request.asp?page=1&start_date=<%= searchStartDate %>&end_date=<%= searchEndDate %>&status=<%= searchStatus %>" class="shadcn-btn shadcn-btn-outline" style="padding: 5px 10px; margin: 0 2px;">처음</a>
                        <a href="vehicle_request.asp?page=<%= currentPage - 1 %>&start_date=<%= searchStartDate %>&end_date=<%= searchEndDate %>&status=<%= searchStatus %>" class="shadcn-btn shadcn-btn-outline" style="padding: 5px 10px; margin: 0 2px;">이전</a>
                    <% End If %>
                    
                    <% For i = pageStart To pageEnd %>
                        <% If i = CInt(currentPage) Then %>
                            <span class="shadcn-btn shadcn-btn-primary" style="padding: 5px 10px; margin: 0 2px;"><%= i %></span>
                        <% Else %>
                            <a href="vehicle_request.asp?page=<%= i %>&start_date=<%= searchStartDate %>&end_date=<%= searchEndDate %>&status=<%= searchStatus %>" class="shadcn-btn shadcn-btn-outline" style="padding: 5px 10px; margin: 0 2px;"><%= i %></a>
                        <% End If %>
                    <% Next %>
                    
                    <% If CInt(currentPage) < totalPages Then %>
                        <a href="vehicle_request.asp?page=<%= currentPage + 1 %>&start_date=<%= searchStartDate %>&end_date=<%= searchEndDate %>&status=<%= searchStatus %>" class="shadcn-btn shadcn-btn-outline" style="padding: 5px 10px; margin: 0 2px;">다음</a>
                        <a href="vehicle_request.asp?page=<%= totalPages %>&start_date=<%= searchStartDate %>&end_date=<%= searchEndDate %>&status=<%= searchStatus %>" class="shadcn-btn shadcn-btn-outline" style="padding: 5px 10px; margin: 0 2px;">마지막</a>
                    <% End If %>
                </div>
            </div>
            <% End If %>
        </div>
    </div>
</div>

<!--#include file="../includes/footer.asp"--> 