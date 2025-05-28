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

<div class="vehicle-request-container">
    <div class="shadcn-card" style="margin-bottom: 20px;">
        <div class="shadcn-card-header">
            <h2 class="shadcn-card-title">개인차량 이용 신청 내역</h2>
            <p class="shadcn-card-description">개인차량 이용 신청 내역을 조회합니다.</p>
        </div>
        
        <% If errorMsg <> "" Then %>
        <div class="shadcn-alert shadcn-alert-error">
            <div>
                <span class="shadcn-alert-title">오류</span>
                <span class="shadcn-alert-description"><%= errorMsg %></span>
            </div>
        </div>
        <% End If %>
        
        <% If successMsg <> "" Then %>
        <div class="shadcn-alert shadcn-alert-success">
            <div>
                <span class="shadcn-alert-title">성공</span>
                <span class="shadcn-alert-description"><%= successMsg %></span>
            </div>
        </div>
        <% End If %>
        
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
                            <option value="작성중" <% If searchStatus = "작성중" Then Response.Write("selected") End If %>>작성중</option>
                            <option value="대기" <% If searchStatus = "대기" Then Response.Write("selected") End If %>>대기</option>
                            <option value="승인" <% If searchStatus = "승인" Then Response.Write("selected") End If %>>승인</option>
                            <option value="반려" <% If searchStatus = "반려" Then Response.Write("selected") End If %>>반려</option>
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
                <table class="shadcn-table">
                    <thead class="shadcn-table-header">
                        <tr>
                            <th>신청일자</th>
                            <th>제목</th>
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
                        <tr>
                            <td colspan="9" class="text-center">등록된 개인차량 이용 신청 내역이 없습니다.</td>
                        </tr>
                    </tbody>
                </table>
            <% Else %>
                <table class="shadcn-table">
                    <thead class="shadcn-table-header">
                        <tr>
                            <th>신청일자</th>
                            <th>제목</th>
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
                            <td><% 
                                If IsNull(rs("title")) Or rs("title") = "" Then
                                    Response.Write(rs("purpose"))
                                Else
                                    Response.Write(rs("title"))
                                End If
                            %></td>
                            <td><%= rs("purpose") %></td>
                            <td><%= rs("start_location") %></td>
                            <td><%= rs("destination") %></td>
                            <td><%= rs("distance") %></td>
                            <td><%= FormatNumber(CDbl(rs("distance")) * CDbl(fuelRate)) %></td>
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
                                    <% If rs("approval_status") <> "완료" And rs("approval_status") <> "승인" Then %>
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