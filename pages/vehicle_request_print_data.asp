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
    Response.Write "로그인이 필요합니다."
    Response.End
End If

' 검색 조건 가져오기
Dim searchStartDate, searchEndDate, searchStatus, searchCondition
searchStartDate = PreventSQLInjection(Request.QueryString("start_date"))
searchEndDate = PreventSQLInjection(Request.QueryString("end_date"))
searchStatus = PreventSQLInjection(Request.QueryString("status"))

' 검색 조건 SQL 생성
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

' 유료비 단가 조회
Dim fuelRate, fuelRateSQL, fuelRateRS
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

' 전체 데이터 조회 (페이징 없이)
Dim SQL, rs
SQL = "SELECT vr.request_id, vr.start_date AS request_date, vr.purpose, vr.start_location, vr.destination, " & _
      "vr.distance, vr.approval_status, vr.title " & _
      "FROM VehicleRequests vr " & _
      searchCondition & " " & _
      "ORDER BY vr.start_date DESC, vr.request_id DESC"

Set rs = db99.Execute(SQL)

' 총 건수 계산
Dim totalCount
totalCount = 0
If Not rs.EOF Then
    rs.MoveLast
    totalCount = rs.RecordCount
    rs.MoveFirst
End If
%>

<div class="table-container">
    <div class="total-info">
        <p><strong>총 건수:</strong> <%= totalCount %>건</p>
    </div>
    
    <% If rs.EOF Then %>
        <div class="empty-message">
            <p>검색 조건에 해당하는 차량 이용 신청 내역이 없습니다.</p>
        </div>
    <% Else %>
        <table class="table">
            <thead>
                <tr>
                    <th>신청일자</th>
                    <th>제목</th>
                    <th>업무 목적</th>
                    <th>출발지</th>
                    <th>목적지</th>
                    <th>거리(km)</th>
                    <th>금액</th>
                    <th>상태</th>
                </tr>
            </thead>
            <tbody>
                <% Do While Not rs.EOF %>
                <tr>
                    <td class="date-cell"><%= FormatDate(rs("request_date")) %></td>
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
                    <td class="amount-cell"><%= FormatNumber(CDbl(rs("distance")) * CDbl(fuelRate)) %>원</td>
                    <td>
                        <span class="status-badge"><%= rs("approval_status") %></span>
                    </td>
                </tr>
                <% 
                    rs.MoveNext
                    Loop 
                %>
            </tbody>
        </table>
    <% End If %>
</div>

<%
' 사용한 객체 해제
If Not rs Is Nothing Then
    If rs.State = 1 Then
        rs.Close
    End If
    Set rs = Nothing
End If
%> 