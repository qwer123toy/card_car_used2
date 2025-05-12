<!--#include file="../includes/connection.asp"-->
<!--#include file="../includes/functions.asp"-->
<%
' 로그인 체크
If Not IsAuthenticated() Then
    RedirectTo("../index.asp")
End If

' 사용자 정보 조회
Dim SQL, rs, userName, userDeptId, userDeptName
SQL = "SELECT u.name, u.department_id, d.name AS department_name FROM Users u " & _
      "LEFT JOIN Department d ON u.department_id = d.department_id " & _
      "WHERE u.user_id = '" & Session("user_id") & "'"
Set rs = dbConn.Execute(SQL)

If Not rs.EOF Then
    userName = rs("name")
    userDeptId = rs("department_id")
    userDeptName = rs("department_name")
End If
rs.Close

' 최근 카드 사용 내역
Dim recentCardSQL, recentCardRS
recentCardSQL = "SELECT TOP 5 c.usage_id, c.usage_date, c.amount, ca.account_name " & _
                "FROM CardUsage c " & _
                "JOIN CardAccount ca ON c.card_id = ca.card_id " & _
                "WHERE c.user_id = '" & Session("user_id") & "' " & _
                "ORDER BY c.usage_date DESC"
Set recentCardRS = dbConn.Execute(recentCardSQL)

' 최근 차량 사용 내역
Dim recentVehicleSQL, recentVehicleRS
recentVehicleSQL = "SELECT TOP 5 v.request_id, v.request_date, v.purpose, v.approval_status " & _
                   "FROM VehicleRequests v " & _
                   "WHERE v.user_id = '" & Session("user_id") & "' AND v.is_deleted = 0 " & _
                   "ORDER BY v.request_date DESC"
Set recentVehicleRS = dbConn.Execute(recentVehicleSQL)
%>
<!--#include file="../includes/header.asp"-->

<div class="dashboard-container">
    <div class="dashboard-welcome shadcn-card" style="margin-bottom: 30px;">
        <div class="shadcn-card-header">
            <h2 class="shadcn-card-title">환영합니다, <%= userName %> 님</h2>
            <p class="shadcn-card-description"><%= userDeptName %> 소속</p>
        </div>
        <div class="shadcn-card-content">
            <p>카드 지출 결의 및 개인차량 이용 내력 관리 시스템에 오신 것을 환영합니다.</p>
            <p>아래에서 최근 활동 내역을 확인하거나, 새로운 요청을 생성할 수 있습니다.</p>
        </div>
    </div>
    
    <div class="dashboard-grid" style="display: grid; grid-template-columns: repeat(auto-fit, minmax(450px, 1fr)); gap: 20px;">
        <!-- 카드 사용 내역 -->
        <div class="shadcn-card">
            <div class="shadcn-card-header">
                <h3 class="shadcn-card-title">최근 카드 사용 내역</h3>
            </div>
            <div class="shadcn-card-content">
                <% If recentCardRS.EOF Then %>
                <p>최근 카드 사용 내역이 없습니다.</p>
                <% Else %>
                <table class="shadcn-table">
                    <thead class="shadcn-table-header">
                        <tr>
                            <th>날짜</th>
                            <th>카드</th>
                            <th>금액</th>
                        </tr>
                    </thead>
                    <tbody>
                        <% Do While Not recentCardRS.EOF %>
                        <tr>
                            <td><%= FormatDate(recentCardRS("usage_date")) %></td>
                            <td><%= recentCardRS("account_name") %></td>
                            <td><%= FormatNumber(recentCardRS("amount")) %></td>
                        </tr>
                        <% 
                            recentCardRS.MoveNext
                            Loop 
                        %>
                    </tbody>
                </table>
                <% End If %>
            </div>
            <div class="shadcn-card-footer">
                <a href="card_usage.asp" class="shadcn-btn shadcn-btn-outline">모든 내역 보기</a>
                <a href="card_usage_add.asp" class="shadcn-btn shadcn-btn-primary">새 내역 등록</a>
            </div>
        </div>
        
        <!-- 차량 사용 내역 -->
        <div class="shadcn-card">
            <div class="shadcn-card-header">
                <h3 class="shadcn-card-title">최근 차량 사용 신청</h3>
            </div>
            <div class="shadcn-card-content">
                <% If recentVehicleRS.EOF Then %>
                <p>최근 차량 사용 신청 내역이 없습니다.</p>
                <% Else %>
                <table class="shadcn-table">
                    <thead class="shadcn-table-header">
                        <tr>
                            <th>날짜</th>
                            <th>용도</th>
                            <th>상태</th>
                        </tr>
                    </thead>
                    <tbody>
                        <% Do While Not recentVehicleRS.EOF %>
                        <tr>
                            <td><%= FormatDate(recentVehicleRS("request_date")) %></td>
                            <td><%= recentVehicleRS("purpose") %></td>
                            <td>
                                <% 
                                Dim statusClass
                                Select Case recentVehicleRS("approval_status")
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
                                <span class="<%= statusClass %>"><%= recentVehicleRS("approval_status") %></span>
                            </td>
                        </tr>
                        <% 
                            recentVehicleRS.MoveNext
                            Loop 
                        %>
                    </tbody>
                </table>
                <% End If %>
            </div>
            <div class="shadcn-card-footer">
                <a href="vehicle_request.asp" class="shadcn-btn shadcn-btn-outline">모든 내역 보기</a>
                <a href="vehicle_request_add.asp" class="shadcn-btn shadcn-btn-primary">새 신청서 작성</a>
            </div>
        </div>
    </div>
</div>

<!--#include file="../includes/footer.asp"--> 