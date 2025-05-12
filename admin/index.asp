<!--#include file="../includes/connection.asp"-->
<!--#include file="../includes/functions.asp"-->
<%
' 로그인 및 관리자 권한 체크
If Not IsAuthenticated() Then
    RedirectTo("../index.asp")
End If

If Not IsAdmin() Then
    RedirectTo("../pages/dashboard.asp")
End If

' 통계 데이터 조회
Dim userCountSQL, userCountRS, userCount
userCountSQL = "SELECT COUNT(*) AS total FROM Users"
Set userCountRS = dbConn.Execute(userCountSQL)
userCount = userCountRS("total")
userCountRS.Close

Dim cardUsageCountSQL, cardUsageCountRS, cardUsageCount
cardUsageCountSQL = "SELECT COUNT(*) AS total FROM CardUsage"
Set cardUsageCountRS = dbConn.Execute(cardUsageCountSQL)
cardUsageCount = cardUsageCountRS("total")
cardUsageCountRS.Close

Dim vehicleRequestCountSQL, vehicleRequestCountRS, vehicleRequestCount
vehicleRequestCountSQL = "SELECT COUNT(*) AS total FROM VehicleRequests WHERE is_deleted = 0"
Set vehicleRequestCountRS = dbConn.Execute(vehicleRequestCountSQL)
vehicleRequestCount = vehicleRequestCountRS("total")
vehicleRequestCountRS.Close

Dim cardUsageTotalSQL, cardUsageTotalRS, cardUsageTotal
cardUsageTotalSQL = "SELECT SUM(amount) AS total FROM CardUsage"
Set cardUsageTotalRS = dbConn.Execute(cardUsageTotalSQL)
If Not IsNull(cardUsageTotalRS("total")) Then
    cardUsageTotal = cardUsageTotalRS("total")
Else
    cardUsageTotal = 0
End If
cardUsageTotalRS.Close

Dim vehicleRequestTotalSQL, vehicleRequestTotalRS, vehicleRequestTotal
vehicleRequestTotalSQL = "SELECT SUM(total_amount) AS total FROM VehicleRequests WHERE is_deleted = 0"
Set vehicleRequestTotalRS = dbConn.Execute(vehicleRequestTotalSQL)
If Not IsNull(vehicleRequestTotalRS("total")) Then
    vehicleRequestTotal = vehicleRequestTotalRS("total")
Else
    vehicleRequestTotal = 0
End If
vehicleRequestTotalRS.Close

' 최근 활동 로그
Dim recentLogsSQL, recentLogsRS
recentLogsSQL = "SELECT TOP 10 al.activity_id, al.user_id, al.action, al.description, al.created_at, u.name " & _
               "FROM ActivityLogs al " & _
               "JOIN Users u ON al.user_id = u.user_id " & _
               "ORDER BY al.created_at DESC"
Set recentLogsRS = dbConn.Execute(recentLogsSQL)

' 승인 대기 중인 차량 신청
Dim pendingVehicleRequestsSQL, pendingVehicleRequestsRS
pendingVehicleRequestsSQL = "SELECT TOP 5 vr.request_id, vr.request_date, vr.purpose, u.name, vr.total_amount " & _
                          "FROM VehicleRequests vr " & _
                          "JOIN Users u ON vr.user_id = u.user_id " & _
                          "WHERE vr.approval_status = '대기' AND vr.is_deleted = 0 " & _
                          "ORDER BY vr.request_date ASC"
Set pendingVehicleRequestsRS = dbConn.Execute(pendingVehicleRequestsSQL)
%>
<!--#include file="../includes/header.asp"-->

<div class="admin-dashboard-container">
    <div class="shadcn-card" style="margin-bottom: 20px;">
        <div class="shadcn-card-header">
            <h2 class="shadcn-card-title">관리자 대시보드</h2>
            <p class="shadcn-card-description">시스템 현황 및 관리 기능에 접근할 수 있습니다.</p>
        </div>
    </div>
    
    <!-- 통계 요약 -->
    <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(250px, 1fr)); gap: 20px; margin-bottom: 20px;">
        <div class="shadcn-card">
            <div class="shadcn-card-header">
                <h3 class="shadcn-card-title">사용자</h3>
            </div>
            <div class="shadcn-card-content" style="text-align: center;">
                <p style="font-size: 2rem; font-weight: bold;"><%= userCount %></p>
                <p>총 사용자</p>
            </div>
            <div class="shadcn-card-footer">
                <a href="users.asp" class="shadcn-btn shadcn-btn-outline" style="width: 100%;">사용자 관리</a>
            </div>
        </div>
        
        <div class="shadcn-card">
            <div class="shadcn-card-header">
                <h3 class="shadcn-card-title">카드 사용</h3>
            </div>
            <div class="shadcn-card-content" style="text-align: center;">
                <p style="font-size: 2rem; font-weight: bold;"><%= cardUsageCount %></p>
                <p>총 카드 사용 내역</p>
                <p style="margin-top: 10px; font-size: 1.2rem;"><%= FormatNumber(cardUsageTotal) %></p>
                <p>총 사용 금액</p>
            </div>
            <div class="shadcn-card-footer">
                <a href="card_usage_manage.asp" class="shadcn-btn shadcn-btn-outline" style="width: 100%;">카드 내역 관리</a>
            </div>
        </div>
        
        <div class="shadcn-card">
            <div class="shadcn-card-header">
                <h3 class="shadcn-card-title">차량 이용</h3>
            </div>
            <div class="shadcn-card-content" style="text-align: center;">
                <p style="font-size: 2rem; font-weight: bold;"><%= vehicleRequestCount %></p>
                <p>총 차량 이용 신청</p>
                <p style="margin-top: 10px; font-size: 1.2rem;"><%= FormatNumber(vehicleRequestTotal) %></p>
                <p>총 비용</p>
            </div>
            <div class="shadcn-card-footer">
                <a href="vehicle_request_manage.asp" class="shadcn-btn shadcn-btn-outline" style="width: 100%;">차량 이용 관리</a>
            </div>
        </div>
        
        <div class="shadcn-card">
            <div class="shadcn-card-header">
                <h3 class="shadcn-card-title">관리 메뉴</h3>
            </div>
            <div class="shadcn-card-content">
                <ul style="list-style: none; padding: 0;">
                    <li style="margin-bottom: 10px;"><a href="card_manage.asp" class="shadcn-btn shadcn-btn-outline" style="width: 100%;">카드 관리</a></li>
                    <li style="margin-bottom: 10px;"><a href="account_types.asp" class="shadcn-btn shadcn-btn-outline" style="width: 100%;">계정 과목 관리</a></li>
                    <li><a href="fuel_rate.asp" class="shadcn-btn shadcn-btn-outline" style="width: 100%;">유류비 단가 관리</a></li>
                </ul>
            </div>
        </div>
    </div>
    
    <!-- 승인 대기 중인 차량 신청 -->
    <div class="shadcn-card" style="margin-bottom: 20px;">
        <div class="shadcn-card-header">
            <h3 class="shadcn-card-title">승인 대기 중인 차량 이용 신청</h3>
        </div>
        <div class="shadcn-card-content">
            <% If pendingVehicleRequestsRS.EOF Then %>
                <div style="text-align: center; padding: 20px;">
                    <p>승인 대기 중인 차량 이용 신청이 없습니다.</p>
                </div>
            <% Else %>
                <table class="shadcn-table">
                    <thead class="shadcn-table-header">
                        <tr>
                            <th>신청 ID</th>
                            <th>신청일자</th>
                            <th>신청자</th>
                            <th>용도</th>
                            <th>금액</th>
                            <th>관리</th>
                        </tr>
                    </thead>
                    <tbody>
                        <% Do While Not pendingVehicleRequestsRS.EOF %>
                        <tr>
                            <td><%= pendingVehicleRequestsRS("request_id") %></td>
                            <td><%= FormatDate(pendingVehicleRequestsRS("request_date")) %></td>
                            <td><%= pendingVehicleRequestsRS("name") %></td>
                            <td><%= pendingVehicleRequestsRS("purpose") %></td>
                            <td><%= FormatNumber(pendingVehicleRequestsRS("total_amount")) %></td>
                            <td>
                                <a href="vehicle_request_approve.asp?id=<%= pendingVehicleRequestsRS("request_id") %>" class="shadcn-btn shadcn-btn-primary" style="padding: 2px 8px; font-size: 0.75rem;">상세/승인</a>
                            </td>
                        </tr>
                        <% 
                            pendingVehicleRequestsRS.MoveNext
                            Loop 
                        %>
                    </tbody>
                </table>
                <div style="margin-top: 10px; text-align: right;">
                    <a href="vehicle_request_manage.asp?status=대기" class="shadcn-btn shadcn-btn-outline">모든 대기 신청 보기</a>
                </div>
            <% End If %>
        </div>
    </div>
    
    <!-- 최근 활동 로그 -->
    <div class="shadcn-card">
        <div class="shadcn-card-header">
            <h3 class="shadcn-card-title">최근 활동 로그</h3>
        </div>
        <div class="shadcn-card-content">
            <% If recentLogsRS.EOF Then %>
                <div style="text-align: center; padding: 20px;">
                    <p>최근 활동 로그가 없습니다.</p>
                </div>
            <% Else %>
                <table class="shadcn-table">
                    <thead class="shadcn-table-header">
                        <tr>
                            <th>시간</th>
                            <th>사용자</th>
                            <th>작업</th>
                            <th>설명</th>
                        </tr>
                    </thead>
                    <tbody>
                        <% Do While Not recentLogsRS.EOF %>
                        <tr>
                            <td><%= FormatDateTime(recentLogsRS("created_at"), 2) & " " & FormatDateTime(recentLogsRS("created_at"), 4) %></td>
                            <td><%= recentLogsRS("name") & " (" & recentLogsRS("user_id") & ")" %></td>
                            <td><%= recentLogsRS("action") %></td>
                            <td><%= recentLogsRS("description") %></td>
                        </tr>
                        <% 
                            recentLogsRS.MoveNext
                            Loop 
                        %>
                    </tbody>
                </table>
                <div style="margin-top: 10px; text-align: right;">
                    <a href="activity_logs.asp" class="shadcn-btn shadcn-btn-outline">모든 로그 보기</a>
                </div>
            <% End If %>
        </div>
    </div>
</div>

<!--#include file="../includes/footer.asp"--> 