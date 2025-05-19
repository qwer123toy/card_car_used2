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
    RedirectTo("/contents/card_car_used/index.asp")
End If

' 사용자 정보 조회
Dim SQL, rs, userName, userDeptId, userDeptName
On Error Resume Next

' 세션에서 사용자 이름 불러오기 (DB 오류 시 대비)
If Session("name") <> "" Then
    userName = Session("name")
Else
    userName = Session("user_id")
End If

' 세션에서 부서 ID 불러오기
If Session("department_id") <> "" Then
    userDeptId = Session("department_id")
Else
    userDeptId = 0
End If

' 부서명 기본값 설정
userDeptName = "부서 정보 없음"

' DB에서 사용자 정보 조회 시도
SQL = "SELECT u.name, u.department_id, d.name AS department_name FROM Users u " & _
      "LEFT JOIN Department d ON u.department_id = d.department_id " & _
      "WHERE u.user_id = '" & Session("user_id") & "'"
      
' 파라미터화된 쿼리 사용을 위해 명령 객체 생성
Dim cmd
Set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = db
cmd.CommandText = "SELECT u.name, u.department_id, d.name AS department_name FROM Users u " & _
                 "LEFT JOIN Department d ON u.department_id = d.department_id " & _
                 "WHERE u.user_id = ?"
cmd.Parameters.Append cmd.CreateParameter("@user_id", 200, 1, 30, Session("user_id"))

' 명령 실행
Set rs = cmd.Execute()

' DB 조회에 성공했고 사용자 정보가 있는 경우
If Err.Number = 0 And Not rs.EOF Then
    userName = rs("name")
    userDeptId = rs("department_id")
    
    ' NULL 값 처리
    If Not IsNull(rs("department_name")) Then
    userDeptName = rs("department_name")
End If
    
rs.Close
End If

' 최근 카드 사용 내역 조회 시도
Dim recentCardRS
Set recentCardRS = Server.CreateObject("ADODB.Recordset")

' 빈 레코드셋 초기화 (오류 발생 시 이 레코드셋 사용)
recentCardRS.Fields.Append "usage_id", 3 ' adInteger
recentCardRS.Fields.Append "usage_date", 7 ' adDate
recentCardRS.Fields.Append "amount", 6 ' adCurrency
recentCardRS.Fields.Append "account_name", 200, 100 ' adVarChar
recentCardRS.Open

' 실제 데이터 조회 시도
Dim recentCardCmd
Set recentCardCmd = Server.CreateObject("ADODB.Command")
recentCardCmd.ActiveConnection = db
recentCardCmd.CommandText = "SELECT TOP 5 c.usage_id, c.usage_date, c.amount, ca.account_name " & _
                "FROM CardUsage c " & _
                "JOIN CardAccount ca ON c.card_id = ca.card_id " & _
                          "WHERE c.user_id = ? " & _
                "ORDER BY c.usage_date DESC"
recentCardCmd.Parameters.Append recentCardCmd.CreateParameter("@user_id", 200, 1, 30, Session("user_id"))

On Error Resume Next
Dim tempRS
Set tempRS = recentCardCmd.Execute()

' 데이터 조회 성공 시 임시 레코드셋에서 실제 레코드셋으로 데이터 복사
If Err.Number = 0 And Not tempRS.EOF Then
    ' 기존 빈 레코드셋 닫고 정상 레코드셋으로 대체
    recentCardRS.Close
    Set recentCardRS = tempRS
End If
On Error GoTo 0

' 최근 차량 사용 내역 조회 시도
Dim recentVehicleRS
Set recentVehicleRS = Server.CreateObject("ADODB.Recordset")

' 빈 레코드셋 초기화 (오류 발생 시 이 레코드셋 사용)
recentVehicleRS.Fields.Append "request_id", 3 ' adInteger
recentVehicleRS.Fields.Append "request_date", 7 ' adDate
recentVehicleRS.Fields.Append "purpose", 200, 100 ' adVarChar
recentVehicleRS.Fields.Append "approval_status", 200, 20 ' adVarChar
recentVehicleRS.Open

' 실제 데이터 조회 시도
Dim recentVehicleCmd
Set recentVehicleCmd = Server.CreateObject("ADODB.Command")
recentVehicleCmd.ActiveConnection = db
recentVehicleCmd.CommandText = "SELECT TOP 5 v.request_id, v.request_date, v.purpose, v.approval_status " & _
                   "FROM VehicleRequests v " & _
                            "WHERE v.user_id = ? AND v.is_deleted = 0 " & _
                   "ORDER BY v.request_date DESC"
recentVehicleCmd.Parameters.Append recentVehicleCmd.CreateParameter("@user_id", 200, 1, 30, Session("user_id"))

On Error Resume Next
Dim tempVehicleRS
Set tempVehicleRS = recentVehicleCmd.Execute()

' 데이터 조회 성공 시 임시 레코드셋에서 실제 레코드셋으로 데이터 복사
If Err.Number = 0 And Not tempVehicleRS.EOF Then
    ' 기존 빈 레코드셋 닫고 정상 레코드셋으로 대체
    recentVehicleRS.Close
    Set recentVehicleRS = tempVehicleRS
End If
On Error GoTo 0
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
        <div class="shadcn-card-footer" style="display: flex; justify-content: flex-end;">
            <a href="/contents/card_car_used/pages/my_profile.asp" class="shadcn-btn shadcn-btn-outline">내 정보 보기</a>
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
                <a href="/contents/card_car_used/pages/card_usage.asp" class="shadcn-btn shadcn-btn-outline">모든 내역 보기</a>
                <a href="/contents/card_car_used/pages/card_usage_add.asp" class="shadcn-btn shadcn-btn-primary">새 내역 등록</a>
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
                <a href="/contents/card_car_used/pages/vehicle_request.asp" class="shadcn-btn shadcn-btn-outline">모든 내역 보기</a>
                <a href="/contents/card_car_used/pages/vehicle_request_add.asp" class="shadcn-btn shadcn-btn-primary">새 신청서 작성</a>
            </div>
        </div>
    </div>
</div>

<!--#include file="../includes/footer.asp"--> 