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

On Error Resume Next

' URL 파라미터에서 신청서 ID 추출
Dim requestId, errorMsg, successMsg
requestId = PreventSQLInjection(Request.QueryString("id"))

If requestId = "" Then
    errorMsg = "잘못된 접근입니다. 신청서 ID가 필요합니다."
    Response.Redirect("vehicle_request.asp")
End If

' 신청서 정보 조회
Dim cmd, rs
Set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = db
cmd.CommandText = "SELECT v.*, u.name AS user_name " & _
                 "FROM VehicleRequests v " & _
                 "INNER JOIN Users u ON v.user_id = u.user_id " & _
                 "WHERE v.request_id = ? AND v.is_deleted = 0"
cmd.Parameters.Append cmd.CreateParameter("@request_id", 3, 1, , CLng(requestId))

Set rs = cmd.Execute()

If Err.Number <> 0 Or rs.EOF Then
    errorMsg = "요청하신 신청서를 찾을 수 없습니다: " & Err.Description
    
    ' 오류 로그 기록
    If Err.Number <> 0 Then
        LogActivity Session("user_id"), "오류", "차량 이용 신청서 조회 실패 (ID: " & requestId & ", 오류: " & Err.Description & ")"
    End If
    
    Set rs = Nothing
End If

' 최신 유류비 단가 조회
Dim fuelRateSQL, fuelRateRS, fuelRate
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

On Error GoTo 0
%>
<!--#include file="../includes/header.asp"-->

<div class="vehicle-request-view-container">
    <div class="shadcn-card" style="max-width: 800px; margin: 30px auto;">
        <div class="shadcn-card-header">
            <h2 class="shadcn-card-title">차량 이용 신청서 상세정보</h2>
            <p class="shadcn-card-description">신청서 세부 내용을 확인합니다.</p>
        </div>
        
        <% If errorMsg <> "" Then %>
        <div class="shadcn-alert shadcn-alert-error">
            <div>
                <span class="shadcn-alert-title">오류</span>
                <span class="shadcn-alert-description"><%= errorMsg %></span>
            </div>
        </div>
        <% End If %>
        
        <% If Not rs Is Nothing And Not rs.EOF Then %>
        <div class="shadcn-card-content">
            <div class="info-group" style="margin-bottom: 20px; padding-bottom: 15px; border-bottom: 1px solid #eee;">
                <div style="display: flex; justify-content: space-between; align-items: center;">
                    <h3 style="margin: 0; font-size: 1.2rem;">신청 정보</h3>
                    <span class="shadcn-badge <%= GetStatusBadgeClass(rs("approval_status")) %>">
                        <%= rs("approval_status") %>
                    </span>
                </div>
            </div>
            
            <div class="info-grid" style="display: grid; grid-template-columns: repeat(2, 1fr); gap: 15px;">
                <div class="info-item">
                    <span class="info-label">신청번호</span>
                    <span class="info-value"><%= rs("request_id") %></span>
                </div>
                
                <div class="info-item">
                    <span class="info-label">신청자</span>
                    <span class="info-value"><%= rs("user_name") %></span>
                </div>
                
                <div class="info-item">
                    <span class="info-label">시작일자</span>
                    <span class="info-value"><%= FormatDate(rs("start_date")) %></span>
                </div>
                
                <div class="info-item">
                    <span class="info-label">종료일자</span>
                    <span class="info-value"><%= FormatDate(rs("end_date")) %></span>
                </div>
                
                <div class="info-item" style="grid-column: span 2;">
                    <span class="info-label">업무 목적</span>
                    <span class="info-value"><%= rs("purpose") %></span>
                </div>
                
                <div class="info-item">
                    <span class="info-label">출발지</span>
                    <span class="info-value"><%= rs("start_location") %></span>
                </div>
                
                <div class="info-item">
                    <span class="info-label">목적지</span>
                    <span class="info-value"><%= rs("destination") %></span>
                </div>
                
                <div class="info-item">
                    <span class="info-label">운행거리</span>
                    <span class="info-value"><%= rs("distance") %> km</span>
                </div>
                
                <div class="info-item">
                    <span class="info-label">유류비 단가</span>
                    <span class="info-value"><%= FormatNumber(fuelRate) %> 원</span>
                </div>
                
                <div class="info-item">
                    <span class="info-label">총 금액</span>
                    <span class="info-value"><%= FormatNumber(CDbl(rs("distance")) * CDbl(fuelRate)) %> 원</span>
                </div>
            </div>
            
            <div class="shadcn-card-footer" style="margin-top: 1.5rem; display: flex; justify-content: space-between;">
                <div>
                    <a href="vehicle_request.asp" class="shadcn-btn shadcn-btn-outline">목록으로</a>
                </div>
                
                <% If Session("user_id") = rs("user_id") And rs("approval_status") = "작성중" Then %>
                <div>
                    <a href="vehicle_request_edit.asp?id=<%= rs("request_id") %>" class="shadcn-btn shadcn-btn-secondary">수정</a>
                    <button type="button" class="shadcn-btn shadcn-btn-destructive" data-request-id="<%= rs("request_id") %>" onclick="confirmDelete(this.getAttribute('data-request-id'))">삭제</button>
                </div>
                <% End If %>
            </div>
        </div>
        <% 
            rs.Close
            Set rs = Nothing
        End If 
        %>
    </div>
</div>

<script>
    function confirmDelete(requestId) {
        if (confirm('정말로 이 신청서를 삭제하시겠습니까?')) {
            window.location.href = 'vehicle_request_delete.asp?id=' + requestId;
        }
    }
</script>

<style>
    .info-item {
        margin-bottom: 8px;
    }
    
    .info-label {
        display: block;
        font-size: 0.85rem;
        color: #666;
        margin-bottom: 2px;
    }
    
    .info-value {
        font-weight: 500;
    }
    
    .shadcn-badge {
        display: inline-flex;
        align-items: center;
        justify-content: center;
        border-radius: 4px;
        padding: 0.25rem 0.5rem;
        font-size: 0.75rem;
        font-weight: 500;
    }
    
    .shadcn-badge-primary {
        background-color: #0070f3;
        color: white;
    }
    
    .shadcn-badge-secondary {
        background-color: #666;
        color: white;
    }
    
    .shadcn-badge-destructive {
        background-color: #ff4c4c;
        color: white;
    }
    
    .shadcn-badge-outline {
        background-color: transparent;
        border: 1px solid #ddd;
        color: #666;
    }
</style>

<%
Function GetStatusBadgeClass(status)
    Select Case status
        Case "승인"
            GetStatusBadgeClass = "shadcn-badge-primary"
        Case "반려"
            GetStatusBadgeClass = "shadcn-badge-destructive"
        Case "작성중"
            GetStatusBadgeClass = "shadcn-badge-secondary"
        Case Else
            GetStatusBadgeClass = "shadcn-badge-outline"
    End Select
End Function
%>

<!--#include file="../includes/footer.asp"--> 