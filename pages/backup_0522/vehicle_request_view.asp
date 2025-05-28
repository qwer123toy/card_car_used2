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

' 결재선 정보 조회
Dim approvalRS, approvalSQL
approvalSQL = "SELECT al.*, u.name AS approver_name, u.department_id, " & _
             "d.name AS department_name, u.job_grade, j.name AS job_grade_name " & _
             "FROM " & dbSchema & ".ApprovalLogs al " & _
             "JOIN " & dbSchema & ".Users u ON al.approver_id = u.user_id " & _
             "LEFT JOIN " & dbSchema & ".Department d ON u.department_id = d.department_id " & _
             "LEFT JOIN " & dbSchema & ".Job_Grade j ON u.job_grade = j.job_grade_id " & _
             "WHERE al.target_table_name = 'VehicleRequests' AND al.target_id = ? " & _
             "ORDER BY al.approval_step"

Set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = db
cmd.CommandText = approvalSQL
cmd.Parameters.Append cmd.CreateParameter("@target_id", 3, 1, , requestId)
Set approvalRS = cmd.Execute()

On Error GoTo 0
%>
<!--#include file="../includes/header.asp"-->

<div class="vehicle-request-view-container">
    <div class="shadcn-card" style="max-width: 850px; margin: 30px auto; box-shadow: 0 4px 12px rgba(0,0,0,0.08); border-radius: 16px;">
        <div class="shadcn-card-header" style="background: linear-gradient(to right, #4A90E2, #5A9EEA); padding: 1.75rem 2rem; border-radius: 16px 16px 0 0;">
            <h2 class="shadcn-card-title" style="color: white; font-size: 1.5rem; margin-bottom: 0.25rem;">차량 이용 신청서 상세정보</h2>
            <p class="shadcn-card-description" style="color: rgba(255,255,255,0.8); margin: 0;">신청서 세부 내용 및 결재 정보를 확인합니다.</p>
        </div>
        
        <% If errorMsg <> "" Then %>
        <div class="shadcn-alert shadcn-alert-error" style="margin: 1.5rem; padding: 1rem; background-color: #FFEEEE; border-left: 4px solid #E53E3E; border-radius: 8px;">
            <div>
                <span class="shadcn-alert-title" style="font-weight: 600; color: #E53E3E;">오류</span>
                <span class="shadcn-alert-description" style="color: #333;"><%= errorMsg %></span>
            </div>
        </div>
        <% End If %>
        
        <% If Not rs Is Nothing And Not rs.EOF Then %>
        <div class="shadcn-card-content" style="padding: 2rem;">
            <div class="info-group" style="margin-bottom: 1.75rem; display: flex; justify-content: space-between; align-items: center; border-bottom: 1px solid #eee; padding-bottom: 1.25rem;">
                <div style="display: flex; align-items: center; gap: 1rem;">
                    <div style="background: #F0F7FF; border-radius: 50%; width: 48px; height: 48px; display: flex; align-items: center; justify-content: center;">
                        <i class="fas fa-car" style="color: #4A90E2; font-size: 1.25rem;"></i>
                    </div>
                    <div>
                        <h3 style="margin: 0; font-size: 1.25rem; font-weight: 600; color: #2C3E50;">신청서 #<%= rs("request_id") %></h3>
                        <p style="margin: 0; font-size: 0.9rem; color: #64748B;"><%= FormatDate(rs("request_date")) %> 신청됨</p>
                    </div>
                </div>
                <span class="shadcn-badge <%= GetStatusBadgeClass(rs("approval_status")) %>" style="font-size: 0.85rem; padding: 0.5rem 1rem; border-radius: 6px;">
                    <%= rs("approval_status") %>
                </span>
            </div>
            
            <div class="info-section" style="background: #F8FAFC; border-radius: 12px; padding: 1.5rem; margin-bottom: 2rem;">
                <h4 style="margin-top: 0; margin-bottom: 1.25rem; font-size: 1.1rem; color: #2C3E50; font-weight: 600; border-bottom: 1px solid #E9ECEF; padding-bottom: 0.75rem;">
                    <i class="fas fa-info-circle" style="margin-right: 0.5rem; color: #4A90E2;"></i>신청 정보
                </h4>
                
                <div class="info-grid" style="display: grid; grid-template-columns: repeat(2, 1fr); gap: 1.25rem;">
                    <div class="info-item" style="background: white; padding: 1rem; border-radius: 8px; box-shadow: 0 1px 3px rgba(0,0,0,0.05);">
                        <span class="info-label" style="display: block; font-size: 0.85rem; color: #64748B; margin-bottom: 0.375rem;">신청자</span>
                        <span class="info-value" style="font-weight: 500; font-size: 1rem; color: #2C3E50;"><%= rs("user_name") %></span>
                    </div>
                    
                    <div class="info-item" style="background: white; padding: 1rem; border-radius: 8px; box-shadow: 0 1px 3px rgba(0,0,0,0.05);">
                        <span class="info-label" style="display: block; font-size: 0.85rem; color: #64748B; margin-bottom: 0.375rem;">기간</span>
                        <span class="info-value" style="font-weight: 500; font-size: 1rem; color: #2C3E50;">
                            <%= FormatDate(rs("start_date")) %> ~ <%= FormatDate(rs("end_date")) %>
                        </span>
                    </div>
                    
                    <div class="info-item" style="background: white; padding: 1rem; border-radius: 8px; box-shadow: 0 1px 3px rgba(0,0,0,0.05);">
                        <span class="info-label" style="display: block; font-size: 0.85rem; color: #64748B; margin-bottom: 0.375rem;">출발지</span>
                        <span class="info-value" style="font-weight: 500; font-size: 1rem; color: #2C3E50;"><%= rs("start_location") %></span>
                    </div>
                    
                    <div class="info-item" style="background: white; padding: 1rem; border-radius: 8px; box-shadow: 0 1px 3px rgba(0,0,0,0.05);">
                        <span class="info-label" style="display: block; font-size: 0.85rem; color: #64748B; margin-bottom: 0.375rem;">목적지</span>
                        <span class="info-value" style="font-weight: 500; font-size: 1rem; color: #2C3E50;"><%= rs("destination") %></span>
                    </div>
                    
                    <div class="info-item" style="background: white; padding: 1rem; border-radius: 8px; box-shadow: 0 1px 3px rgba(0,0,0,0.05); grid-column: span 2;">
                        <span class="info-label" style="display: block; font-size: 0.85rem; color: #64748B; margin-bottom: 0.375rem;">업무 목적</span>
                        <span class="info-value" style="font-weight: 500; font-size: 1rem; color: #2C3E50;"><%= rs("purpose") %></span>
                    </div>
                </div>
            </div>
            
            <div class="info-section" style="background: #F8FAFC; border-radius: 12px; padding: 1.5rem; margin-bottom: 2rem;">
                <h4 style="margin-top: 0; margin-bottom: 1.25rem; font-size: 1.1rem; color: #2C3E50; font-weight: 600; border-bottom: 1px solid #E9ECEF; padding-bottom: 0.75rem;">
                    <i class="fas fa-calculator" style="margin-right: 0.5rem; color: #4A90E2;"></i>비용 정보
                </h4>
                
                <div class="info-grid" style="display: grid; grid-template-columns: repeat(3, 1fr); gap: 1.25rem;">
                    <div class="info-item" style="background: white; padding: 1rem; border-radius: 8px; box-shadow: 0 1px 3px rgba(0,0,0,0.05);">
                        <span class="info-label" style="display: block; font-size: 0.85rem; color: #64748B; margin-bottom: 0.375rem;">운행거리</span>
                        <span class="info-value" style="font-weight: 500; font-size: 1rem; color: #2C3E50;"><%= FormatNumber(rs("distance")) %> km</span>
                    </div>
                    
                    <div class="info-item" style="background: white; padding: 1rem; border-radius: 8px; box-shadow: 0 1px 3px rgba(0,0,0,0.05);">
                        <span class="info-label" style="display: block; font-size: 0.85rem; color: #64748B; margin-bottom: 0.375rem;">유류비 단가</span>
                        <span class="info-value" style="font-weight: 500; font-size: 1rem; color: #2C3E50;"><%= FormatNumber(fuelRate) %> 원</span>
                    </div>
                    
                    <div class="info-item" style="background: white; padding: 1rem; border-radius: 8px; box-shadow: 0 1px 3px rgba(0,0,0,0.05);">
                        <span class="info-label" style="display: block; font-size: 0.85rem; color: #64748B; margin-bottom: 0.375rem;">유류비 합계</span>
                        <span class="info-value" style="font-weight: 500; font-size: 1rem; color: #2C3E50;"><%= FormatNumber(CDbl(rs("distance")) * CDbl(fuelRate)) %> 원</span>
                    </div>
                    
                    <div class="info-item" style="background: white; padding: 1rem; border-radius: 8px; box-shadow: 0 1px 3px rgba(0,0,0,0.05);">
                        <span class="info-label" style="display: block; font-size: 0.85rem; color: #64748B; margin-bottom: 0.375rem;">통행료</span>
                        <span class="info-value" style="font-weight: 500; font-size: 1rem; color: #2C3E50;">
                            <% If Not IsNull(rs("toll_fee")) Then %>
                                <%= FormatNumber(rs("toll_fee")) %> 원
                            <% Else %>
                                0 원
                            <% End If %>
                        </span>
                    </div>
                    
                    <div class="info-item" style="background: white; padding: 1rem; border-radius: 8px; box-shadow: 0 1px 3px rgba(0,0,0,0.05);">
                        <span class="info-label" style="display: block; font-size: 0.85rem; color: #64748B; margin-bottom: 0.375rem;">주차비</span>
                        <span class="info-value" style="font-weight: 500; font-size: 1rem; color: #2C3E50;">
                            <% If Not IsNull(rs("parking_fee")) Then %>
                                <%= FormatNumber(rs("parking_fee")) %> 원
                            <% Else %>
                                0 원
                            <% End If %>
                        </span>
                    </div>
                    
                    <div class="info-item" style="background: #F0F7FF; padding: 1rem; border-radius: 8px; box-shadow: 0 1px 3px rgba(0,0,0,0.05); border-left: 3px solid #4A90E2;">
                        <span class="info-label" style="display: block; font-size: 0.85rem; color: #64748B; margin-bottom: 0.375rem;">총 예상 비용</span>
                        <span class="info-value" style="font-weight: 600; font-size: 1.1rem; color: #2C3E50;">
                            <% 
                            Dim totalCost, tollFeeCost, parkingFeeCost
                            
                            tollFeeCost = 0
                            If Not IsNull(rs("toll_fee")) Then
                                tollFeeCost = CDbl(rs("toll_fee"))
                            End If
                            
                            parkingFeeCost = 0
                            If Not IsNull(rs("parking_fee")) Then
                                parkingFeeCost = CDbl(rs("parking_fee"))
                            End If
                            
                            totalCost = (CDbl(rs("distance")) * CDbl(fuelRate)) + tollFeeCost + parkingFeeCost
                            Response.Write FormatNumber(totalCost) & " 원"
                            %>
                        </span>
                    </div>
                </div>
            </div>
            
            <!-- 결재선 정보 -->
            <% If Not approvalRS.EOF Then %>
            <div class="info-section" style="background: #F8FAFC; border-radius: 12px; padding: 1.5rem; margin-bottom: 2rem;">
                <h4 style="margin-top: 0; margin-bottom: 1.25rem; font-size: 1.1rem; color: #2C3E50; font-weight: 600; border-bottom: 1px solid #E9ECEF; padding-bottom: 0.75rem;">
                    <i class="fas fa-user-check" style="margin-right: 0.5rem; color: #4A90E2;"></i>결재선 정보
                </h4>
                
                <div class="approval-line" style="display: flex; flex-direction: column; gap: 1rem;">
                    <% 
                    approvalRS.MoveFirst
                    Do While Not approvalRS.EOF 
                        Dim statusClass
                        Select Case approvalRS("status")
                            Case "승인"
                                statusClass = "shadcn-badge-success"
                            Case "반려"
                                statusClass = "shadcn-badge-destructive"
                            Case "대기"
                                statusClass = "shadcn-badge-secondary"
                        End Select
                    %>
                    <div class="approval-step" style="display: flex; padding: 1.25rem; border: 1px solid #E9ECEF; border-radius: 10px; background-color: white; box-shadow: 0 1px 3px rgba(0,0,0,0.05); transition: transform 0.2s ease;">
                        <div style="width: 100px; text-align: center; padding: 0.5rem 0.75rem; background-color: #F1F5F9; border-radius: 6px; margin-right: 1.25rem; font-weight: 600; color: #475569; font-size: 0.9rem;">
                            <%= approvalRS("approval_step") %>차 결재
                        </div>
                        <div style="flex-grow: 1;">
                            <div style="font-weight: 600; font-size: 1.05rem; color: #2C3E50; margin-bottom: 0.25rem;"><%= approvalRS("approver_name") %></div>
                            <div style="font-size: 0.9rem; color: #64748B; margin-bottom: 0.75rem;">
                                <% If Not IsNull(approvalRS("department_name")) Then %>
                                    <%= approvalRS("department_name") %>
                                    <% If Not IsNull(approvalRS("job_grade_name")) Then %> / <%= approvalRS("job_grade_name") %><% End If %>
                                <% ElseIf Not IsNull(approvalRS("job_grade_name")) Then %>
                                    <%= approvalRS("job_grade_name") %>
                                <% End If %>
                            </div>
                            <div style="display: flex; justify-content: space-between; align-items: center; margin-top: 0.5rem;">
                                <span class="shadcn-badge <%= statusClass %>" style="font-size: 0.85rem; padding: 0.375rem 0.75rem; border-radius: 6px;"><%= approvalRS("status") %></span>
                                <% If Not IsNull(approvalRS("approved_at")) Then %>
                                    <span style="font-size: 0.9rem; color: #64748B;"><i class="far fa-clock" style="margin-right: 0.375rem;"></i><%= FormatDateTime(approvalRS("approved_at"), 2) %></span>
                                <% End If %>
                            </div>
                            <% If Not IsNull(approvalRS("comments")) And approvalRS("comments") <> "" Then %>
                                <div style="margin-top: 0.875rem; padding: 0.75rem; background-color: #F1F5F9; border-radius: 6px; font-size: 0.9rem; color: #475569;">
                                    <i class="fas fa-comment" style="margin-right: 0.5rem; color: #64748B;"></i>
                                    <%= approvalRS("comments") %>
                                </div>
                            <% End If %>
                        </div>
                    </div>
                    <%
                        approvalRS.MoveNext
                    Loop
                    %>
                </div>
            </div>
            <% End If %>
            
            <div class="shadcn-card-footer" style="margin-top: 1.5rem; display: flex; justify-content: space-between; border-top: 1px solid #E9ECEF; padding-top: 1.5rem;">
                <div>
                    <a href="vehicle_request.asp" class="shadcn-btn" style="background-color: #F8FAFC; color: #475569; border: 1px solid #E9ECEF; padding: 0.75rem 1.5rem; border-radius: 8px; font-weight: 600; text-decoration: none; display: inline-flex; align-items: center; transition: all 0.2s ease;">
                        <i class="fas fa-arrow-left" style="margin-right: 0.5rem;"></i>목록으로
                    </a>
                </div>
                
                <div style="display: flex; gap: 0.75rem;">
                    <% If Session("user_id") = rs("user_id") And (rs("approval_status") = "반려" Or rs("approval_status") = "대기") Then %>
                        <a href="vehicle_request_edit.asp?id=<%= rs("request_id") %>" class="shadcn-btn" style="background-color: #4A90E2; color: white; border: none; padding: 0.75rem 1.5rem; border-radius: 8px; font-weight: 600; text-decoration: none; display: inline-flex; align-items: center; transition: all 0.2s ease;">
                            <i class="fas fa-edit" style="margin-right: 0.5rem;"></i>수정
                        </a>
                        <button type="button" class="shadcn-btn" style="background-color: #E53E3E; color: white; border: none; padding: 0.75rem 1.5rem; border-radius: 8px; font-weight: 600; cursor: pointer; display: inline-flex; align-items: center; transition: all 0.2s ease;" data-request-id="<%= rs("request_id") %>" onclick="confirmDelete(this.getAttribute('data-request-id'))">
                            <i class="fas fa-trash-alt" style="margin-right: 0.5rem;"></i>삭제
                        </button>
                    <% End If %>
                    
                    <% If Session("user_id") = rs("user_id") And (rs("approval_status") = "대기" Or rs("approval_status") = "완료") Then %>
                        <a href="approval_detail.asp?id=<%= rs("request_id") %>&type=VehicleRequests" class="shadcn-btn" style="background-color: #2C3E50; color: white; border: none; padding: 0.75rem 1.5rem; border-radius: 8px; font-weight: 600; text-decoration: none; display: inline-flex; align-items: center; transition: all 0.2s ease;">
                            <i class="fas fa-file-alt" style="margin-right: 0.5rem;"></i>결재 정보 상세
                        </a>
                    <% End If %>
                </div>
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
    
    // 호버 효과 추가
    document.addEventListener('DOMContentLoaded', function() {
        const approvalSteps = document.querySelectorAll('.approval-step');
        if (approvalSteps) {
            approvalSteps.forEach(step => {
                step.addEventListener('mouseenter', function() {
                    this.style.transform = 'translateY(-3px)';
                    this.style.boxShadow = '0 4px 12px rgba(0,0,0,0.1)';
                });
                
                step.addEventListener('mouseleave', function() {
                    this.style.transform = 'translateY(0)';
                    this.style.boxShadow = '0 1px 3px rgba(0,0,0,0.05)';
                });
            });
        }
        
        const infoItems = document.querySelectorAll('.info-item');
        if (infoItems) {
            infoItems.forEach(item => {
                item.addEventListener('mouseenter', function() {
                    this.style.boxShadow = '0 4px 12px rgba(0,0,0,0.07)';
                });
                
                item.addEventListener('mouseleave', function() {
                    this.style.boxShadow = '0 1px 3px rgba(0,0,0,0.05)';
                });
            });
        }
    });
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
    
    .shadcn-badge-success {
        background-color: #E3F9E5;
        color: #1B873F;
    }
    
    .shadcn-badge-destructive {
        background-color: #FFE9E9;
        color: #DA3633;
    }
    
    .shadcn-badge-secondary {
        background-color: #F6F8FA;
        color: #57606A;
    }
</style>

<%
Function GetStatusBadgeClass(status)
    Select Case status
        Case "승인"
            GetStatusBadgeClass = "shadcn-badge-success"
        Case "반려"
            GetStatusBadgeClass = "shadcn-badge-destructive"
        Case "대기"
            GetStatusBadgeClass = "shadcn-badge-secondary"
        Case "완료"
            GetStatusBadgeClass = "shadcn-badge-success"
        Case Else
            GetStatusBadgeClass = "shadcn-badge-outline"
    End Select
End Function
%>

<!--#include file="../includes/footer.asp"--> 