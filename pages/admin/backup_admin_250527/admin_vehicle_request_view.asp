<%@ Language="VBScript" CodePage="65001" %>
<% 
Response.CodePage = 65001
Response.CharSet = "utf-8"
%>

<!--#include file="../../db.asp"-->
<!--#include file="../../includes/functions.asp"-->
<%
' 로그인 체크
If Not IsAuthenticated() Then
    RedirectTo("../../index.asp")
End If

' 관리자 권한 체크
If Not IsAdmin() Then
    Response.Write("<script>alert('관리자 권한이 필요합니다.'); window.location.href='../dashboard.asp';</script>")
    Response.End
End If

' 요청 ID 확인
Dim requestId
requestId = Request.QueryString("id")

If requestId = "" Then
    Response.Write("<script>alert('잘못된 접근입니다.'); window.location.href='admin_vehicle_requests.asp';</script>")
    Response.End
End If

' 차량 이용 신청 정보 조회
Dim requestSQL, requestRS
requestSQL = "SELECT vr.*, " & _
             "u.user_id as user_id, u.email AS user_email, d.name AS department_name " & _
             "FROM " & dbSchema & ".VehicleRequests vr " & _
             "LEFT JOIN " & dbSchema & ".Users u ON vr.user_id = u.user_id " & _
             "LEFT JOIN " & dbSchema & ".Department d ON u.department_id = d.department_id " & _
             "WHERE vr.request_id = " & requestId

Set requestRS = db99.Execute(requestSQL)

' 데이터가 없으면 목록으로 리다이렉션
If requestRS.EOF Then
    Response.Write("<script>alert('해당 차량 이용 신청 정보를 찾을 수 없습니다.'); window.location.href='admin_vehicle_requests.asp';</script>")
    Response.End
End If

' 상태 처리 (승인/거부)
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    Dim action, updateSQL, rejectReason
    action = Request.Form("action")
    
    If action = "approve" Then
        ' 승인 처리
        updateSQL = "UPDATE " & dbSchema & ".VehicleRequests " & _
                   "SET status = '승인됨', approval_date = GETDATE() " & _
                   "WHERE request_id = " & requestId
        
        On Error Resume Next
        db99.Execute(updateSQL)
        
        If Err.Number <> 0 Then
            Response.Write("<script>alert('차량 이용 신청 승인 중 오류가 발생했습니다: " & Server.HTMLEncode(Err.Description) & "'); window.location.href='admin_vehicle_request_view.asp?id=" & requestId & "';</script>")
        Else
            ' 활동 로그 기록
            LogActivity Session("user_id"), "차량이용신청승인", "차량 이용 신청 승인 (ID: " & requestId & ")"
            Response.Write("<script>alert('차량 이용 신청이 승인되었습니다.'); window.location.href='admin_vehicle_requests.asp';</script>")
        End If
        On Error GoTo 0
        Response.End
        
    ElseIf action = "reject" Then
        ' 거부 처리
        rejectReason = PreventSQLInjection(Request.Form("rejection_reason"))
        
        updateSQL = "UPDATE " & dbSchema & ".VehicleRequests " & _
                   "SET status = '거부됨', rejection_reason = '" & rejectReason & "', approval_date = GETDATE() " & _
                   "WHERE request_id = " & requestId
        
        On Error Resume Next
        db99.Execute(updateSQL)
        
        If Err.Number <> 0 Then
            Response.Write("<script>alert('차량 이용 신청 거부 중 오류가 발생했습니다: " & Server.HTMLEncode(Err.Description) & "'); window.location.href='admin_vehicle_request_view.asp?id=" & requestId & "';</script>")
        Else
            ' 활동 로그 기록
            LogActivity Session("user_id"), "차량이용신청거부", "차량 이용 신청 거부 (ID: " & requestId & ")"
            Response.Write("<script>alert('차량 이용 신청이 거부되었습니다.'); window.location.href='admin_vehicle_requests.asp';</script>")
        End If
        On Error GoTo 0
        Response.End
    End If
End If

' 날짜 포맷
Function FormatDate(dateValue)
    If IsNull(dateValue) Or Not IsDate(dateValue) Then
        FormatDate = ""
    Else
        FormatDate = Year(dateValue) & "-" & Right("0" & Month(dateValue), 2) & "-" & Right("0" & Day(dateValue), 2)
    End If
End Function


%>

<!--#include file="../../includes/header.asp"-->

<div class="container-fluid my-4">
    <div class="row">
        <div class="col-md-3">
            <!-- 사이드바 메뉴 -->
            <div class="card shadow-sm mb-4">
                <div class="card-header bg-primary text-white">
                    <h5 class="mb-0"><i class="fas fa-cog me-2"></i>관리 메뉴</h5>
                </div>
                <div class="list-group list-group-flush">
                    <a href="admin_dashboard.asp" class="list-group-item list-group-item-action">
                        <i class="fas fa-tachometer-alt me-2"></i>대시보드
                    </a>
                    <a href="admin_cardaccount.asp" class="list-group-item list-group-item-action">
                        <i class="fas fa-credit-card me-2"></i>카드 계정 관리
                    </a>
                    <a href="admin_cardaccounttypes.asp" class="list-group-item list-group-item-action">
                        <i class="fas fa-tags me-2"></i>카드 계정 유형 관리
                    </a>
                    <a href="admin_fuelrate.asp" class="list-group-item list-group-item-action">
                        <i class="fas fa-gas-pump me-2"></i>유류비 단가 관리
                    </a>
                    <a href="admin_job_grade.asp" class="list-group-item list-group-item-action">
                        <i class="fas fa-user-tie me-2"></i>직급 관리
                    </a>
                    <a href="admin_department.asp" class="list-group-item list-group-item-action">
                        <i class="fas fa-sitemap me-2"></i>부서 관리
                    </a>
                    <a href="admin_users.asp" class="list-group-item list-group-item-action">
                        <i class="fas fa-users me-2"></i>사용자 관리
                    </a>
                    <a href="admin_card_usage.asp" class="list-group-item list-group-item-action">
                        <i class="fas fa-receipt me-2"></i>카드 사용 내역 관리
                    </a>
                    <a href="admin_vehicle_requests.asp" class="list-group-item list-group-item-action active">
                        <i class="fas fa-car me-2"></i>차량 이용 신청 관리
                    </a>
                    <a href="admin_approvals.asp" class="list-group-item list-group-item-action">
                        <i class="fas fa-file-signature me-2"></i>결재 로그 관리
                    </a>
                </div>
            </div>
        </div>
        
        <div class="col-md-9">
            <div class="card shadow-sm mb-4">
                <div class="card-header bg-white d-flex justify-content-between align-items-center">
                    <h4 class="mb-0"><i class="fas fa-car me-2"></i>차량 이용 신청 상세보기</h4>
                    <div>
                        <a href="admin_vehicle_requests.asp" class="btn btn-secondary">
                            <i class="fas fa-arrow-left me-1"></i> 목록으로
                        </a>
                    </div>
                </div>
                <div class="card-body">
                    <div class="row mb-4">
                        <form method="post" action="admin_vehicle_request_process.asp">
                            <div class="col-md-6">
                                <h5 class="border-bottom pb-2 mb-3">신청 정보</h5>
                                <table class="table table-bordered">
                                    <tr>
                                        <th class="bg-light">신청일</th>
                                        <td><input type="date" name="request_date" value="<%= FormatDate(requestRS("request_date")) %>" class="form-control"></td>
                                    </tr>
                                    <tr>
                                        <th class="bg-light">결재 상태</th>
                                        <td><%= requestRS("approval_status") %></td>
                                    </tr>
                                    <tr>
                                        <th class="bg-light">이용 기간</th>
                                        <td>
                                            <input type="date" name="start_date" value="<%= FormatDate(requestRS("start_date")) %>" class="form-control"> ~
                                            <input type="date" name="end_date" value="<%= FormatDate(requestRS("end_date")) %>" class="form-control">
                                        </td>
                                    </tr>
                                    <tr>
                                        <th class="bg-light">제목</th>
                                        <td><input type="text" name="title" value="<%= IIf(IsNull(requestRS("title")), "", requestRS("title")) %>" class="form-control"></td>
                                    </tr>
                                    <tr>
                                        <th class="bg-light">목적지</th>
                                        <td><input type="text" name="destination" value="<%= IIf(IsNull(requestRS("destination")), "", requestRS("destination")) %>" class="form-control"></td>
                                    </tr>
                                    <tr>
                                        <th class="bg-light">사용 목적</th>
                                        <td><input type="text" name="purpose" value="<%= IIf(IsNull(requestRS("purpose")), "", requestRS("purpose")) %>" class="form-control"></td>
                                    </tr>
                                </table>
                                <input type="hidden" name="request_id" value="<%= requestRS("request_id") %>">
                                <button type="submit" class="btn btn-primary mt-3">수정</button>
                            </div>
                        </form>
                        
                        <div class="col-md-6">
                            <h5 class="border-bottom pb-2 mb-3">신청자 정보</h5>
                            <table class="table table-bordered">
                                <tr>
                                    <th class="bg-light" width="30%">이름</th>
                                    <td><%= IIf(IsNull(requestRS("user_id")), "-", requestRS("user_id")) %></td>
                                </tr>
                                <tr>
                                    <th class="bg-light">이메일</th>
                                    <td><%= IIf(IsNull(requestRS("user_email")), "-", requestRS("user_email")) %></td>
                                </tr>
                                
                                <tr>
                                    <th class="bg-light">부서</th>
                                    <td><%= IIf(IsNull(requestRS("department_name")), "-", requestRS("department_name")) %></td>
                                </tr>
                            </table>
                            
                            
                        </div>
                    </div>
                    
                </div>
            </div>
        </div>
    </div>
</div>

<%
' 사용한 객체 해제
If Not requestRS Is Nothing Then
    If requestRS.State = 1 Then
        requestRS.Close
    End If
    Set requestRS = Nothing
End If
%>

<!--#include file="../../includes/footer.asp"--> 