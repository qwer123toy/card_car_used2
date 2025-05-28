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

' 사용자 정보 조회
Dim userSQL, userRS
userSQL = "SELECT name, department_id, job_grade FROM " & dbSchema & ".Users WHERE user_id = '" & Session("user_id") & "'"
Set userRS = db99.Execute(userSQL)

Dim userName
If Not userRS.EOF Then
    userName = userRS("name")
Else
    userName = Session("user_id")
End If

' 부서명 가져오기
Function GetDepartmentName(deptId)
    If IsNull(deptId) Or deptId = "" Then
        GetDepartmentName = "-"
        Exit Function
    End If
    
    Dim deptName, deptSQL, deptRS
    deptSQL = "SELECT name FROM " & dbSchema & ".Department WHERE department_id = " & deptId
    
    On Error Resume Next
    Set deptRS = db99.Execute(deptSQL)
    
    If Err.Number = 0 And Not deptRS.EOF Then
        deptName = deptRS("name")
    Else
        deptName = deptId
    End If
    
    If Not deptRS Is Nothing Then
        If deptRS.State = 1 Then
            deptRS.Close
        End If
        Set deptRS = Nothing
    End If
    
    GetDepartmentName = deptName
End Function

' 직급명 가져오기
Function GetJobGradeName(gradeId)
    If IsNull(gradeId) Or gradeId = "" Then
        GetJobGradeName = "-"
        Exit Function
    End If
    
    Dim gradeName, gradeSQL, gradeRS
    gradeSQL = "SELECT name FROM " & dbSchema & ".Job_Grade WHERE job_grade_id = " & gradeId
    
    On Error Resume Next
    Set gradeRS = db99.Execute(gradeSQL)
    
    If Err.Number = 0 And Not gradeRS.EOF Then
        gradeName = gradeRS("name")
    Else
        gradeName = gradeId
    End If
    
    If Not gradeRS Is Nothing Then
        If gradeRS.State = 1 Then
            gradeRS.Close
        End If
        Set gradeRS = Nothing
    End If
    
    GetJobGradeName = gradeName
End Function

' 통계 정보 가져오기
Dim statSQL, statRS
Dim userCount, cardCount, vehicleCount, approvalCount

' 사용자 수
statSQL = "SELECT COUNT(*) AS cnt FROM " & dbSchema & ".Users WHERE is_active = 1"
Set statRS = db99.Execute(statSQL)
If Not statRS.EOF Then
    userCount = statRS("cnt")
Else
    userCount = 0
End If
Set statRS = Nothing

' 카드 계정 수
statSQL = "SELECT COUNT(*) AS cnt FROM " & dbSchema & ".CardAccount"
Set statRS = db99.Execute(statSQL)
If Not statRS.EOF Then
    cardCount = statRS("cnt")
Else
    cardCount = 0
End If
Set statRS = Nothing

' 차량 신청 수 (최근 30일)
statSQL = "SELECT COUNT(*) AS cnt FROM " & dbSchema & ".VehicleRequests WHERE request_date >= DATEADD(day, -30, GETDATE())"
Set statRS = db99.Execute(statSQL)
If Not statRS.EOF Then
    vehicleCount = statRS("cnt")
Else
    vehicleCount = 0
End If
Set statRS = Nothing

' 결재 수 (최근 30일)
statSQL = "SELECT COUNT(*) AS cnt FROM " & dbSchema & ".ApprovalLogs WHERE created_at >= DATEADD(day, -30, GETDATE())"
Set statRS = db99.Execute(statSQL)
If Not statRS.EOF Then
    approvalCount = statRS("cnt")
Else
    approvalCount = 0
End If
Set statRS = Nothing
%>

<!--#include file="../../includes/header.asp"-->

<div class="container-fluid my-4">
    <div class="row">
        <div class="col-md-3">            
            <div class="card shadow-sm mb-4">
                <div class="card-header bg-primary text-white">
                    <h5 class="mb-0"><i class="fas fa-cog me-2"></i>관리 메뉴</h5>
                </div>
                <div class="list-group list-group-flush">
                    <a href="admin_dashboard.asp" class="list-group-item list-group-item-action active">
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
                    <a href="admin_vehicle_requests.asp" class="list-group-item list-group-item-action">
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
                <div class="card-header bg-white">
                    <h4 class="mb-0"><i class="fas fa-tachometer-alt me-2"></i>관리자 대시보드</h4>
                </div>
                <div class="card-body">
                    <div class="alert alert-info" role="alert">
                        <i class="fas fa-info-circle me-2"></i>관리자 대시보드에 오신 것을 환영합니다. 여기에서 시스템의 주요 설정을 관리할 수 있습니다.
                    </div>
                    
                    <div class="row">
                        <div class="col-md-3 mb-4">
                            <div class="card bg-primary text-white">
                                <div class="card-body">
                                    <div class="d-flex justify-content-between align-items-center">
                                        <div>
                                            <h6 class="mb-0">등록 사용자</h6>
                                            <h2 class="mb-0"><%= userCount %></h2>
                                        </div>
                                        <div>
                                            <i class="fas fa-users fa-2x"></i>
                                        </div>
                                    </div>
                                </div>
                                <div class="card-footer d-flex justify-content-between align-items-center">
                                    <span>사용자 관리</span>
                                    <a href="admin_users.asp" class="text-white"><i class="fas fa-arrow-right"></i></a>
                                </div>
                            </div>
                        </div>
                        
                        <div class="col-md-3 mb-4">
                            <div class="card bg-success text-white">
                                <div class="card-body">
                                    <div class="d-flex justify-content-between align-items-center">
                                        <div>
                                            <h6 class="mb-0">카드 계정</h6>
                                            <h2 class="mb-0"><%= cardCount %></h2>
                                        </div>
                                        <div>
                                            <i class="fas fa-credit-card fa-2x"></i>
                                        </div>
                                    </div>
                                </div>
                                <div class="card-footer d-flex justify-content-between align-items-center">
                                    <span>카드 관리</span>
                                    <a href="admin_cardaccount.asp" class="text-white"><i class="fas fa-arrow-right"></i></a>
                                </div>
                            </div>
                        </div>
                        
                        <div class="col-md-3 mb-4">
                            <div class="card bg-warning text-white">
                                <div class="card-body">
                                    <div class="d-flex justify-content-between align-items-center">
                                        <div>
                                            <h6 class="mb-0">차량 신청</h6>
                                            <h2 class="mb-0"><%= vehicleCount %></h2>
                                        </div>
                                        <div>
                                            <i class="fas fa-car fa-2x"></i>
                                        </div>
                                    </div>
                                </div>
                                <div class="card-footer d-flex justify-content-between align-items-center">
                                    <span>최근 30일</span>
                                    <a href="admin_vehicle_requests.asp" class="text-white"><i class="fas fa-arrow-right"></i></a>
                                </div>
                            </div>
                        </div>
                        
                        <div class="col-md-3 mb-4">
                            <div class="card bg-info text-white">
                                <div class="card-body">
                                    <div class="d-flex justify-content-between align-items-center">
                                        <div>
                                            <h6 class="mb-0">결재 처리</h6>
                                            <h2 class="mb-0"><%= approvalCount %></h2>
                                        </div>
                                        <div>
                                            <i class="fas fa-file-signature fa-2x"></i>
                                        </div>
                                    </div>
                                </div>
                                <div class="card-footer d-flex justify-content-between align-items-center">
                                    <span>최근 30일</span>
                                    <a href="admin_approvals.asp" class="text-white"><i class="fas fa-arrow-right"></i></a>
                                </div>
                            </div>
                        </div>
                    </div>
         
                </div>
            </div>
        </div>
    </div>
</div>

<%
' 사용한 객체 해제
If Not userRS Is Nothing Then
    If userRS.State = 1 Then
        userRS.Close
    End If
    Set userRS = Nothing
End If
%>

<!--#include file="../../includes/footer.asp"--> 