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

' 사용자 ID 확인
Dim userId
userId = Request.QueryString("id")

If userId = "" Then
    Response.Write("<script>alert('잘못된 접근입니다.'); window.location.href='admin_users.asp';</script>")
    Response.End
End If

' 사용자 정보 조회
Dim userSQL, userRS
userSQL = "SELECT u.*, d.name as department_name, j.name as grade_name " & _
          "FROM " & dbSchema & ".Users u " & _
          "LEFT JOIN " & dbSchema & ".Department d ON u.department_id = d.department_id " & _
          "LEFT JOIN " & dbSchema & ".Job_Grade j ON u.job_grade = j.job_grade_id " & _
          "WHERE u.user_id = '" & PreventSQLInjection(userId) & "'"


Set userRS = db99.Execute(userSQL)

' 데이터가 없으면 목록으로 리다이렉션
If userRS.EOF Then
    Response.Write("<script>alert('해당 사용자 정보를 찾을 수 없습니다.'); window.location.href='admin_users.asp';</script>")
    Response.End
End If

' 부서 목록 조회
Dim deptSQL, deptRS
deptSQL = "SELECT department_id, name FROM " & dbSchema & ".Department ORDER BY name"
Set deptRS = db99.Execute(deptSQL)

' 직급 목록 조회
Dim gradeSQL, gradeRS
gradeSQL = "SELECT job_grade_id, name FROM " & dbSchema & ".Job_Grade ORDER BY sort_order"
Set gradeRS = db99.Execute(gradeSQL)

' 사용자 카드 사용 내역 조회
Dim cardUsageSQL, cardUsageRS
cardUsageSQL = "SELECT TOP 5 cu.usage_id, cu.usage_date, cu.amount, cu.store_name, ca.account_name AS card_name " & _
               "FROM " & dbSchema & ".CardUsage cu " & _
               "LEFT JOIN " & dbSchema & ".CardAccount ca ON cu.card_id = ca.card_id " & _
               "WHERE cu.user_id = '" & PreventSQLInjection(userId) & "' " & _
               "ORDER BY cu.usage_date DESC"

Set cardUsageRS = db99.Execute(cardUsageSQL)

' 사용자 차량 이용 신청 조회
Dim vehicleReqSQL, vehicleReqRS
vehicleReqSQL = "SELECT TOP 5 vr.request_id, vr.request_date, vr.start_date, vr.end_date " & _
                "FROM " & dbSchema & ".VehicleRequests vr " & _
                "WHERE vr.user_id = '" & PreventSQLInjection(userId) & "' " & _
                "ORDER BY vr.request_date DESC"

Set vehicleReqRS = db99.Execute(vehicleReqSQL)

' 날짜 포맷
Function FormatDate(dateValue)
    If IsNull(dateValue) Or Not IsDate(dateValue) Then
        FormatDate = "-"
    Else
        FormatDate = FormatDateTime(dateValue, 2) & " " & FormatDateTime(dateValue, 4)
    End If
End Function

' 상태 표시
Function GetStatusBadge(status)
    Select Case status
        Case "대기중"
            GetStatusBadge = "<span class=""badge bg-warning"">대기중</span>"
        Case "승인됨"
            GetStatusBadge = "<span class=""badge bg-success"">승인됨</span>"
        Case "거부됨"
            GetStatusBadge = "<span class=""badge bg-danger"">거부됨</span>"
        Case "취소됨"
            GetStatusBadge = "<span class=""badge bg-secondary"">취소됨</span>"
        Case "완료됨"
            GetStatusBadge = "<span class=""badge bg-info"">완료됨</span>"
        Case Else
            GetStatusBadge = "<span class=""badge bg-light text-dark"">기타</span>"
    End Select
End Function


' 천 단위 쉼표 함수
Function AddComma(n)
    Dim s, l, i, result
    s = CStr(n)
    l = Len(s)
    result = ""
    
    For i = 1 To l
        result = Mid(s, l - i + 1, 1) & result
        If i Mod 3 = 0 And i <> l Then
            result = "," & result
        End If
    Next

    AddComma = result
End Function


' 활성 상태 표시
Function GetActiveStatus(isActive)
    If isActive Then
        GetActiveStatus = "<span class=""badge bg-success"">활성</span>"
    Else
        GetActiveStatus = "<span class=""badge bg-danger"">비활성</span>"
    End If
End Function

' 관리자 상태 표시
Function GetAdminStatus(isAdmin)
    If isAdmin Then
        GetAdminStatus = "<span class=""badge bg-primary"">관리자</span>"
    Else
        GetAdminStatus = "<span class=""badge bg-secondary"">일반 사용자</span>"
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
                    <a href="admin_users.asp" class="list-group-item list-group-item-action active">
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
                <div class="card-header bg-white d-flex justify-content-between align-items-center">
                    <h4 class="mb-0"><i class="fas fa-user me-2"></i>사용자 상세정보</h4>
                    <div>
                        <a href="admin_users.asp" class="btn btn-secondary">
                            <i class="fas fa-arrow-left me-1"></i> 목록으로
                        </a>
                    </div>
                  
                </div>
                <div class="card-body">
                    <div class="row mb-4">
                        <div class="col-md-12">
                            <form method="post" action="admin_users_process.asp">
                                <input type="hidden" name="user_id" value="<%= userRS("user_id") %>">
                                <h5 class="border-bottom pb-2 mb-3">기본 정보</h5>
                                <div class="row">
                                    <div class="col-md-6">
                                        <table class="table table-bordered">
                                            <tr>
                                                <th class="bg-light" width="30%">사용자 ID</th>
                                                <td><%= userRS("user_id") %></td>
                                            </tr>
                                            <tr>
                                                <th class="bg-light">이름</th>
                                                <td><input type="text" class="form-control" name="name" value="<%= userRS("name") %>"></td>
                                            </tr>
                                            <tr>
                                                <th class="bg-light">이메일</th>
                                                <td><input type="email" class="form-control" name="email" value="<%= userRS("email") %>"></td>
                                            </tr>
                                        </table>
                                    </div>
                                    <div class="col-md-6">
                                        <table class="table table-bordered">
                                            <tr>
                                                <th class="bg-light">부서</th>
                                                <td>
                                                    <select name="department_id" class="form-select">
                                                        <% 
                                                        Do While Not deptRS.EOF 
                                                            Dim selectedDept
                                                            selectedDept = ""
                                                            If userRS("department_id") = deptRS("department_id") Then selectedDept = "selected"
                                                        %>
                                                            <option value="<%= deptRS("department_id") %>" <%= selectedDept %>><%= deptRS("name") %></option>
                                                        <% 
                                                            deptRS.MoveNext 
                                                        Loop 
                                                        %>
                                                    </select>
                                                </td>
                                            </tr>
                                            
                                            <tr>
                                                <th class="bg-light">직급</th>
                                                <td>
                                                    <select name="job_grade_id" class="form-select">
                                                        <% 
                                                        Do While Not gradeRS.EOF 
                                                            Dim selectedGrade
                                                            selectedGrade = ""
                                                            If CInt(userRS("job_grade")) = CInt(gradeRS("job_grade_id")) Then selectedGrade = "selected"
                                                            %>  
                                                            <option value="<%= gradeRS("job_grade_id") %>" <%= selectedGrade %>><%= gradeRS("name") %></option>
                                                        <% 
                                                            gradeRS.MoveNext 
                                                        Loop 
                                                        %>
                                                    </select>
                                                </td>
                                            </tr>
                                            <tr>
                                                <th class="bg-light">상태</th>
                                                <td><%= GetActiveStatus(userRS("is_active")) %></td>
                                            </tr>
                                        </table>
                                    </div>
                                </div>
                                <button type="submit" class="btn btn-primary">
                                    <i class="fas fa-save me-1"></i> 수정
                                </button>
                            </form>
                        </div>
                    </div>
                    
                    <div class="row mb-4">
                        <div class="col-md-12">
                            <h5 class="border-bottom pb-2 mb-3">최근 카드 사용 내역</h5>
                            <div class="table-responsive">
                                <table class="table table-striped table-bordered table-hover">
                                    <thead class="table-dark">
                                        <tr>
                                            <th>ID</th>
                                            <th>사용일</th>
                                            <th>카드</th>
                                            <th>가맹점</th>
                                            <th>금액</th>
                                            <th>관리</th>
                                        </tr>
                                    </thead>
                                    <tbody>
                                        <% 
                                        If cardUsageRS.EOF Then 
                                        %>
                                        <tr>
                                            <td colspan="6" class="text-center">등록된 카드 사용 내역이 없습니다.</td>
                                        </tr>
                                        <% 
                                        Else
                                            Do While Not cardUsageRS.EOF 
                                        %>
                                        <tr>
                                            <td><%= cardUsageRS("usage_id") %></td>
                                            <td><%= FormatDateTime(cardUsageRS("usage_date"), 2) %></td>
                                            <td><%= IIf(IsNull(cardUsageRS("card_name")), "-", cardUsageRS("card_name")) %></td>
                                            <td><%= cardUsageRS("store_name") %></td>
                                            <td class="text-end"><%= FormatCurrency(cardUsageRS("amount")) %></td>
                                            <td>
                                                <a href="admin_card_usage_view.asp?id=<%= cardUsageRS("usage_id") %>" class="btn btn-sm btn-primary">
                                                    <i class="fas fa-eye"></i> 상세보기
                                                </a>
                                            </td>
                                        </tr>
                                        <% 
                                                cardUsageRS.MoveNext
                                            Loop
                                        End If
                                        %>
                                    </tbody>
                                </table>
                                <% If Not cardUsageRS.EOF Then %>
                                <div class="text-end">
                                    <a href="admin_card_usage.asp?field=user_id&keyword=<%= userRS("name") %>" class="btn btn-outline-primary">
                                        <i class="fas fa-list me-1"></i> 모든 카드 사용 내역 보기
                                    </a>
                                </div>
                                <% End If %>
                            </div>
                        </div>
                    </div>
                    
                    <div class="row mb-4">
                        <div class="col-md-12">
                            <h5 class="border-bottom pb-2 mb-3">최근 차량 이용 신청</h5>
                            <div class="table-responsive">
                                <table class="table table-striped table-bordered table-hover">
                                    <thead class="table-dark">
                                        <tr>
                                            <th>ID</th>
                                            <th>신청일</th>
                                            <th>사용기간</th>
                                            
                                            
                                            
                                            <th>관리</th>
                                        </tr>
                                    </thead>
                                    <tbody>
                                        <% 
                                        If vehicleReqRS.EOF Then 
                                        %>
                                        <tr>
                                            <td colspan="7" class="text-center">등록된 차량 이용 신청이 없습니다.</td>
                                        </tr>
                                        <% 
                                        Else
                                            Do While Not vehicleReqRS.EOF 
                                        %>
                                        <tr>
                                            <td><%= vehicleReqRS("request_id") %></td>
                                            <td><%= FormatDateTime(vehicleReqRS("request_date"), 2) %></td>
                                            <td>
                                                <%= FormatDateTime(vehicleReqRS("start_date"), 2) %> ~ 
                                                <%= FormatDateTime(vehicleReqRS("end_date"), 2) %>
                                            </td>
                                            
                                            
                                            <td>
                                                <a href="admin_vehicle_request_view.asp?id=<%= vehicleReqRS("request_id") %>" class="btn btn-sm btn-primary">
                                                    <i class="fas fa-eye"></i> 상세보기
                                                </a>
                                            </td>
                                        </tr>
                                        <% 
                                                vehicleReqRS.MoveNext
                                            Loop
                                        End If
                                        %>
                                    </tbody>
                                </table>
                                <% If Not vehicleReqRS.EOF Then %>
                                <div class="text-end">
                                    <a href="admin_vehicle_requests.asp?field=user_id&keyword=<%= userRS("name") %>" class="btn btn-outline-primary">
                                        <i class="fas fa-list me-1"></i> 모든 차량 이용 신청 보기
                                    </a>
                                </div>
                                <% End If %>
                            </div>
                        </div>
                    </div>

                    </div>
                    
                    <div class="row">
                        <div class="col-md-12">
                            <div class="d-flex justify-content-between">
                                <a href="admin_users.asp" class="btn btn-secondary">
                                    <i class="fas fa-arrow-left me-1"></i> 목록으로
                                </a>

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

If Not cardUsageRS Is Nothing Then
    Set activityRS = Nothing
End If

If Not cardUsageRS Is Nothing Then
    If cardUsageRS.State = 1 Then
        cardUsageRS.Close
    End If
    Set cardUsageRS = Nothing
End If

If Not vehicleReqRS Is Nothing Then
    If vehicleReqRS.State = 1 Then
        vehicleReqRS.Close
    End If
    Set vehicleReqRS = Nothing
End If
%>

<!--#include file="../../includes/footer.asp"--> 