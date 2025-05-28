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

' 결재 문서 삭제 처리
If Request.QueryString("action") = "delete" And Request.QueryString("id") <> "" Then
    Dim deleteId
    deleteId = PreventSQLInjection(Request.QueryString("id"))
    
    ' 삭제 쿼리 실행
    Dim deleteSQL
    deleteSQL = "DELETE FROM " & dbSchema & ".Approvals WHERE approval_id = " & deleteId
    
    On Error Resume Next
    db.Execute(deleteSQL)
    
    If Err.Number <> 0 Then
        Response.Write("<script>alert('결재 문서 삭제 중 오류가 발생했습니다: " & Server.HTMLEncode(Err.Description) & "'); window.location.href='admin_approvals.asp';</script>")
    Else
        ' 활동 로그 기록
        LogActivity Session("user_id"), "결재문서삭제", "결재 문서 삭제 (ID: " & deleteId & ")"
        Response.Write("<script>alert('결재 문서가 삭제되었습니다.'); window.location.href='admin_approvals.asp';</script>")
    End If
    On Error GoTo 0
    Response.End
End If

' 페이징 처리
Dim pageNo, pageSize, totalCount, totalPages
pageSize = 15 ' 페이지당 표시할 레코드 수

' 현재 페이지 번호
If Request.QueryString("page") = "" Then
    pageNo = 1
Else
    pageNo = CInt(Request.QueryString("page"))
End If

' 검색 조건에 따른 SQL 쿼리 구성
Dim searchKeyword, searchField, searchDateFrom, searchDateTo, searchStatus, whereClause
searchKeyword = Trim(Request.QueryString("keyword"))
searchField = Request.QueryString("field")
searchDateFrom = Request.QueryString("date_from")
searchDateTo = Request.QueryString("date_to")
searchStatus = Request.QueryString("status")

whereClause = ""
Dim whereConditions : whereConditions = Array()
Dim conditionIndex : conditionIndex = 0

' 키워드 검색 조건
If searchKeyword <> "" Then
    If searchField = "requester_id" Then
        ReDim Preserve whereConditions(conditionIndex)
        whereConditions(conditionIndex) = "u1.name LIKE '%" & PreventSQLInjection(searchKeyword) & "%'"
        conditionIndex = conditionIndex + 1
    ElseIf searchField = "approver_id" Then
        ReDim Preserve whereConditions(conditionIndex)
        whereConditions(conditionIndex) = "u2.name LIKE '%" & PreventSQLInjection(searchKeyword) & "%'"
        conditionIndex = conditionIndex + 1
    ElseIf searchField = "title" Then
        ReDim Preserve whereConditions(conditionIndex)
        whereConditions(conditionIndex) = "a.title LIKE '%" & PreventSQLInjection(searchKeyword) & "%'"
        conditionIndex = conditionIndex + 1
    End If
End If

' 날짜 범위 검색 조건
If IsDate(searchDateFrom) Then
    ReDim Preserve whereConditions(conditionIndex)
    whereConditions(conditionIndex) = "a.request_date >= '" & CDate(searchDateFrom) & "'"
    conditionIndex = conditionIndex + 1
End If

If IsDate(searchDateTo) Then
    ReDim Preserve whereConditions(conditionIndex)
    whereConditions(conditionIndex) = "a.request_date <= '" & CDate(searchDateTo) & " 23:59:59'"
    conditionIndex = conditionIndex + 1
End If

' 상태 검색 조건
If searchStatus <> "" Then
    ReDim Preserve whereConditions(conditionIndex)
    whereConditions(conditionIndex) = "a.status = '" & PreventSQLInjection(searchStatus) & "'"
    conditionIndex = conditionIndex + 1
End If

' WHERE 절 구성
If conditionIndex > 0 Then
    whereClause = " WHERE " & Join(whereConditions, " AND ")
End If

' 전체 레코드 수
Dim countSQL, countRS
countSQL = "SELECT COUNT(*) AS cnt " & _
           "FROM " & dbSchema & ".Approvallogs a " & _
           "LEFT JOIN " & dbSchema & ".Users u1 ON a.approver_id = u1.user_id " & _
           "LEFT JOIN " & dbSchema & ".Users u2 ON a.approver_id = u2.user_id " & _
           IIf(whereClause <> "", whereClause, "")

Set countRS = db99.Execute(countSQL)
totalCount = countRS("cnt")
totalPages = (totalCount + pageSize - 1) \ pageSize

' 결재 문서 목록 조회
Dim listSQL, listRS
' WHERE 절이 있을 경우 "AND"로 연결 (앞의 WHERE 제거하고)
listSQL = "SELECT * FROM (" & _
          "SELECT TOP " & pageSize & " * FROM (" & _
          "SELECT TOP " & (pageNo * pageSize) & " a.approval_log_id, a.approver_id, " & _
          "a.target_table_name, a.target_id, a.approval_step, a.status, a.approved_at, " & _
          "a.comments, a.created_at, u1.name AS approver_name, " & _
          "ISNULL(cu.title, vr.title) AS title " & _
          "FROM " & dbSchema & ".Approvallogs a " & _
          "INNER JOIN (" & _
          "    SELECT MAX(approval_log_id) AS approval_log_id " & _
          "    FROM " & dbSchema & ".Approvallogs " & _
          "    GROUP BY target_id" & _
          ") AS grouped ON a.approval_log_id = grouped.approval_log_id " & _
          "LEFT JOIN " & dbSchema & ".Users u1 ON a.approver_id = u1.user_id " & _
          "LEFT JOIN " & dbSchema & ".CardUsage cu ON a.target_table_name = 'CardUsage' AND a.target_id = cu.usage_id " & _
          "LEFT JOIN " & dbSchema & ".VehicleRequests vr ON a.target_table_name = 'VehicleRequests' AND a.target_id = vr.request_id " & _
          IIf(whereClause <> "", whereClause, "") & " " & _
          "ORDER BY a.created_at DESC) AS T1 " & _
          "ORDER BY created_at ASC) AS T2 " & _
          "ORDER BY created_at DESC"


Set listRS = db99.Execute(listSQL)

' 상태 옵션
Dim statusOptions : statusOptions = Array("대기", "승인", "반려")

' 사용자 목록 조회
Dim userSQL, userRS
userSQL = "SELECT user_id, name FROM " & dbSchema & ".Users WHERE is_active = 1 ORDER BY name"
Set userRS = db99.Execute(userSQL)

' 날짜 포맷
Function FormatDate(dateValue)
    If IsNull(dateValue) Or Not IsDate(dateValue) Then
        FormatDate = "-"
    Else
        FormatDate = FormatDateTime(dateValue, 2)
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
                    <a href="admin_vehicle_requests.asp" class="list-group-item list-group-item-action">
                        <i class="fas fa-car me-2"></i>차량 이용 신청 관리
                    </a>
                    <a href="admin_approvals.asp" class="list-group-item list-group-item-action active">
                        <i class="fas fa-file-signature me-2"></i>결재 로그 관리
                    </a>
                </div>
            </div>
        </div>
        
        <div class="col-md-9">
            <div class="card shadow-sm mb-4">
                <div class="card-header bg-white d-flex justify-content-between align-items-center">
                    <h4 class="mb-0"><i class="fas fa-file-signature me-2"></i>결재 문서 관리</h4>
                </div>
                <div class="card-body">
                    <!-- 검색 폼 -->
                    <form action="admin_approvals.asp" method="get" class="mb-4">
                        <div class="row g-2 mb-2">
                            <div class="col-md-3">
                                <select name="field" class="form-select">
                                    <option value="requester_id" <% If searchField = "requester_id" Then Response.Write("selected") %>>신청자</option>
                                    <option value="approver_id" <% If searchField = "approver_id" Then Response.Write("selected") %>>결재자</option>
                                    <option value="title" <% If searchField = "title" Then Response.Write("selected") %>>제목</option>
                                </select>
                            </div>
                            <div class="col-md-4">
                                <input type="text" class="form-control" name="keyword" value="<%= searchKeyword %>" placeholder="검색어를 입력하세요">
                            </div>
                            <div class="col-md-2">
                                <button type="submit" class="btn btn-primary w-100">검색</button>
                            </div>
                        </div>
                        <div class="row g-2 mb-2">
                            <div class="col-md-5">
                                <div class="input-group">
                                    <span class="input-group-text">시작일</span>
                                    <input type="date" class="form-control" name="date_from" value="<%= searchDateFrom %>">
                                </div>
                            </div>
                            <div class="col-md-5">
                                <div class="input-group">
                                    <span class="input-group-text">종료일</span>
                                    <input type="date" class="form-control" name="date_to" value="<%= searchDateTo %>">
                                </div>
                            </div>
                            <div class="col-md-2">
                                <button type="button" class="btn btn-secondary w-100" onclick="clearSearch()">초기화</button>
                            </div>
                        </div>
                        <div class="row g-2">
                            <div class="col-md-10">
                                <select name="status" class="form-select">
                                    <option value="">모든 상태</option>
                                    <% For Each statusOption In statusOptions %>
                                    <option value="<%= statusOption %>" <% If searchStatus = statusOption Then Response.Write("selected") %>><%= statusOption %></option>
                                    <% Next %>
                                </select>
                            </div>
                        </div>
                    </form>

                    <!-- 결재 문서 목록 -->
                    <div class="table-responsive">
                        <table class="table table-striped table-bordered table-hover">
                            <thead class="table-dark">
                                <tr>
                                    <th>결재 번호</th>
                                    <th>신청일</th>
                                    <th>제목</th>
                                    <th>문서 종류</th>
                                    <th>최종 결재자</th>
                                    <th>상태</th>
                                    <th>관리</th>
                                </tr>
                            </thead>
                            <tbody>
                                <% 
                                If listRS.EOF Then 
                                %>
                                <tr>
                                    <td colspan="8" class="text-center">등록된 결재 문서가 없습니다.</td>
                                </tr>
                                <% 
                                Else
                                    Do While Not listRS.EOF 
                                %>
                                <tr>
                                    <td><%= listRS("approval_log_id") %></td>
                                    <td><%= FormatDate(listRS("created_at")) %></td>
                                    <td><%= listRS("title") %></td>
                                    <td><%= listRS("target_table_name") %></td>
                                    <td><%= IIf(IsNull(listRS("approver_name")), "-", listRS("approver_name")) %></td>
                                    <td><%= listRS("status") %></td>
                                    <td>
                                        <a href="admin_approval_view.asp?target_id=<%= listRS("target_id") %>&target_table_name=<%= listRS("target_table_name") %>" class="btn btn-sm btn-primary view-approval">
                                            <i class="fas fa-eye"></i> 상세보기
                                        </a>
                                        <button class="btn btn-sm btn-danger" onclick="confirmDelete('<%= listRS("approval_log_id") %>')">
                                            <i class="fas fa-trash"></i> 삭제
                                        </button>
                                    </td>
                                </tr>
                                <% 
                                        listRS.MoveNext
                                    Loop
                                End If
                                %>
                            </tbody>
                        </table>
                    </div>
                    
                    <!-- 페이징 -->
                    <% If totalPages > 1 Then %>
                    <nav aria-label="Page navigation">
                        <ul class="pagination justify-content-center">
                            <% If pageNo > 1 Then %>
                            <li class="page-item">
                                <a class="page-link" href="admin_approvals.asp?page=<%= pageNo - 1 %>&field=<%= searchField %>&keyword=<%= searchKeyword %>&date_from=<%= searchDateFrom %>&date_to=<%= searchDateTo %>&status=<%= searchStatus %>">이전</a>
                            </li>
                            <% End If %>
                            
                            <%
                            ' 두 숫자 중 큰 값을 반환
                            Function Max(a, b)
                                If a > b Then
                                    Max = a
                                Else
                                    Max = b
                                End If
                            End Function

                            ' 두 숫자 중 작은 값을 반환
                            Function Min(a, b)
                                If a < b Then
                                    Min = a
                                Else
                                    Min = b
                                End If
                            End Function
                            
                            Dim startPage, endPage
                            startPage = Max(1, pageNo - 5)
                            endPage = Min(totalPages, pageNo + 5)
                            
                            For i = startPage To endPage
                            %>
                            <li class="page-item <% If i = pageNo Then %>active<% End If %>">
                                <a class="page-link" href="admin_approvals.asp?page=<%= i %>&field=<%= searchField %>&keyword=<%= searchKeyword %>&date_from=<%= searchDateFrom %>&date_to=<%= searchDateTo %>&status=<%= searchStatus %>"><%= i %></a>
                            </li>
                            <% Next %>
                            
                            <% If pageNo < totalPages Then %>
                            <li class="page-item">
                                <a class="page-link" href="admin_approvals.asp?page=<%= pageNo + 1 %>&field=<%= searchField %>&keyword=<%= searchKeyword %>&date_from=<%= searchDateFrom %>&date_to=<%= searchDateTo %>&status=<%= searchStatus %>">다음</a>
                            </li>
                            <% End If %>
                        </ul>
                    </nav>
                    <% End If %>
                </div>
            </div>
        </div>
    </div>
</div>

<script>
// 삭제 확인
function confirmDelete(id) {
    if (confirm("정말로 이 결재 문서를 삭제하시겠습니까? 이 작업은 되돌릴 수 없습니다.")) {
        window.location.href = "admin_approvals.asp?action=delete&id=" + id;
    }
}

// 검색 초기화
function clearSearch() {
    window.location.href = "admin_approvals.asp";
}
</script>

<%
' 날짜를 input type="date"에 사용할 수 있는 형식으로 변환
Function FormatDateForInput(dateValue)
    If IsDate(dateValue) Then
        FormatDateForInput = Year(dateValue) & "-" & Right("0" & Month(dateValue), 2) & "-" & Right("0" & Day(dateValue), 2)
    Else
        FormatDateForInput = ""
    End If
End Function

' 사용한 객체 해제
If Not listRS Is Nothing Then
    If listRS.State = 1 Then
        listRS.Close
    End If
    Set listRS = Nothing
End If

If Not userRS Is Nothing Then
    If userRS.State = 1 Then
        userRS.Close
    End If
    Set userRS = Nothing
End If
%>

<!--#include file="../../includes/footer.asp"--> 