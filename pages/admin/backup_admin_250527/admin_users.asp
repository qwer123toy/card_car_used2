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

' 사용자 삭제 처리
If Request.QueryString("action") = "delete" And Request.QueryString("id") <> "" Then
    Dim deleteId
    deleteId = PreventSQLInjection(Request.QueryString("id"))
    
    ' 현재 로그인된 사용자는 삭제할 수 없음
    If deleteId = Session("user_id") Then
        Response.Write("<script>alert('현재 로그인된 계정은 삭제할 수 없습니다.'); window.location.href='admin_users.asp';</script>")
        Response.End
    End If
    
    ' 삭제 쿼리 실행 (실제로는 비활성화)
    Dim deleteSQL
    deleteSQL = "UPDATE " & dbSchema & ".Users SET is_active = 0 WHERE user_id = '" & deleteId & "'"
    
    On Error Resume Next
    db.Execute(deleteSQL)
    
    If Err.Number <> 0 Then
        Response.Write("<script>alert('사용자 비활성화 중 오류가 발생했습니다: " & Server.HTMLEncode(Err.Description) & "'); window.location.href='admin_users.asp';</script>")
    Else
        ' 활동 로그 기록
        LogActivity Session("user_id"), "사용자비활성화", "사용자 비활성화 (ID: " & deleteId & ")"
        Response.Write("<script>alert('사용자가 비활성화되었습니다.'); window.location.href='admin_users.asp';</script>")
    End If
    On Error GoTo 0
    Response.End
End If

' 사용자 활성화 처리
If Request.QueryString("action") = "active" And Request.QueryString("id") <> "" Then
    Dim activeId
    activeId = PreventSQLInjection(Request.QueryString("id"))
    
    ' 활성화 쿼리 실행
    Dim activeSQL
    activeSQL = "UPDATE " & dbSchema & ".Users SET is_active = 1 WHERE user_id = '" & activeId & "'"
    
    On Error Resume Next
    db99.Execute(activeSQL)
    
    If Err.Number <> 0 Then
        Response.Write("<script>alert('사용자 활성화 중 오류가 발생했습니다: " & Server.HTMLEncode(Err.Description) & "'); window.location.href='admin_users.asp';</script>")
    Else
        ' 활동 로그 기록
        LogActivity Session("user_id"), "사용자활성화", "사용자 활성화 (ID: " & activeId & ")"
        Response.Write("<script>alert('사용자가 활성화되었습니다.'); window.location.href='admin_users.asp';</script>")
    End If
    On Error GoTo 0
    Response.End
End If

' 페이징 처리
Dim pageNo, pageSize, totalCount, totalPages
pageSize = 10 ' 페이지당 표시할 레코드 수

' 현재 페이지 번호
If Request.QueryString("page") = "" Then
    pageNo = 1
Else
    pageNo = CInt(Request.QueryString("page"))
End If

' 검색 조건에 따른 SQL 쿼리 구성
Dim searchKeyword, searchField, whereClause
searchKeyword = Trim(Request.QueryString("keyword"))
searchField = Request.QueryString("field")

whereClause = ""
If searchKeyword <> "" Then
    If searchField = "user_id" Then
        whereClause = " WHERE user_id LIKE '%" & PreventSQLInjection(searchKeyword) & "%'"
    ElseIf searchField = "name" Then
        whereClause = " WHERE name LIKE '%" & PreventSQLInjection(searchKeyword) & "%'"
    ElseIf searchField = "email" Then
        whereClause = " WHERE email LIKE '%" & PreventSQLInjection(searchKeyword) & "%'"
    ElseIf searchField = "department" Then
        ' 부서명으로 검색
        whereClause = " WHERE department_id IN (SELECT department_id FROM " & dbSchema & ".Department WHERE name LIKE '%" & PreventSQLInjection(searchKeyword) & "%')"
    End If
End If

' 전체 레코드 수
Dim countSQL, countRS
countSQL = "SELECT COUNT(*) AS cnt FROM " & dbSchema & ".Users" & whereClause
Set countRS = db99.Execute(countSQL)
totalCount = countRS("cnt")
totalPages = totalCount / pageSize

' 사용자 목록 조회
Dim listSQL, listRS
listSQL = "SELECT * FROM (" & _
          "SELECT TOP " & pageSize & " * FROM (" & _
          "SELECT TOP " & (pageNo * pageSize) & " U.user_id, U.name, U.email, U.department_id, U.is_active, J.name AS job_grade " & _
          "FROM " & dbSchema & ".Users U " & _
          "JOIN " & dbSchema & ".Job_Grade J ON U.job_grade = J.job_grade_id " & _
          "WHERE U.user_id <> 'admin'"

If whereClause <> "" Then
    listSQL = listSQL & " AND " & Mid(whereClause, 8) ' " WHERE " 제거하고 조건만 붙이기
End If

listSQL = listSQL & " ORDER BY U.user_id DESC) AS T1 " & _
          "ORDER BY user_id ASC) AS T2 " & _
          "ORDER BY user_id DESC"

          
Set listRS = db99.Execute(listSQL)

' 부서 목록 조회 (등록/수정용)
Dim deptSQL, deptRS
deptSQL = "SELECT department_id, name FROM " & dbSchema & ".Department ORDER BY name"
Set deptRS = db99.Execute(deptSQL)

' 직급 목록 조회 (등록/수정용)
Dim gradeSQL, gradeRS
gradeSQL = "SELECT job_grade_id, name FROM " & dbSchema & ".Job_Grade ORDER BY sort_order"
Set gradeRS = db99.Execute(gradeSQL)

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
                    <h4 class="mb-0"><i class="fas fa-users me-2"></i>사용자 관리</h4>
                    <button class="btn btn-primary" data-bs-toggle="modal" data-bs-target="#addUserModal">
                        <i class="fas fa-plus me-1"></i> 사용자 등록
                    </button>
                </div>
                <div class="card-body">
                    <!-- 검색 폼 -->
                    <form action="admin_users.asp" method="get" class="mb-4">
                        <div class="row g-2">
                            <div class="col-md-3">
                                <select class="form-select" name="field">
                                    <option value="user_id" <% If searchField = "user_id" Then %>selected<% End If %>>사용자 ID</option>
                                    <option value="name" <% If searchField = "name" Then %>selected<% End If %>>이름</option>
                                    <option value="email" <% If searchField = "email" Then %>selected<% End If %>>이메일</option>
                                    <option value="department" <% If searchField = "department" Then %>selected<% End If %>>부서</option>
                                </select>
                            </div>
                            <div class="col-md-7">
                                <input type="text" class="form-control" name="keyword" value="<%= searchKeyword %>" placeholder="검색어를 입력하세요">
                            </div>
                            <div class="col-md-2">
                                <button type="submit" class="btn btn-primary w-100">검색</button>
                            </div>
                        </div>
                    </form>
                    
                    <!-- 사용자 목록 -->
                    <div class="table-responsive">
                        <table class="table table-striped table-bordered table-hover">
                            <thead class="table-dark">
                                <tr>
                                    <th>ID</th>
                                    <th>이름</th>
                                    <th>이메일</th>
                                    <th>부서</th>
                                    <th>직급</th>
                                    <th>상태</th>
                                    <th>관리</th>
                                </tr>
                            </thead>
                            <tbody>
                                <% 
                                If listRS.EOF Then 
                                %>
                                <tr>
                                    <td colspan="8" class="text-center">등록된 사용자가 없습니다.</td>
                                </tr>
                                <% 
                                Else
                                    Do While Not listRS.EOF 
                                %>
                                <tr>
                                    <td><%= listRS("user_id") %></td>
                                    <td><%= listRS("name") %></td>
                                    <td><%= listRS("email") %></td>
                                    <td><%= GetDepartmentName(listRS("department_id")) %></td>
                                    <td><%= listRS("job_grade") %></td>
                                    <td><%= IIf(listRS("is_active") = 1, "활성", "비활성") %></td>
                                    <td>
                                        <a href="admin_user_view.asp?id=<%= listRS("user_id") %>" class="btn btn-sm btn-primary">
                                            <i class="fas fa-eye"></i> 상세보기
                                        </a>
                                        
                                        <% if listRS("is_active") = 1 then %>
                                            <button class="btn btn-sm btn-danger" onclick="confirmDelete('<%= listRS("user_id") %>')">
                                                <i class="fas fa-trash">삭제</i>
                                            </button>
                                        <% else %>
                                            <button class="btn btn-sm btn-success" onclick="confirmActive('<%= listRS("user_id") %>')">
                                                <i class="fas fa-trash">활성화</i>
                                            </button>
                                        <% end if %>
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
                                <a class="page-link" href="admin_users.asp?page=<%= pageNo - 1 %>&field=<%= searchField %>&keyword=<%= searchKeyword %>">이전</a>
                            </li>
                            <% End If %>
                            
                            <% 
                            Dim startPage, endPage
                            startPage = Max(1, pageNo - 5)
                            endPage = Min(totalPages, pageNo + 5)
                            
                            For i = startPage To endPage
                            %>
                            <li class="page-item <% If i = pageNo Then %>active<% End If %>">
                                <a class="page-link" href="admin_users.asp?page=<%= i %>&field=<%= searchField %>&keyword=<%= searchKeyword %>"><%= i %></a>
                            </li>
                            <% Next %>
                            
                            <% If pageNo < totalPages Then %>
                            <li class="page-item">
                                <a class="page-link" href="admin_users.asp?page=<%= pageNo + 1 %>&field=<%= searchField %>&keyword=<%= searchKeyword %>">다음</a>
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
    if (confirm("정말로 이 사용자를 비활성화하시겠습니까? 이 작업은 계정을 비활성화하며 실제로 삭제하지는 않습니다.")) {
        window.location.href = "admin_users.asp?action=delete&id=" + id;
    }
}

function confirmActive(id) {
    if (confirm("정말로 이 사용자를 활성화하시겠습니까?")) {
        window.location.href = "admin_users.asp?action=active&id=" + id;
    }
}

</script>

<%
' 사용한 객체 해제
If Not listRS Is Nothing Then
    If listRS.State = 1 Then
        listRS.Close
    End If
    Set listRS = Nothing
End If

If Not deptRS Is Nothing Then
    If deptRS.State = 1 Then
        deptRS.Close
    End If
    Set deptRS = Nothing
End If

If Not gradeRS Is Nothing Then
    If gradeRS.State = 1 Then
        gradeRS.Close
    End If
    Set gradeRS = Nothing
End If
%>

<!--#include file="../../includes/footer.asp"--> 