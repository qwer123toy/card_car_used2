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

' 부서 삭제 처리
If Request.QueryString("action") = "delete" And Request.QueryString("id") <> "" Then
    Dim deleteId
    deleteId = PreventSQLInjection(Request.QueryString("id"))
    
    ' 삭제 전 해당 부서가 사용중인지 확인
    Dim checkUseSQL, checkUseRS
    checkUseSQL = "SELECT COUNT(*) AS cnt FROM " & dbSchema & ".Users WHERE department_id = " & deleteId
    Set checkUseRS = db.Execute(checkUseSQL)
    
    If Not checkUseRS.EOF And checkUseRS("cnt") > 0 Then
        Response.Write("<script>alert('이 부서는 사용자에게 할당되어 있어 삭제할 수 없습니다.'); window.location.href='admin_department.asp';</script>")
        Response.End
    End If
    
    ' 삭제 쿼리 실행
    Dim deleteSQL
    deleteSQL = "DELETE FROM " & dbSchema & ".Department WHERE department_id = " & deleteId
    
    On Error Resume Next
    db.Execute(deleteSQL)
    
    If Err.Number <> 0 Then
        Response.Write("<script>alert('부서 삭제 중 오류가 발생했습니다: " & Server.HTMLEncode(Err.Description) & "'); window.location.href='admin_department.asp';</script>")
    Else
        ' 활동 로그 기록
        LogActivity Session("user_id"), "부서삭제", "부서 삭제 (ID: " & deleteId & ")"
        Response.Write("<script>alert('부서가 삭제되었습니다.'); window.location.href='admin_department.asp';</script>")
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
Dim searchKeyword, whereClause
searchKeyword = Trim(Request.QueryString("keyword"))

whereClause = ""
If searchKeyword <> "" Then
    whereClause = " WHERE name LIKE '%" & PreventSQLInjection(searchKeyword) & "%' OR code LIKE '%" & PreventSQLInjection(searchKeyword) & "%'"
End If

' 전체 레코드 수
Dim countSQL, countRS
countSQL = "SELECT COUNT(*) AS cnt FROM " & dbSchema & ".Department" & whereClause
Set countRS = db99.Execute(countSQL)
totalCount = countRS("cnt")
totalPages = totalCount / pageSize

' 부서 목록 조회
Dim listSQL, listRS
listSQL = "SELECT * FROM (" & _
          "SELECT TOP " & pageSize & " * FROM (" & _
          "SELECT TOP " & (pageNo * pageSize) & " department_id, name, parent_id, created_at " & _
          "FROM " & dbSchema & ".Department" & whereClause & " " & _
          "ORDER BY department_id ASC) AS T1 " & _
          "ORDER BY department_id DESC) AS T2 " & _
          "ORDER BY department_id ASC"
          
Set listRS = db99.Execute(listSQL)

' 부서 목록 조회 (상위 부서 선택용)
Dim departmentsSQL, departmentsRS
departmentsSQL = "SELECT department_id, name FROM " & dbSchema & ".Department ORDER BY department_id"
Set departmentsRS = db99.Execute(departmentsSQL)  

' 상위 부서 이름 가져오기
Function GetParentDepartmentName(parentId)
    If IsNull(parentId) Or parentId = "" Then
        GetParentDepartmentName = "-"
        Exit Function
    End If
    
    Dim deptName, deptSQL, deptRS
    deptSQL = "SELECT name FROM " & dbSchema & ".Department WHERE department_id = " & parentId
    
    On Error Resume Next
    Set deptRS = db99.Execute(deptSQL)
    
    If Err.Number = 0 And Not deptRS.EOF Then
        deptName = deptRS("name")
    Else
        deptName = parentId
    End If
    
    If Not deptRS Is Nothing Then
        If deptRS.State = 1 Then
            deptRS.Close
        End If
        Set deptRS = Nothing
    End If
    
    GetParentDepartmentName = deptName
End Function

' POST 처리 - 부서 추가/수정
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    Dim action, departmentId, departmentCode, departmentName, description, parentId, isActive
    
    action = Request.Form("action")
    departmentCode = PreventSQLInjection(Request.Form("code"))
    departmentName = PreventSQLInjection(Request.Form("name"))
    description = PreventSQLInjection(Request.Form("description"))
    
    ' 상위 부서 ID 처리
    If Request.Form("parent_id") <> "" Then
        parentId = CInt(Request.Form("parent_id"))
    Else
        parentId = Null
    End If
    
    ' 활성 상태 확인
    If Request.Form("is_active") = "1" Then
        isActive = True
    Else
        isActive = False
    End If
    
    ' 유효성 검사
    If departmentName = "" Then
        Response.Write("<script>alert('부서명을 입력해주세요.'); history.back();</script>")
        Response.End
    End If
    
    On Error Resume Next
    
    If action = "add" Then
        ' 중복 코드 확인
        Dim checkSQL, checkRS
        checkSQL = "SELECT COUNT(*) AS cnt FROM " & dbSchema & ".Department WHERE name = '" & departmentName & "'"
        Set checkRS = db99.Execute(checkSQL)
        
        If Not checkRS.EOF And checkRS("cnt") > 0 Then
            Response.Write("<script>alert('이미 사용 중인 부서명입니다. 다른 이름을 입력해주세요.'); history.back();</script>")
            Response.End
        End If
        
        ' 부서 추가
        Dim addSQL
        addSQL = "INSERT INTO " & dbSchema & ".Department " & _
                 "( name, parent_id) " & _
                 "VALUES ('" & departmentName & "', " & parentId & ")"
        
        db99.Execute(addSQL)
        
        
        
            ' 활동 로그 기록
            LogActivity Session("user_id"), "부서추가", "부서 추가 (코드: " & departmentCode & ", 이름: " & departmentName & ")"
    
    ElseIf action = "edit" Then
        ' 부서 수정
        departmentId = Request.Form("department_id")
        
        If departmentId = "" Then
            Response.Write("<script>alert('부서 ID가 필요합니다.'); window.location.href='admin_department.asp';</script>")
            Response.End
        End If
        
        ' 중복 코드 확인 (자신 제외)
        Dim checkEditSQL, checkEditRS
        checkEditSQL = "SELECT COUNT(*) AS cnt FROM " & dbSchema & ".Department WHERE code = '" & departmentCode & "' AND department_id <> " & departmentId
        Set checkEditRS = db.Execute(checkEditSQL)
        
        If Not checkEditRS.EOF And checkEditRS("cnt") > 0 Then
            Response.Write("<script>alert('이미 사용 중인 부서 코드입니다. 다른 코드를 입력해주세요.'); history.back();</script>")
            Response.End
        End If
        
        ' 순환 참조 확인 (자신을 상위 부서로 설정하는 경우)
        If Not IsNull(parentId) And CStr(parentId) = CStr(departmentId) Then
            Response.Write("<script>alert('부서를 자신의 상위 부서로 설정할 수 없습니다.'); history.back();</script>")
            Response.End
        End If
        
        ' 부서 수정
        Dim editSQL
        editSQL = "UPDATE " & dbSchema & ".Department SET " & _
                  "code = ?, name = ?, description = ?, parent_id = ?, is_active = ? " & _
                  "WHERE department_id = ?"
        
        Dim cmdEdit
        Set cmdEdit = Server.CreateObject("ADODB.Command")
        cmdEdit.ActiveConnection = db
        cmdEdit.CommandText = editSQL
        cmdEdit.Parameters.Append cmdEdit.CreateParameter("@code", 200, 1, 20, departmentCode)
        cmdEdit.Parameters.Append cmdEdit.CreateParameter("@name", 200, 1, 50, departmentName)
        cmdEdit.Parameters.Append cmdEdit.CreateParameter("@description", 200, 1, 200, IIf(description = "", Null, description))
        cmdEdit.Parameters.Append cmdEdit.CreateParameter("@parent_id", 3, 1, , IIf(IsNull(parentId), Null, parentId))
        cmdEdit.Parameters.Append cmdEdit.CreateParameter("@is_active", 11, 1, , isActive)
        cmdEdit.Parameters.Append cmdEdit.CreateParameter("@department_id", 3, 1, , departmentId)
        
        cmdEdit.Execute
        
        If Err.Number <> 0 Then
            Response.Write("<script>alert('부서 수정 중 오류가 발생했습니다: " & Server.HTMLEncode(Err.Description) & "'); history.back();</script>")
            Response.End
        Else
            ' 활동 로그 기록
            LogActivity Session("user_id"), "부서수정", "부서 수정 (ID: " & departmentId & ", 코드: " & departmentCode & ")"
            Response.Write("<script>alert('부서가 수정되었습니다.'); window.location.href='admin_department.asp';</script>")
            Response.End
        End If
    End If
    
    On Error GoTo 0
End If
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
                    <a href="admin_department.asp" class="list-group-item list-group-item-action active">
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
                <div class="card-header bg-white d-flex justify-content-between align-items-center">
                    <h4 class="mb-0"><i class="fas fa-sitemap me-2"></i>부서 관리</h4>
                    <button class="btn btn-primary" data-bs-toggle="modal" data-bs-target="#addDeptModal">
                        <i class="fas fa-plus me-1"></i> 부서 등록
                    </button>
                </div>
                <div class="card-body">
                    <!-- 검색 폼 -->
                    <form action="admin_department.asp" method="get" class="mb-4">
                        <div class="row g-2">
                            <div class="col-md-10">
                                <input type="text" class="form-control" name="keyword" value="<%= searchKeyword %>" placeholder="부서명 또는 부서코드로 검색">
                            </div>
                            <div class="col-md-2">
                                <button type="submit" class="btn btn-primary w-100">검색</button>
                            </div>
                        </div>
                    </form>
                    
                    <!-- 부서 목록 -->
                    <div class="table-responsive">
                        <table class="table table-striped table-bordered table-hover">
                            <thead class="table-dark">
                                <tr>
                                    <th>ID</th>
                                    <th>부서명</th>
                                    <th>상위부서</th>
                                    <th>등록일</th>
                                    <th>관리</th>
                                </tr>
                            </thead>
                            <tbody>
                                <% 
                                If listRS.EOF Then 
                                %>
                                <tr>
                                    <td colspan="7" class="text-center">등록된 부서가 없습니다.</td>
                                </tr>
                                <% 
                                Else
                                    Do While Not listRS.EOF 
                                %>
                                <tr>
                                    <td><%= listRS("department_id") %></td>
                                   
                                    <td><%= listRS("name") %></td>
                                    <td><%= GetParentDepartmentName(listRS("parent_id")) %></td>
                                   
                                    <td><%= FormatDateTime(listRS("created_at"), 2) %></td>
                                    <td>
                                       <button class="btn btn-sm btn-danger" onclick="confirmDelete('<%=listRS("department_id")%>')">                                            <i class="fas fa-trash">삭제</i>                                        </button>
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
                                <a class="page-link" href="admin_department.asp?page=<%= pageNo - 1 %>&keyword=<%= searchKeyword %>">이전</a>
                            </li>
                            <% End If %>
                            
                            <% 
                            Dim startPage, endPage
                            startPage = Max(1, pageNo - 5)
                            endPage = Min(totalPages, pageNo + 5)
                            
                            For i = startPage To endPage
                            %>
                            <li class="page-item <% If i = pageNo Then %>active<% End If %>">
                                <a class="page-link" href="admin_department.asp?page=<%= i %>&keyword=<%= searchKeyword %>"><%= i %></a>
                            </li>
                            <% Next %>
                            
                            <% If pageNo < totalPages Then %>
                            <li class="page-item">
                                <a class="page-link" href="admin_department.asp?page=<%= pageNo + 1 %>&keyword=<%= searchKeyword %>">다음</a>
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

<!-- 부서 등록 모달 -->
<div class="modal fade" id="addDeptModal" tabindex="-1" aria-labelledby="addDeptModalLabel" aria-hidden="true">
    <div class="modal-dialog">
        <div class="modal-content">
            <form action="admin_department.asp" method="post" id="addDeptForm">
                <input type="hidden" name="action" value="add">
                <div class="modal-header">
                    <h5 class="modal-title" id="addDeptModalLabel">부서 등록</h5>
                    <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Close"></button>
                </div>
                <div class="modal-body">
                    <div class="row mb-3">
                       
                        <div class="col-md-12">
                            <label for="name" class="form-label">부서명 <span class="text-danger">*</span></label>
                            <input type="text" class="form-control" id="name" name="name" required maxlength="50">
                        </div>
                    </div>
                    <div class="mb-3">
                        <label for="parent_id" class="form-label">상위부서</label>
                        <select class="form-select" id="parent_id" name="parent_id">
                            <option value="">선택하세요</option>
                            <option value="">없음</option>
                            <% 
                            If Not departmentsRS.EOF Then
                                departmentsRS.MoveFirst
                                Do While Not departmentsRS.EOF 
                            %>
                            <option value="<%= departmentsRS("department_id") %>"><%= departmentsRS("name") %></option>
                            <% 
                                    departmentsRS.MoveNext
                                Loop
                            End If
                            %>
                        </select>
                    </div>
                   
                   
                </div>
                <div class="modal-footer">
                    <button type="button" class="btn btn-secondary" data-bs-dismiss="modal">취소</button>
                    <button type="submit" class="btn btn-primary">등록</button>
                </div>
            </form>
        </div>
    </div>
</div>

<script>
// 삭제 확인
function confirmDelete(id) {
    if (confirm("정말로 이 부서를 삭제하시겠습니까? 이 작업은 되돌릴 수 없습니다.")) {
        window.location.href = "admin_department.asp?action=delete&id=" + id;
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

If Not departmentsRS Is Nothing Then
    If departmentsRS.State = 1 Then
        departmentsRS.Close
    End If
    Set departmentsRS = Nothing
End If
%>

<!--#include file="../../includes/footer.asp"--> 