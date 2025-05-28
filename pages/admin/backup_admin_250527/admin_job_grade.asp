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

' 직급 삭제 처리
If Request.QueryString("action") = "delete" And Request.QueryString("id") <> "" Then
    Dim deleteId
    deleteId = PreventSQLInjection(Request.QueryString("id"))
    
    ' 삭제 전 해당 직급이 사용중인지 확인
    Dim checkUseSQL, checkUseRS
    checkUseSQL = "SELECT COUNT(*) AS cnt FROM " & dbSchema & ".Users WHERE job_grade = " & deleteId
    Set checkUseRS = db.Execute(checkUseSQL)
    
    If Not checkUseRS.EOF And checkUseRS("cnt") > 0 Then
        Response.Write("<script>alert('이 직급은 사용자에게 할당되어 있어 삭제할 수 없습니다.'); window.location.href='admin_job_grade.asp';</script>")
        Response.End
    End If
    
    ' 삭제 쿼리 실행
    Dim deleteSQL
    deleteSQL = "DELETE FROM " & dbSchema & ".Job_Grade WHERE job_grade_id = " & deleteId
    
    On Error Resume Next
    db.Execute(deleteSQL)
    
    If Err.Number <> 0 Then
        Response.Write("<script>alert('직급 삭제 중 오류가 발생했습니다: " & Replace(Err.Description, "'", "\'") & "'); window.location.href='admin_job_grade.asp';</script>")
    Else
        ' 활동 로그 기록
        LogActivity Session("user_id"), "직급삭제", "직급 삭제 (ID: " & deleteId & ")"
        Response.Write("<script>alert('직급이 삭제되었습니다.'); window.location.href='admin_job_grade.asp';</script>")
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
    whereClause = " WHERE name LIKE '%" & PreventSQLInjection(searchKeyword) & "%'"
End If

' 전체 레코드 수
Dim countSQL, countRS
countSQL = "SELECT COUNT(*) AS cnt FROM " & dbSchema & ".Job_Grade" & whereClause
Set countRS = db99.Execute(countSQL)
totalCount = countRS("cnt")
totalPages = totalCount / pageSize

' 직급 목록 조회
Dim listSQL, listRS
listSQL = "SELECT * FROM " & dbSchema & ".Job_Grade" & whereClause & " ORDER BY sort_order"
          
Set listRS = db99.Execute(listSQL)

' POST 처리 - 직급 추가/수정
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    Dim action, jobGradeId, jobGradeName, sortOrder, description
    
    action = Request.Form("action")
    jobGradeName = PreventSQLInjection(Request.Form("name"))
    sortOrder = CInt(Request.Form("sort_order"))
    description = PreventSQLInjection(Request.Form("description"))
    
    ' 유효성 검사
    If jobGradeName = "" Then
        Response.Write("<script>alert('직급명을 입력해주세요.'); history.back();</script>")
        Response.End
    End If
    
    On Error Resume Next
    
    If action = "add" Then
        ' 직급 추가
        Dim addSQL
        addSQL = "INSERT INTO " & dbSchema & ".Job_Grade " & _
                 "(name, sort_order) " & _
                 "VALUES ('" & jobGradeName & "', " & sortOrder & ")"& _
                 "ORDER BY sort_order DESC"
        
        db99.Execute(addSQL)
        
        If Err.Number <> 0 Then
            Response.Write("<script>alert('직급 추가 중 오류가 발생했습니다: " & Replace(Err.Description, "'", "\'") & "'); history.back();</script>")
            Response.End
        Else
            ' 활동 로그 기록
            LogActivity Session("user_id"), "직급추가", "직급 추가 (이름: " & jobGradeName & ")"
            Response.Write("<script>alert('직급이 추가되었습니다.'); window.location.href='admin_job_grade.asp';</script>")
            Response.End
        End If
    
    ElseIf action = "edit" Then
        ' 직급 수정
        jobGradeId = Request.Form("job_grade_id")
        
        If jobGradeId = "" Then
            Response.Write("<script>alert('직급 ID가 필요합니다.'); window.location.href='admin_job_grade.asp';</script>")
            Response.End
        End If
        
        Dim editSQL
        editSQL = "UPDATE " & dbSchema & ".Job_Grade SET " & _
                  "name = ?, sort_order = ?, description = ? " & _
                  "WHERE job_grade_id = ?"
        
        Dim cmdEdit
        Set cmdEdit = Server.CreateObject("ADODB.Command")
        cmdEdit.ActiveConnection = db
        cmdEdit.CommandText = editSQL
        cmdEdit.Parameters.Append cmdEdit.CreateParameter("@name", 200, 1, 50, jobGradeName)
        cmdEdit.Parameters.Append cmdEdit.CreateParameter("@sort_order", 3, 1, , sortOrder)  ' adInteger
        cmdEdit.Parameters.Append cmdEdit.CreateParameter("@description", 200, 1, 200, IIf(description = "", Null, description))
        cmdEdit.Parameters.Append cmdEdit.CreateParameter("@job_grade_id", 3, 1, , jobGradeId)
        
        cmdEdit.Execute
        
        If Err.Number <> 0 Then
            Response.Write("<script>alert('직급 수정 중 오류가 발생했습니다: " & Replace(Err.Description, "'", "\'") & "'); history.back();</script>")
            Response.End
        Else
            ' 활동 로그 기록
            LogActivity Session("user_id"), "직급수정", "직급 수정 (ID: " & jobGradeId & ", 이름: " & jobGradeName & ")"
            Response.Write("<script>alert('직급이 수정되었습니다.'); window.location.href='admin_job_grade.asp';</script>")
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
                    <a href="admin_job_grade.asp" class="list-group-item list-group-item-action active">
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
                <div class="card-header bg-white d-flex justify-content-between align-items-center">
                    <h4 class="mb-0"><i class="fas fa-user-tie me-2"></i>직급 관리</h4>

                </div>
                <div class="card-body">
                    <!-- 검색 폼 -->
                    <form action="admin_job_grade.asp" method="get" class="mb-4">
                        <div class="row g-2">
                            <div class="col-md-10">
                                <input type="text" class="form-control" name="keyword" value="<%= searchKeyword %>" placeholder="직급명으로 검색">
                            </div>
                            <div class="col-md-2">
                                <button type="submit" class="btn btn-primary w-100">검색</button>
                            </div>
                        </div>
                    </form>
                    
                    <!-- 직급 목록 -->
                    <div class="table-responsive">
                        <table class="table table-striped table-bordered table-hover">
                            <thead class="table-dark">
                                <tr>
                                    <th>ID</th>
                                    <th>직급명</th>
                                    <th>직급순서</th>
                                    <th>관리</th>
                                </tr>
                            </thead>
                            <tbody>
                                <% 
                                If listRS.EOF Then 
                                %>
                                <tr>
                                    <td colspan="6" class="text-center">등록된 직급이 없습니다.</td>
                                </tr>
                                <% 
                                Else
                                    Do While Not listRS.EOF 
                                %>
                                <tr>
                                    <td><%= listRS("job_grade_id") %></td>
                                    <td><%= listRS("name") %></td>
                                    <td><%= listRS("sort_order") %></td>
                                    <td>
                                        <button class="btn btn-sm btn-danger" onclick="confirmDelete(<%= listRS("job_grade_id") %>)">
                                            <i class="fas fa-trash">삭제</i>
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
                                <a class="page-link" href="admin_job_grade.asp?page=<%= pageNo - 1 %>&keyword=<%= searchKeyword %>">이전</a>
                            </li>
                            <% End If %>
                            
                            <% 
                            Dim startPage, endPage
                            startPage = Max(1, pageNo - 5)
                            endPage = Min(totalPages, pageNo + 5)
                            
                            For i = startPage To endPage
                            %>
                            <li class="page-item <% If i = pageNo Then %>active<% End If %>">
                                <a class="page-link" href="admin_job_grade.asp?page=<%= i %>&keyword=<%= searchKeyword %>"><%= i %></a>
                            </li>
                            <% Next %>
                            
                            <% If pageNo < totalPages Then %>
                            <li class="page-item">
                                <a class="page-link" href="admin_job_grade.asp?page=<%= pageNo + 1 %>&keyword=<%= searchKeyword %>">다음</a>
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

<!-- 직급 등록 모달 -->
<div class="modal fade" id="addGradeModal" tabindex="-1" aria-labelledby="addGradeModalLabel" aria-hidden="true">
    <div class="modal-dialog">
        <div class="modal-content">
            <form action="admin_job_grade.asp" method="post" id="addGradeForm">
                <input type="hidden" name="action" value="add">
                <div class="modal-header">
                    <h5 class="modal-title" id="addGradeModalLabel">직급 추가</h5>
                    <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Close"></button>
                </div>
                <div class="modal-body">
                    <div class="mb-3">
                        <label for="name" class="form-label">직급명 <span class="text-danger">*</span></label>
                        <input type="text" class="form-control" id="name" name="name" required maxlength="50">
                    </div>
                    <div class="mb-3">
                        <label for="name" class="form-label">직급순서 <span class="text-danger">*</span></label>
                        <input type="text" class="form-control" id="sort_order" name="sort_order" required maxlength="50">
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
    if (confirm('정말로 이 직급을 삭제하시겠습니까? 이 작업은 되돌릴 수 없습니다.')) {
        window.location.href = 'admin_job_grade.asp?action=delete&id=' + id;
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
%>

<!--#include file="../../includes/footer.asp"--> 