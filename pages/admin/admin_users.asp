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
        whereClause = " WHERE U.user_id LIKE '%" & PreventSQLInjection(searchKeyword) & "%'"
    ElseIf searchField = "name" Then
        whereClause = " WHERE U.name LIKE '%" & PreventSQLInjection(searchKeyword) & "%'"
    ElseIf searchField = "email" Then
        whereClause = " WHERE U.email LIKE '%" & PreventSQLInjection(searchKeyword) & "%'"
    ElseIf searchField = "department" Then
        ' 부서명으로 검색
        whereClause = " WHERE department_id IN (SELECT department_id FROM " & dbSchema & ".Department WHERE name LIKE '%" & PreventSQLInjection(searchKeyword) & "%')"
    End If
End If

' 전체 레코드 수
Dim countSQL, countRS
countSQL = "SELECT COUNT(*) AS cnt FROM " & dbSchema & ".Users U" & whereClause
Set countRS = db99.Execute(countSQL)
totalCount = countRS("cnt")
totalPages = (totalCount + pageSize - 1) \ pageSize

' 사용자 목록 조회
Dim listSQL, listRS
listSQL = "SELECT * FROM (" & _
          "SELECT TOP " & pageSize & " * FROM (" & _
          "SELECT TOP " & (pageNo * pageSize) & " U.user_id, U.name, U.email, U.phone, U.department_id, U.is_active, J.name AS job_grade " & _
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

<style>
.admin-container {
    max-width: 1400px;
    margin: 0 auto;
    padding: 2rem 1rem;
}

.admin-nav {
    background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
    border-radius: 16px;
    padding: 1.5rem;
    margin-bottom: 2rem;
    box-shadow: 0 8px 32px rgba(0,0,0,0.1);
}

.admin-nav-title {
    color: white;
    font-size: 1.25rem;
    font-weight: 600;
    margin-bottom: 1.5rem;
    display: flex;
    align-items: center;
}

.admin-nav-grid {
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
    gap: 0.75rem;
}

.admin-nav-item {
    background: rgba(255,255,255,0.1);
    border: 1px solid rgba(255,255,255,0.2);
    border-radius: 12px;
    padding: 1rem;
    color: white;
    text-decoration: none;
    transition: all 0.3s ease;
    display: flex;
    align-items: center;
    font-size: 0.9rem;
    font-weight: 500;
}

.admin-nav-item:hover {
    background: rgba(255,255,255,0.2);
    transform: translateY(-2px);
    color: white;
    text-decoration: none;
}

.admin-nav-item.active {
    background: rgba(255,255,255,0.25);
    border-color: rgba(255,255,255,0.4);
}

.admin-nav-item i {
    margin-right: 0.75rem;
    font-size: 1.1rem;
}

.page-header {
    display: flex;
    justify-content: space-between;
    align-items: center;
    margin-bottom: 2rem;
    padding: 1.5rem;
    background: white;
    border-radius: 12px;
    box-shadow: 0 2px 4px rgba(0,0,0,0.05);
}

.page-title {
    font-size: 1.5rem;
    font-weight: 600;
    color: #2C3E50;
    margin: 0;
}

.search-section {
    background: white;
    border-radius: 16px;
    padding: 2rem;
    margin-bottom: 2rem;
    box-shadow: 0 4px 20px rgba(0,0,0,0.08);
}

.search-title {
    font-size: 1.1rem;
    font-weight: 600;
    color: #2C3E50;
    margin-bottom: 1.5rem;
    display: flex;
    align-items: center;
}

.form-control, .form-select {
    border-radius: 8px;
    border: 2px solid #E9ECEF;
    padding: 0.875rem 1rem;
    font-size: 1rem;
    transition: all 0.2s ease;
}

.form-control:focus, .form-select:focus {
    border-color: #4A90E2;
    box-shadow: 0 0 0 4px rgba(74,144,226,0.1);
}

.btn {
    padding: 0.875rem 1.5rem;
    font-weight: 600;
    border-radius: 8px;
    transition: all 0.2s ease;
    margin: 0 0.25rem;
}

.btn-primary {
    background: linear-gradient(to right, #4A90E2, #5A9EEA);
    border: none;
    color: white;
}

.btn-primary:hover {
    transform: translateY(-2px);
    box-shadow: 0 4px 12px rgba(74,144,226,0.2);
}

.btn-success {
    background: linear-gradient(to right, #2ECC71, #27AE60);
    border: none;
    color: white;
}

.btn-success:hover {
    transform: translateY(-2px);
    box-shadow: 0 4px 12px rgba(46,204,113,0.2);
}

.btn-danger {
    background: linear-gradient(to right, #E74C3C, #C0392B);
    border: none;
    color: white;
}

.btn-danger:hover {
    transform: translateY(-2px);
    box-shadow: 0 4px 12px rgba(231,76,60,0.2);
}

.btn-secondary {
    background: #F8FAFC;
    border: 2px solid #E9ECEF;
    color: #2C3E50;
}

.btn-secondary:hover {
    background: #E9ECEF;
    transform: translateY(-2px);
}

.table-section {
    background: white;
    border-radius: 16px;
    padding: 2rem;
    box-shadow: 0 4px 20px rgba(0,0,0,0.08);
    margin-bottom: 2rem;
}

.table-title {
    font-size: 1.1rem;
    font-weight: 600;
    color: #2C3E50;
    margin-bottom: 1.5rem;
    display: flex;
    align-items: center;
}

.table {
    margin-bottom: 0;
    border-radius: 12px;
    overflow: hidden;
    box-shadow: 0 2px 8px rgba(0,0,0,0.05);
}

.table th {
    background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
    color: white;
    font-weight: 600;
    border: none;
    padding: 1rem;
    font-size: 0.95rem;
}

.table td {
    padding: 1rem;
    vertical-align: middle;
    border-bottom: 1px solid #E9ECEF;
    color: #2C3E50;
}

.table tbody tr:hover {
    background-color: #F8FAFC;
    transition: background-color 0.2s ease;
}

.status-badge {
    padding: 0.5rem 1rem;
    border-radius: 20px;
    font-size: 0.875rem;
    font-weight: 600;
}

.status-active {
    background: #E3F9E5;
    color: #1B873F;
}

.status-inactive {
    background: #FFE9E9;
    color: #DA3633;
}

.btn-sm {
    padding: 0.5rem 1rem;
    font-size: 0.875rem;
    margin: 0 0.125rem;
}

.pagination {
    margin-top: 2rem;
}

.page-link {
    border-radius: 8px;
    border: 2px solid #E9ECEF;
    color: #4A90E2;
    padding: 0.75rem 1rem;
    margin: 0 0.125rem;
    font-weight: 500;
}

.page-link:hover {
    background-color: #4A90E2;
    border-color: #4A90E2;
    color: white;
}

.page-item.active .page-link {
    background-color: #4A90E2;
    border-color: #4A90E2;
}

.empty-state {
    text-align: center;
    padding: 3rem;
    color: #64748B;
}

.empty-state i {
    font-size: 3rem;
    margin-bottom: 1rem;
    color: #CBD5E1;
}
</style>

<div class="admin-container">
    <!-- 관리자 네비게이션 -->
    <div class="admin-nav">
        <div class="admin-nav-title">
            <i class="fas fa-cog me-2"></i>관리자 메뉴
        </div>
        <div class="admin-nav-grid">
            <a href="admin_dashboard.asp" class="admin-nav-item">
                <i class="fas fa-tachometer-alt"></i>대시보드
            </a>
            <a href="admin_cardaccount.asp" class="admin-nav-item">
                <i class="fas fa-credit-card"></i>카드 계정 관리
            </a>
            <a href="admin_cardaccounttypes.asp" class="admin-nav-item">
                <i class="fas fa-tags"></i>카드 계정 유형 관리
            </a>
            <a href="admin_fuelrate.asp" class="admin-nav-item">
                <i class="fas fa-gas-pump"></i>유류비 단가 관리
            </a>
            <a href="admin_job_grade.asp" class="admin-nav-item">
                <i class="fas fa-user-tie"></i>직급 관리
            </a>
            <a href="admin_department.asp" class="admin-nav-item">
                <i class="fas fa-sitemap"></i>부서 관리
            </a>
            <a href="admin_users.asp" class="admin-nav-item active">
                <i class="fas fa-users"></i>사용자 관리
            </a>
            <a href="admin_card_usage.asp" class="admin-nav-item">
                <i class="fas fa-receipt"></i>카드 사용 내역 관리
            </a>
            <a href="admin_vehicle_requests.asp" class="admin-nav-item">
                <i class="fas fa-car"></i>차량 이용 신청 관리
            </a>
            <a href="admin_approvals.asp" class="admin-nav-item">
                <i class="fas fa-file-signature"></i>결재 로그 관리
            </a>
        </div>
    </div>

    <!-- 페이지 헤더 -->
    <div class="page-header">
        <h2 class="page-title">
            <i class="fas fa-users me-2"></i>사용자 관리
        </h2>
       
    </div>

    <!-- 검색 섹션 -->
    <div class="search-section">
        <div class="search-title">
            <i class="fas fa-search me-2"></i>사용자 검색
        </div>
        <form action="admin_users.asp" method="get">
            <div class="row g-3">
                <div class="col-md-3">
                    <label class="form-label">검색 필드</label>
                    <select class="form-select" name="field">
                        <option value="user_id" <% If searchField = "user_id" Then %>selected<% End If %>>사용자 ID</option>
                        <option value="name" <% If searchField = "name" Then %>selected<% End If %>>이름</option>
                        <option value="email" <% If searchField = "email" Then %>selected<% End If %>>이메일</option>
                        <option value="department" <% If searchField = "department" Then %>selected<% End If %>>부서</option>
                    </select>
                </div>
                <div class="col-md-7">
                    <label class="form-label">검색어</label>
                    <input type="text" class="form-control" name="keyword" value="<%= searchKeyword %>" placeholder="검색어를 입력하세요">
                </div>
                <div class="col-md-2">
                    <label class="form-label">&nbsp;</label>
                    <button type="submit" class="btn btn-primary w-100">
                        <i class="fas fa-search me-1"></i>검색
                    </button>
                </div>
            </div>
        </form>
    </div>

    <!-- 사용자 목록 -->
    <div class="table-section">
        <div class="table-title">
            <i class="fas fa-list me-2"></i>사용자 목록 (총 <%= totalCount %>명)
        </div>
        
        <% If listRS.EOF Then %>
        <div class="empty-state">
            <i class="fas fa-users"></i>
            <h5>등록된 사용자가 없습니다</h5>
            <p>새로운 사용자를 등록해보세요.</p>
        </div>
        <% Else %>
        <div class="table-responsive">
            <table class="table">
                <thead>
                    <tr>
                        <th style="text-align: center;">사용자 ID</th>
                        <th style="text-align: center;">이름</th>
                        <th style="text-align: center;">이메일</th>
                        <th style="text-align: center;">전화번호</th>
                        <th style="text-align: center;">부서</th>
                        <th style="text-align: center;">직급</th>
                        <th style="text-align: center;">상태</th>
                        <th style="text-align: center;">관리</th>
                    </tr>
                </thead>
                <tbody>
                    <% Do While Not listRS.EOF %>
                    <tr>
                        <td style="text-align: center;"><strong><%= listRS("user_id") %></strong></td>
                        <td style="text-align: center;"><%= listRS("name") %></td>
                        <td style="text-align: center;"><%= listRS("email") %></td>
                        <td style="text-align: center;"><%= listRS("phone") %></td>
                        <td style="text-align: center;"><%= GetDepartmentName(listRS("department_id")) %></td>
                        <td style="text-align: center;"><%= listRS("job_grade") %></td>
                        <td style="text-align: center;">
                            <% If listRS("is_active") = 1 Then %>
                                <span class="status-badge status-active">활성</span>
                            <% Else %>
                                <span class="status-badge status-inactive">비활성</span>
                            <% End If %>
                        </td>
                        <td style="text-align: center;">
                            <a href="admin_user_view.asp?id=<%= listRS("user_id") %>" class="btn btn-sm btn-primary">
                                <i class="fas fa-eye"></i> 상세
                            </a>
                            
                            <% If listRS("is_active") = 1 Then %>
                                <button class="btn btn-sm btn-danger" data-user-id="<%= listRS("user_id") %>" onclick="confirmDelete(this.getAttribute('data-user-id'))">
                                    <i class="fas fa-ban"></i> 비활성화
                                </button>
                            <% Else %>
                                <button class="btn btn-sm btn-success" data-user-id="<%= listRS("user_id") %>" onclick="confirmActive(this.getAttribute('data-user-id'))">
                                    <i class="fas fa-check"></i> 활성화
                                </button>
                            <% End If %>
                        </td>
                    </tr>
                    <% 
                    listRS.MoveNext
                    Loop
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
                    <a class="page-link" href="admin_users.asp?page=<%= pageNo - 1 %>&field=<%= searchField %>&keyword=<%= searchKeyword %>">
                        <i class="fas fa-chevron-left"></i> 이전
                    </a>
                </li>
                <% End If %>
                
                <% 
                Dim startPage, endPage
                If pageNo - 5 > 1 Then
                    startPage = pageNo - 5
                Else
                    startPage = 1
                End If
                
                If pageNo + 5 < totalPages Then
                    endPage = pageNo + 5
                Else
                    endPage = totalPages
                End If
                
                For i = startPage To endPage
                %>
                <li class="page-item <% If i = pageNo Then %>active<% End If %>">
                    <a class="page-link" href="admin_users.asp?page=<%= i %>&field=<%= searchField %>&keyword=<%= searchKeyword %>"><%= i %></a>
                </li>
                <% Next %>
                
                <% If pageNo < totalPages Then %>
                <li class="page-item">
                    <a class="page-link" href="admin_users.asp?page=<%= pageNo + 1 %>&field=<%= searchField %>&keyword=<%= searchKeyword %>">
                        다음 <i class="fas fa-chevron-right"></i>
                    </a>
                </li>
                <% End If %>
            </ul>
        </nav>
        <% End If %>
        <% End If %>
    </div>
</div>

<script>
// 삭제 확인
function confirmDelete(id) {
    if (confirm("정말로 이 사용자를 비활성화하시겠습니까?\n\n이 작업은 계정을 비활성화하며 실제로 삭제하지는 않습니다.")) {
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