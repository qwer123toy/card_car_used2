<%@ Language="VBScript" CodePage="65001" %>
<% 
Response.CodePage = 65001
Response.CharSet = "utf-8"
%>

<!--#include file="../../db.asp"-->
<!--#include file="../../includes/functions.asp"-->
!-- Bootstrap 5 JS (Popper 포함된 번들) -->


<%
If Not IsAuthenticated() Then
    RedirectTo("../../index.asp")
End If

If Not IsAdmin() Then
    Response.Write("<script>alert('관리자 권한이 필요합니다.'); window.location.href='../dashboard.asp';</script>")
    Response.End
End If

If Request.QueryString("action") = "delete" And Request.QueryString("id") <> "" Then
    Dim deleteId
    deleteId = PreventSQLInjection(Request.QueryString("id"))

    Dim checkUseSQL, checkUseRS
    checkUseSQL = "SELECT COUNT(*) AS cnt FROM " & dbSchema & ".CardUsage WHERE card_id = " & deleteId
    Set checkUseRS = db.Execute(checkUseSQL)

    If Not checkUseRS.EOF And checkUseRS("cnt") > 0 Then
        Response.Write("<script>alert('이 카드는 사용 내역이 있어 삭제할 수 없습니다.'); window.location.href='admin_cardaccount.asp';</script>")
        Response.End
    End If

    Dim deleteSQL
    deleteSQL = "DELETE FROM " & dbSchema & ".CardAccount WHERE card_id = " & deleteId

    On Error Resume Next
    db.Execute(deleteSQL)

    If Err.Number <> 0 Then
        Response.Write("<script>alert('카드 계정 삭제 중 오류가 발생했습니다: " & Replace(Err.Description, "'", "\'") & "'); window.location.href='admin_cardaccount.asp';</script>")
    Else
        LogActivity Session("user_id"), "카드계정삭제", "카드 계정 삭제 (ID: " & deleteId & ")"
        Response.Write("<script>alert('카드 계정이 삭제되었습니다.'); window.location.href='admin_cardaccount.asp';</script>")
    End If
    On Error GoTo 0
    Response.End
End If

Dim pageNo, pageSize, totalCount, totalPages
pageSize = 10
If Request.QueryString("page") = "" Then
    pageNo = 1
Else
    pageNo = CInt(Request.QueryString("page"))
End If

Dim searchKeyword, whereClause
searchKeyword = Trim(Request.QueryString("keyword"))
whereClause = ""
If searchKeyword <> "" Then
    whereClause = " WHERE account_name LIKE '%" & PreventSQLInjection(searchKeyword) & "%'"
End If

Dim countSQL, countRS
countSQL = "SELECT COUNT(*) AS cnt FROM " & dbSchema & ".CardAccount AS ca " & whereClause
Set countRS = db99.Execute(countSQL)
totalCount = countRS("cnt")
totalPages = (totalCount + pageSize - 1) \ pageSize

Dim listSQL, listRS
listSQL = "SELECT * FROM (" & _
          "SELECT TOP " & pageSize & " * FROM (" & _
          "SELECT TOP " & (pageNo * pageSize) & " ca.card_id, ca.account_name, ca.issuer " & _
          "FROM " & dbSchema & ".CardAccount AS ca " & _
          whereClause & " ORDER BY ca.card_id) AS T1 ORDER BY card_id ASC) AS T2 ORDER BY card_id"
Set listRS = db99.Execute(listSQL)

Dim cardTypesSQL, cardTypesRS
cardTypesSQL = "SELECT account_type_id, type_name FROM " & dbSchema & ".CardAccountTypes ORDER BY type_name"
Set cardTypesRS = db99.Execute(cardTypesSQL)
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

.btn-danger {
    background: linear-gradient(to right, #E74C3C, #C0392B);
    border: none;
    color: white;
}

.btn-danger:hover {
    transform: translateY(-2px);
    box-shadow: 0 4px 12px rgba(231,76,60,0.2);
}

.btn-sm {
    padding: 0.5rem 1rem;
    font-size: 0.875rem;
    margin: 0 0.125rem;
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

/* 모달 스타일 개선 */
.modal-content {
    border: none;
    border-radius: 16px;
    box-shadow: 0 20px 40px rgba(0,0,0,0.15);
}

.modal-header {
    background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
    color: white;
    border-radius: 16px 16px 0 0;
    padding: 1.5rem 2rem;
    border-bottom: none;
}

.modal-title {
    font-weight: 600;
    font-size: 1.25rem;
    display: flex;
    align-items: center;
}



.btn-close {
    background: none;
    border: none;
    color: white;
    opacity: 0.8;
    font-size: 1.25rem;
    padding: 0.5rem;
    border-radius: 50%;
    transition: all 0.2s ease;
}

.btn-close:hover {
    opacity: 1;
    background: rgba(255,255,255,0.1);
    transform: scale(1.1);
}

.modal-body {
    padding: 2rem;
    background: #fff;
}

.modal-footer {
    padding: 1.5rem 2rem;
    background: #F8FAFC;
    border-radius: 0 0 16px 16px;
    border-top: 1px solid #E9ECEF;
}

.form-label {
    font-weight: 600;
    color: #2C3E50;
    margin-bottom: 0.75rem;
    font-size: 0.95rem;
}

.form-control {
    border-radius: 8px;
    border: 2px solid #E9ECEF;
    padding: 0.875rem 1rem;
    font-size: 1rem;
    transition: all 0.2s ease;
    background: #fff;
}

.form-control:focus {
    border-color: #4A90E2;
    box-shadow: 0 0 0 4px rgba(74,144,226,0.1);
    background: #fff;
}

.form-control::placeholder {
    color: #94A3B8;
    font-style: italic;
}

.mb-3 {
    margin-bottom: 1.75rem !important;
}

.modal-footer .btn {
    padding: 0.875rem 1.5rem;
    font-weight: 600;
    border-radius: 8px;
    transition: all 0.2s ease;
    margin-left: 0.5rem;
}

.modal-footer .btn-secondary {
    background: #F8FAFC;
    border: 2px solid #E9ECEF;
    color: #2C3E50;
}

.modal-footer .btn-secondary:hover {
    background: #E9ECEF;
    transform: translateY(-2px);
    color: #2C3E50;
}

.modal-footer .btn-primary {
    background: linear-gradient(to right, #4A90E2, #5A9EEA);
    border: none;
    box-shadow: 0 4px 12px rgba(74,144,226,0.2);
}

.modal-footer .btn-primary:hover {
    transform: translateY(-2px);
    box-shadow: 0 6px 16px rgba(74,144,226,0.3);
}
</style>
<script>
    function scrollToAddCardForm() {
        const form = document.getElementById('addCardModal');
        if (form) {
            form.scrollIntoView({ behavior: 'smooth' });
        }
    }
    </script>
    
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
            <a href="admin_cardaccount.asp" class="admin-nav-item active">
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
            <a href="admin_users.asp" class="admin-nav-item">
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
            <i class="fas fa-credit-card me-2"></i>카드 계정 관리
        </h2>
        <div>
            <button class="btn btn-primary" onclick="scrollToAddCardForm()">
                <i class="fas fa-plus me-1"></i> 카드 계정 등록
            </button>
        </div>
    </div>

    <!-- 검색 섹션 -->
    <div class="search-section">
        <div class="search-title">
            <i class="fas fa-search me-2"></i>카드 계정 검색
        </div>
        <form action="admin_cardaccount.asp" method="get">
            <div class="row g-3">
                <div class="col-md-10">
                    <label class="form-label">카드명</label>
                    <input type="text" class="form-control" name="keyword" value="<%= searchKeyword %>" placeholder="카드명을 입력하세요">
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

    <!-- 카드 계정 목록 -->
    <div class="table-section">
        <div class="table-title">
            <i class="fas fa-list me-2"></i>카드 계정 목록 (총 <%= totalCount %>개)
        </div>
        
        <% If listRS.EOF Then %>
        <div class="empty-state">
            <i class="fas fa-credit-card"></i>
            <h5>등록된 카드 계정이 없습니다</h5>
            <p>새로운 카드 계정을 등록해보세요.</p>
        </div>
        <% Else %>
        <div class="table-responsive">
            <table class="table">
                <thead>
                    <tr>
                        <th style="text-align: center;">카드 ID</th>
                        <th style="text-align: center;">카드명</th>
                        <th style="text-align: center;">카드회사</th>
                        <th style="text-align: center;">관리</th>
                    </tr>
                </thead>
                <tbody>
                    <% Do While Not listRS.EOF %>
                    <tr>
                        <td style="text-align: center;"><strong><%= listRS("card_id") %></strong></td>
                        <td style="text-align: center;"><%= listRS("account_name") %></td>
                        <td style="text-align: center;"><%= listRS("issuer") %></td>
                        <td style="text-align: center;">
                            <button class="btn btn-sm btn-danger" onclick="confirmDelete(<%= listRS("card_id") %>)">
                                <i class="fas fa-trash"></i> 삭제
                            </button>
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
                    <a class="page-link" href="admin_cardaccount.asp?page=<%= pageNo - 1 %>&keyword=<%= searchKeyword %>">
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
                    <a class="page-link" href="admin_cardaccount.asp?page=<%= i %>&keyword=<%= searchKeyword %>"><%= i %></a>
                </li>
                <% Next %>
                
                <% If pageNo < totalPages Then %>
                <li class="page-item">
                    <a class="page-link" href="admin_cardaccount.asp?page=<%= pageNo + 1 %>&keyword=<%= searchKeyword %>">
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

<!-- 카드 계정 등록 모달 -->
<div class="modal fade" id="addCardModal" tabindex="-1" aria-labelledby="addCardModalLabel" aria-hidden="true">
    <div class="modal-dialog modal-lg">
        <div class="modal-content">
            <form action="admin_cardaccount_process.asp" method="post" onsubmit="return confirm('등록하시겠습니까?');">
                <input type="hidden" name="action" value="add">
                <div class="modal-header">
                    <h5 class="modal-title">
                        <i class="fas fa-plus me-2"></i>카드 계정 등록
                    </h5>
                    <button type="button" class="btn-close" data-bs-dismiss="modal"></button>
                </div>
                <div class="modal-body">
                    <div class="mb-3">
                        <label class="form-label">카드명</label>
                        <input type="text" name="account_name" class="form-control" required>
                    </div>
                    <div class="mb-3">
                        <label class="form-label">카드회사</label>
                        <input type="text" name="issuer" class="form-control" required>
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
function confirmDelete(id) {
    if (confirm("정말로 이 카드 계정을 삭제하시겠습니까?")) {
        window.location.href = "admin_cardaccount.asp?action=delete&id=" + id;
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

If Not cardTypesRS Is Nothing Then
    If cardTypesRS.State = 1 Then
        cardTypesRS.Close
    End If
    Set cardTypesRS = Nothing
End If
%>

<!--#include file="../../includes/footer.asp"-->
