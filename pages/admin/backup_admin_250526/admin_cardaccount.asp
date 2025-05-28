<%@ Language="VBScript" CodePage="65001" %>
<% 
Response.CodePage = 65001
Response.CharSet = "utf-8"
%>

<!--#include file="../../db.asp"-->
<!--#include file="../../includes/functions.asp"-->
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
totalPages = totalCount / pageSize

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
                <div class="card-body">
                    <form action="admin_cardaccount.asp" method="get" class="mb-4">
                        <div class="row g-2">
                            <div class="col-md-10">
                                <input type="text" class="form-control" name="keyword" value="<%= searchKeyword %>" placeholder="카드번호을 입력하세요">
                            </div>
                            <div class="col-md-2">
                                <button type="submit" class="btn btn-primary w-100">검색</button>
                            </div>
                        </div>
                    </form>

                    <div class="table-responsive">
                        <table class="table table-bordered">
                            <thead class="table-dark">
                                <tr>
                                    <th>순번</th>
                                    <th>카드명</th>
                                    <th>카드회사</th>
                                    
                                    <th>관리</th>
                                </tr>
                            </thead>
                            <tbody>
                                <% If listRS.EOF Then %>
                                <tr><td colspan="5" class="text-center">등록된 카드 계정이 없습니다.</td></tr>
                                <% Else
                                    Do While Not listRS.EOF %>
                                <tr>
                                    <td><%= listRS("card_id") %></td>
                                    <td><%= listRS("account_name") %></td>
                                    <td><%= listRS("issuer") %></td>
                                    <td>
                                        <button class="btn btn-sm btn-danger" onclick="confirmDelete(<%= listRS("card_id") %>)">삭제</button>
                                    </td>
                                </tr>
                                <% listRS.MoveNext: Loop: End If %>
                            </tbody>
                        </table>
                    </div>
                </div>
            </div>
        </div>
    </div>
</div>

<div class="modal fade" id="addCardModal" tabindex="-1" aria-labelledby="addCardModalLabel" aria-hidden="true">
    <div class="modal-dialog modal-lg">
        <div class="modal-content">
            <form action="admin_cardaccount_process.asp" method="post">
                <input type="hidden" name="action" value="add">
                <div class="modal-header">
                    <h5 class="modal-title">카드 계정 등록</h5>
                    <button type="button" class="btn-close" data-bs-dismiss="modal"></button>
                </div>
                <div class="modal-body">
                    <div class="mb-3">
                        <label>카드명</label>
                        <input type="text" name="account_name" class="form-control" required>
                    </div>
                    <div class="mb-3">
                        <label>카드회사</label>
                        <input type="text" name="issuer" class="form-control" required>
                    </div>
                </div>
                <div class="modal-footer">
                    <button type="submit" class="btn btn-primary">등록</button>
                </div>
            </form>
        </div>
    </div>
</div>

<script>
function confirmDelete(id) {
    if (confirm('정말 삭제하시겠습니까?')) {
        location.href = 'admin_cardaccount.asp?action=delete&id=' + id;
    }
}
</script>

<%
If Not listRS Is Nothing Then If listRS.State = 1 Then listRS.Close
Set listRS = Nothing
If Not cardTypesRS Is Nothing Then If cardTypesRS.State = 1 Then cardTypesRS.Close
Set cardTypesRS = Nothing
%>

<!--#include file="../../includes/footer.asp"-->
