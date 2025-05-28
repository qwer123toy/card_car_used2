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

' 사용 내역 ID 확인
Dim usageId
usageId = Request.QueryString("id")

If usageId = "" Then
    Response.Write("<script>alert('잘못된 접근입니다.'); window.location.href='admin_card_usage.asp';</script>")
    Response.End
End If

' 카드 사용 내역 정보 조회
Dim usageSQL, usageRS
usageSQL = "SELECT cu.*, " & _
           "u.user_id AS user_id, u.name AS user_name, u.email AS user_email, " & _
           "d.name AS department_name, ca.account_name AS card_name, ca.issuer, " & _
           "cat.type_name AS category_name, cat.account_type_id AS category_id " & _
           "FROM " & dbSchema & ".CardUsage cu " & _
           "LEFT JOIN " & dbSchema & ".Users u ON cu.user_id = u.user_id " & _
           "LEFT JOIN " & dbSchema & ".Department d ON u.department_id = d.department_id " & _
           "LEFT JOIN " & dbSchema & ".CardAccount ca ON cu.card_id = ca.card_id " & _
           "LEFT JOIN " & dbSchema & ".CardAccountTypes cat ON cu.expense_category_id = cat.account_type_id " & _
           "WHERE cu.usage_id = " & usageId

Set usageRS = db99.Execute(usageSQL)

' 계정과목 조회
Dim categorySQL, categoryRS
categorySQL = "SELECT account_type_id, type_name FROM " & dbSchema & ".CardAccountTypes ORDER BY type_name"
Set categoryRS = db99.Execute(categorySQL)

'카드 목록 조회
Dim cardSQL, cardRS
cardSQL = "SELECT card_id, account_name, issuer FROM " & dbSchema & ".CardAccount ORDER BY card_id"
Set cardRS = db99.Execute(cardSQL)


' 데이터가 없으면 목록으로 리다이렉션
If usageRS.EOF Then
    Response.Write("<script>alert('해당 카드 사용 내역 정보를 찾을 수 없습니다.'); window.location.href='admin_card_usage.asp';</script>")
    Response.End
End If

' 날짜 포맷
Function FormatDate(dateValue)
    If IsNull(dateValue) Or Not IsDate(dateValue) Then
        FormatDate = "-"
    Else
        FormatDate = FormatDateTime(dateValue, 2) & " " & FormatDateTime(dateValue, 4)
    End If
End Function


' 금액 포맷
Function FormatCurrency(amount)
    If IsNull(amount) Or Not IsNumeric(amount) Then
        FormatCurrency = "0원"
    Else
        FormatCurrency = AddComma(CLng(amount)) & "원"
    End If
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


' 영수증 파일 경로 가져오기
Function GetReceiptFilePath(fileName)
    If IsNull(fileName) Or fileName = "" Then
        GetReceiptFilePath = ""
    Else
        GetReceiptFilePath = "../../uploads/" & fileName
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
                    <a href="admin_card_usage.asp" class="list-group-item list-group-item-action active">
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
                    <h4 class="mb-0"><i class="fas fa-receipt me-2"></i>카드 사용 내역 상세보기</h4>
                    <div>
                        <a href="admin_card_usage.asp" class="btn btn-secondary">
                            <i class="fas fa-arrow-left me-1"></i> 목록으로
                        </a>
                    </div>
                </div>
                <form method="post" action="admin_card_usage_process.asp">
                    <input type="hidden" name="usage_id" value="<%= usageRS("usage_id") %>">
                <div class="card-body">
                    <div class="row mb-4">
                        <div class="col-md-6">
                            <h5 class="border-bottom pb-2 mb-3">사용 정보</h5>
                            <table class="table table-bordered">
                                <tr>
                                    <th class="bg-light">사용일</th>
                                    <td><input type="date" class="form-control" name="usage_date" value="<%= FormatDateTime(usageRS("usage_date"), 2) %>"></td>
                                </tr>
                                <tr>
                                    <th class="bg-light">금액</th>
                                    <td><input type="number" class="form-control" name="amount" value="<%= usageRS("amount") %>"> 원</td>
                                </tr>
                                <tr>
                                    <th class="bg-light">가맹점</th>
                                    <td><input type="text" class="form-control" name="store_name" value="<%= usageRS("store_name") %>"></td>

                                </tr>
                                <tr>
                                    <th class="bg-light">계정과목</th>
                                    <td>
                                        <select name="category_id" class="form-select">
                                            <% 
                                            ' 카테고리 목록 출력
                                            Do While Not categoryRS.EOF 
                                                Dim selectedCat
                                                selectedCat = ""
                                                If usageRS("category_id") = categoryRS("account_type_id") Then selectedCat = "selected"
                                            %>
                                                <option value="<%= categoryRS("account_type_id") %>" <%= selectedCat %>><%= categoryRS("type_name") %></option> 
                                            <% 
                                                categoryRS.MoveNext 
                                            Loop 
                                            %>
                                        </select>
                                    </td>
                                </tr>
                                <tr>
                                    <th class="bg-light">상태</th>
                                    <td><%= IIf(IsNull(usageRS("approval_status")), "-", usageRS("approval_status")) %></td>
                                </tr>
                                <tr>
                                    <th class="bg-light">사용 목적</th>
                                    <td><input type="text" class="form-control" name="purpose" value="<%= usageRS("purpose") %>"></td>
                                </tr>
                            </table>
                        </div>
                        <div class="col-md-6">
                            <h5 class="border-bottom pb-2 mb-3">카드 정보</h5>
                            <table class="table table-bordered">
                                <tr>
                                    <th class="bg-light" width="30%">카드명</th>
                                    <td>
                                        <select name="card_id" class="form-select">
                                            <% 
                                            ' 카드 목록 출력
                                            Do While Not cardRS.EOF 
                                                Dim selectedCard
                                                selectedCard = ""
                                                If usageRS("card_id") = cardRS("card_id") Then selectedCard = "selected"
                                            %>
                                                <option value="<%= cardRS("card_id") %>" <%= selectedCard %>><%= cardRS("account_name") %> (<%= cardRS("issuer") %>)</option> 
                                            <% 
                                                cardRS.MoveNext 
                                            Loop 
                                            %>
                                        </select>
                                    </td>
                                </tr>
                               
                            </table>
                            
                            <h5 class="border-bottom pb-2 mb-3 mt-4">사용자 정보</h5>
                            <table class="table table-bordered">
                                <tr>
                                    <th class="bg-light" width="30%">이름</th>
                                    <td><%= IIf(IsNull(usageRS("user_name")), "-", usageRS("user_name")) %></td>
                                </tr>
                                <tr>
                                    <th class="bg-light">이메일</th>
                                    <td><%= IIf(IsNull(usageRS("user_email")), "-", usageRS("user_email")) %></td>
                                </tr>
                               
                                <tr>
                                    <th class="bg-light">부서</th>
                                    <td><%= IIf(IsNull(usageRS("department_name")), "-", usageRS("department_name")) %></td>
                                </tr>
                            </table>
                        </div>
                    </div>
                    
                    
                    <div class="row">
                        <div class="col-md-12">
                            <div class="d-flex justify-content-between">
                                <div class="text-end">
                                    <button type="submit" class="btn btn-primary">
                                        <i class="fas fa-save me-1"></i> 수정
                            </div>
                        </div>
                    </div>
                </div>
                </form>
            </div>
        </div>
    </div>
</div>

<%
' 사용한 객체 해제
If Not usageRS Is Nothing Then
    If usageRS.State = 1 Then
        usageRS.Close
    End If
    Set usageRS = Nothing
End If
%>

<!--#include file="../../includes/footer.asp"--> 