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


' 유류비 단가 삭제 처리
If Request.QueryString("action") = "delete" And Request.QueryString("id") <> "" Then
    Dim deleteId
    deleteId = PreventSQLInjection(Request.QueryString("id"))

    Dim checkLatestSQL, checkLatestRS
    checkLatestSQL = "SELECT TOP 1 fuel_rate_id FROM " & dbSchema & ".FuelRate ORDER BY date DESC"
    Set checkLatestRS = db.Execute(checkLatestSQL)

    If Not checkLatestRS.EOF And CStr(checkLatestRS("fuel_rate_id")) = CStr(deleteId) Then
        Response.Write("<script>alert('가장 최근 단가는 삭제할 수 없습니다.'); window.location.href='admin_fuelrate.asp';</script>")
        Response.End
    End If

    Dim deleteSQL
    deleteSQL = "DELETE FROM " & dbSchema & ".FuelRate WHERE fuel_rate_id = " & deleteId

    On Error Resume Next
    db.Execute(deleteSQL)

    If Err.Number <> 0 Then
        Response.Write("<script>alert('유류비 단가 삭제 중 오류가 발생했습니다: " & Replace(Err.Description, "'", "\'") & "'); window.location.href='admin_fuelrate.asp';</script>")
    Else
        LogActivity Session("user_id"), "유류비단가삭제", "유류비 단가 삭제 (ID: " & deleteId & ")"
        Response.Write("<script>alert('유류비 단가가 삭제되었습니다.'); window.location.href='admin_fuelrate.asp';</script>")
    End If
    On Error GoTo 0
    Response.End
End If

' 페이징 처리
Dim pageNo, pageSize, totalCount, totalPages
pageSize = 10
If Request.QueryString("page") = "" Then
    pageNo = 1
Else
    pageNo = CInt(Request.QueryString("page"))
End If

Dim countSQL, countRS
countSQL = "SELECT COUNT(*) AS cnt FROM " & dbSchema & ".FuelRate"
Set countRS = db99.Execute(countSQL)
totalCount = countRS("cnt")
totalPages = totalCount / pageSize

Dim listSQL, listRS
listSQL = "SELECT * FROM " & dbSchema & ".FuelRate ORDER BY date DESC"
Set listRS = db99.Execute(listSQL)

Dim latestRateSQL, latestRateRS, latestRate
latestRateSQL = "SELECT TOP 1 rate FROM " & dbSchema & ".FuelRate ORDER BY date DESC"
Set latestRateRS = db.Execute(latestRateSQL)

If Not latestRateRS.EOF Then
    latestRate = latestRateRS("rate")
Else
    latestRate = 0
End If

' POST 처리 - 유류비 단가 추가/수정
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    Dim action, rate, rateDate, rateId
    action = Request.Form("action")
    rate = CDbl(Replace(Request.Form("rate"), ",", ""))
    rateDate = Request.Form("date")
    rateId = Request.Form("id")

    If rate <= 0 Or rateDate = "" Then
        Response.Write("<script>alert('필수 항목을 모두 입력해주세요.'); history.back();</script>")
        Response.End
    End If

    On Error Resume Next

    If action = "add" Then
        Dim addSQL
        addSQL = "INSERT INTO " & dbSchema & ".FuelRate (rate, date) VALUES (" & rate & ", '" & rateDate & "')"
        db99.Execute(addSQL)

        If Err.Number <> 0 Then
            Dim msg
            msg = Replace(Server.HTMLEncode(Err.Description), "'", "\'")
            Response.Write("<script>alert('유류비 단가 추가 중 오류가 발생했습니다: " & msg & "'); history.back();</script>")
            Response.End
        Else
            LogActivity Session("user_id"), "유류비단가추가", "유류비 단가 추가 (단가: " & rate & ", 날짜: " & rateDate & ")"
            Response.Write("<script>alert('유류비 단가가 추가되었습니다.'); window.location.href='admin_fuelrate.asp';</script>")
            Response.End
        End If

    ElseIf action = "edit" Then
        If rateId = "" Then
            Response.Write("<script>alert('단가 ID가 필요합니다.'); window.location.href='admin_fuelrate.asp';</script>")
            Response.End
        End If

        Dim editSQL
        editSQL = "UPDATE " & dbSchema & ".FuelRate SET rate = " & rate & ", date = '" & rateDate & "' WHERE fuel_rate_id = " & rateId
        db99.Execute(editSQL)

        If Err.Number <> 0 Then
            
            msg = Replace(Server.HTMLEncode(Err.Description), "'", "\'")
            Response.Write("<script>alert('유류비 단가 수정 중 오류가 발생했습니다: " & msg & "'); history.back();</script>")
            Response.End
        Else
            LogActivity Session("user_id"), "유류비단가수정", "유류비 단가 수정 (ID: " & rateId & ", 단가: " & rate & ")"
            Response.Write("<script>alert('유류비 단가가 수정되었습니다.'); window.location.href='admin_fuelrate.asp';</script>")
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
                    <a href="admin_fuelrate.asp" class="list-group-item list-group-item-action active">
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
                <div class="card-header bg-white d-flex justify-content-between align-items-center">
                    <h4 class="mb-0"><i class="fas fa-gas-pump me-2"></i>유류비 단가 관리</h4>
                </div>
                <div class="card-body">
                    <!-- 현재 단가 표시 -->
                    <div class="alert alert-info mb-4">
                        <div class="d-flex justify-content-between align-items-center">
                            <div>
                                <h5 class="mb-1">현재 적용 단가</h5>
                                <p class="mb-0">차량 이용 신청서 작성 시 자동으로 적용되는 단가입니다.</p>
                            </div>
                            <div class="text-end">
                                <h3 class="mb-0 text-primary"><%= FormatNumber(latestRate) %> 원</h3>
                            </div>
                        </div>
                    </div>
                    
                    <!-- 유류비 단가 목록 -->
                    <div class="table-responsive">
                        <table class="table table-striped table-bordered table-hover">
                            <thead class="table-dark">
                                <tr>
                                    <th>ID</th>
                                    <th>단가(원)</th>
                                    <th>적용 일자</th>
                              
                                    <th>관리</th>
                                </tr>
                            </thead>
                            <tbody>
                                <% 
                                If listRS.EOF Then 
                                %>
                                <tr>
                                    <td colspan="7" class="text-center">등록된 유류비 단가가 없습니다.</td>
                                </tr>
                                <% 
                                Else
                                    Dim isLatest
                                    isLatest = True ' 첫 번째 행이 최신
                                    
                                    Do While Not listRS.EOF 
                                %>
                                <tr>
                                    <td><%= listRS("fuel_rate_id") %></td>
                                    <td class="text-end"><%= FormatNumber(listRS("rate")) %> 원</td>
                                    <td><%= FormatDate(listRS("date")) %></td>
                                    
                                       
                                    <td>
                                        
                                        <% If Not isLatest Then ' 최신 단가가 아닌 경우에만 삭제 버튼 표시 %>
                                        <button class="btn btn-sm btn-danger" onclick="confirmDelete(<%= listRS("fuel_rate_id") %>)">
                                            <i class="fas fa-trash">삭제</i>
                                        </button>
                                        <% End If %>
                                    </td>
                                </tr>
                                <% 
                                        isLatest = False ' 첫 번째 행 이후부터는 최신이 아님
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
                                <a class="page-link" href="admin_fuelrate.asp?page=<%= pageNo - 1 %>">이전</a>
                            </li>
                            <% End If %>
                            
                            <% 
                            Dim startPage, endPage
                            startPage = Max(1, pageNo - 5)
                            endPage = Min(totalPages, pageNo + 5)
                            
                            For i = startPage To endPage
                            %>
                            <li class="page-item <% If i = pageNo Then %>active<% End If %>">
                                <a class="page-link" href="admin_fuelrate.asp?page=<%= i %>"><%= i %></a>
                            </li>
                            <% Next %>
                            
                            <% If pageNo < totalPages Then %>
                            <li class="page-item">
                                <a class="page-link" href="admin_fuelrate.asp?page=<%= pageNo + 1 %>">다음</a>
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

<!-- 단가 등록 모달 -->
<div class="modal fade" id="addRateModal" tabindex="-1" aria-labelledby="addRateModalLabel" aria-hidden="true">
    <div class="modal-dialog">
        <div class="modal-content">
            <form action="admin_fuelrate.asp" method="post" id="addRateForm">
                <input type="hidden" name="action" value="add">
                <div class="modal-header">
                    <h5 class="modal-title" id="addRateModalLabel">유류비 단가 등록</h5>
                    <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Close"></button>
                </div>
                <div class="modal-body">
                    <div class="mb-3">
                        <label for="rate" class="form-label">단가(원) <span class="text-danger">*</span></label>
                        <input type="text" class="form-control" id="rate" name="rate" required onkeyup="formatNumberInput(this)">
                        <div class="form-text">숫자만 입력하세요. (예: 2,000)</div>
                    </div>
                    <div class="mb-3">
                        <label for="date" class="form-label">적용 일자 <span class="text-danger">*</span></label>
                        <input type="date" class="form-control" id="date" name="date" required value="<%= FormatDateForInput(Date()) %>">
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
    if (confirm('정말로 이 유류비 단가를 삭제하시겠습니까? 이 작업은 되돌릴 수 없습니다.')) {
        window.location.href = 'admin_fuelrate.asp?action=delete&id=' + id;
    }
}

// 숫자 입력 필드 포맷팅
function formatNumberInput(input) {
    // 콤마 제거 및 숫자만 추출
    let value = input.value.replace(/,/g, '');
    value = value.replace(/[^\d]/g, '');
    
    // 숫자 포맷팅 (천 단위 콤마)
    if (value) {
        value = parseInt(value, 10).toLocaleString('ko-KR');
    }
    
    // 값 업데이트
    input.value = value;
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

If Not latestRateRS Is Nothing Then
    If latestRateRS.State = 1 Then
        latestRateRS.Close
    End If
    Set latestRateRS = Nothing
End If
%>

<!--#include file="../../includes/footer.asp"--> 