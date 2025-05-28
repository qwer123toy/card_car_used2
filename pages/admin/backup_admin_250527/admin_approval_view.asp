<%@ Language="VBScript" CodePage="65001" %>
<% 
Response.CodePage = 65001
Response.CharSet = "utf-8"
%>

<!--#include file="../../db.asp"-->
<!--#include file="../../includes/functions.asp"-->
<%
If Not IsAuthenticated() Then RedirectTo("../../index.asp")
If Not IsAdmin() Then
    Response.Write("<script>alert('관리자 권한이 필요합니다.'); window.location.href='../dashboard.asp';</script>")
    Response.End
End If

Dim targetId, targetTable

targetId = Request.QueryString("target_id")
targetTable = Request.QueryString("target_table_name")

If targetId = "" Or targetTable = "" Then
    Response.Write("<script>alert('잘못된 접근입니다.'); location.href='admin_approvals.asp';</script>")
    Response.End
End If

Dim docSQL, docRS
Select Case LCase(targetTable)
    Case "cardusage"
        docSQL = "SELECT cu.*, u.name AS requester_name, u.email AS requester_email, d.name AS requester_department " & _
                 "FROM " & dbSchema & ".CardUsage cu " & _
                 "LEFT JOIN " & dbSchema & ".Users u ON cu.user_id = u.user_id " & _
                 "LEFT JOIN " & dbSchema & ".Department d ON u.department_id = d.department_id " & _
                 "WHERE cu.usage_id = " & targetId
    Case "vehiclerequests"
        docSQL = "SELECT vr.*, u.name AS requester_name, u.email AS requester_email, d.name AS requester_department " & _
                 "FROM " & dbSchema & ".VehicleRequests vr " & _
                 "LEFT JOIN " & dbSchema & ".Users u ON vr.user_id = u.user_id " & _
                 "LEFT JOIN " & dbSchema & ".Department d ON u.department_id = d.department_id " & _
                 "WHERE vr.request_id = " & targetId
    Case Else
        Response.Write("<script>alert('지원하지 않는 문서 유형입니다.'); location.href='admin_approvals.asp';</script>")
        Response.End
End Select

Set docRS = db.Execute(docSQL)
If docRS.EOF Then
    Response.Write("<script>alert('문서를 찾을 수 없습니다.'); location.href='admin_approvals.asp';</script>")
    Response.End
End If

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
                    <a href="admin_dashboard.asp" class="list-group-item list-group-item-action"><i class="fas fa-tachometer-alt me-2"></i>대시보드</a>
                    <a href="admin_cardaccount.asp" class="list-group-item list-group-item-action"><i class="fas fa-credit-card me-2"></i>카드 계정 관리</a>
                    <a href="admin_cardaccounttypes.asp" class="list-group-item list-group-item-action"><i class="fas fa-tags me-2"></i>카드 계정 유형 관리</a>
                    <a href="admin_fuelrate.asp" class="list-group-item list-group-item-action"><i class="fas fa-gas-pump me-2"></i>유류비 단가 관리</a>
                    <a href="admin_job_grade.asp" class="list-group-item list-group-item-action"><i class="fas fa-user-tie me-2"></i>직급 관리</a>
                    <a href="admin_department.asp" class="list-group-item list-group-item-action"><i class="fas fa-sitemap me-2"></i>부서 관리</a>
                    <a href="admin_users.asp" class="list-group-item list-group-item-action"><i class="fas fa-users me-2"></i>사용자 관리</a>
                    <a href="admin_card_usage.asp" class="list-group-item list-group-item-action"><i class="fas fa-receipt me-2"></i>카드 사용 내역 관리</a>
                    <a href="admin_vehicle_requests.asp" class="list-group-item list-group-item-action"><i class="fas fa-car me-2"></i>차량 이용 신청 관리</a>
                    <a href="admin_approvals.asp" class="list-group-item list-group-item-action active"><i class="fas fa-file-signature me-2"></i>결재 로그 관리</a>
                </div>
            </div>
        </div>
        <div class="col-md-9">
            <div class="card shadow-sm mb-4">
                <div class="card-header bg-white d-flex justify-content-between align-items-center">
                    <h4 class="mb-0"><i class="fas fa-file-signature me-2"></i>결재 문서 상세보기</h4>
                    <div>
                        <a href="admin_approvals.asp" class="btn btn-secondary">
                            <i class="fas fa-arrow-left me-1"></i> 목록으로
                        </a>
                    </div>
                </div>
                <div class="card-body">
                    <div class="row mb-4">
                        <div class="col-md-12">
                            <h5 class="border-bottom pb-2 mb-3">문서 정보</h5>
                            <table class="table table-bordered">
                                <tr>
                                    <th class="bg-light" style="width: 20%">문서 유형</th>
                                    <td><%= targetTable %></td>
                                </tr>
                                <tr>
                                    <th class="bg-light">제목</th>
                                    <td><%= IIf(IsNull(docRS("title")), "-", docRS("title")) %></td>
                                </tr>
                                <tr>
                                    <th class="bg-light">신청일</th>
                                    <td><%= FormatDate(docRS("created_at")) %></td>
                                </tr>
                                <tr>
                                    <th class="bg-light">사용 목적</th>
                                    <td><%= IIf(IsNull(docRS("purpose")), "-", docRS("purpose")) %></td>
                                </tr>
                            </table>
                        </div>
                    </div>
                    <div class="row mb-4">
                        <div class="col-md-12">
                            <h5 class="border-bottom pb-2 mb-3">신청자 정보</h5>
                            <table class="table table-bordered">
                                <tr>
                                    <th class="bg-light" style="width: 20%">이름</th>
                                    <td><%= IIf(IsNull(docRS("requester_name")), "-", docRS("requester_name")) %></td>
                                </tr>
                                <tr>
                                    <th class="bg-light">이메일</th>
                                    <td><%= IIf(IsNull(docRS("requester_email")), "-", docRS("requester_email")) %></td>
                                </tr>
                                <tr>
                                    <th class="bg-light">부서</th>
                                    <td><%= IIf(IsNull(docRS("requester_department")), "-", docRS("requester_department")) %></td>
                                </tr>
                            </table>
                        </div>
                    </div>
                </div>
            </div>
        </div>
    </div>
</div>

<% If Not docRS Is Nothing Then If docRS.State = 1 Then docRS.Close : Set docRS = Nothing %>
<!--#include file="../../includes/footer.asp"-->
