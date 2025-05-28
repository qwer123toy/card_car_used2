<%@ Language="VBScript" CodePage="65001" %>
<%
Response.CodePage = 65001
Response.CharSet = "utf-8"

' 오류 처리 초기화
On Error Resume Next
%>

<!--#include file="../db.asp"-->
<!--#include file="../includes/functions.asp"-->

<%
' 로그인 상태 확인
If Not IsAuthenticated() Then
    Response.Write "<script>alert('로그인이 필요합니다.'); location.href='../index.asp';</script>"
    Response.End
End If

' POST 요청 확인
If Request.ServerVariables("REQUEST_METHOD") <> "POST" Then
    Response.Write "<script>alert('잘못된 접근입니다.'); history.back();</script>"
    Response.End
End If

' 파라미터 받기
Dim requestId, startDate, endDate, purpose, startLocation, destination, distance, title
requestId = Request.Form("request_id")
startDate = Request.Form("start_date")
endDate = Request.Form("end_date")
distance = Request.Form("distance")
purpose = Request.Form("purpose")
startLocation = Request.Form("start_location")
destination = Request.Form("destination")
title = Request.Form("title")

' 필수 값 검증
If requestId = "" Or startDate = "" Or endDate = "" Or distance = "" Or purpose = "" Or startLocation = "" Or destination = "" Then
    Response.Write "<script>alert('필수 항목이 누락되었습니다.'); history.back();</script>"
    Response.End
End If

' 종료일자가 비어있으면 시작일자와 동일하게 설정
If endDate = "" Then
    endDate = startDate
End If

' 숫자 값 안전하게 변환
If IsNumeric(Replace(distance, ",", "")) Then
    distance = CDbl(Replace(distance, ",", ""))
Else
    Response.Write "<script>alert('운행거리는 숫자로 입력해주세요.'); history.back();</script>"
    Response.End
End If

' 현재 사용자가 문서 작성자인지 확인
Dim checkOwnerSQL, checkOwnerRS
checkOwnerSQL = "SELECT user_id, approval_status FROM " & dbSchema & ".VehicleRequests WHERE request_id = " & requestId
Set checkOwnerRS = db.Execute(checkOwnerSQL)

If checkOwnerRS.EOF Then
    Response.Write "<script>alert('존재하지 않는 문서입니다.'); history.back();</script>"
    Response.End
End If

If checkOwnerRS("user_id") <> Session("user_id") Then
    Response.Write "<script>alert('수정 권한이 없습니다.'); history.back();</script>"
    Response.End
End If

If checkOwnerRS("approval_status") = "완료" Then
    Response.Write "<script>alert('완료 상태의 문서는 수정할 수 없습니다.'); history.back();</script>"
    Response.End
End If

' 트랜잭션 시작
db.BeginTrans

' 1. VehicleRequests 테이블 업데이트
Dim updateSQL
updateSQL = "UPDATE " & dbSchema & ".VehicleRequests SET " & _
            "start_date = '" & startDate & "', " & _
            "end_date = '" & endDate & "', " & _
            "distance = " & distance & ", " & _
            "purpose = '" & PreventSQLInjection(purpose) & "', " & _
            "start_location = '" & PreventSQLInjection(startLocation) & "', " & _
            "destination = '" & PreventSQLInjection(destination) & "', " & _
            "approval_status = '대기'"

' title이 제공된 경우 업데이트에 포함
If title <> "" Then
    updateSQL = updateSQL & ", title = '" & PreventSQLInjection(title) & "' "
End If

updateSQL = updateSQL & " WHERE request_id = " & requestId

db.Execute(updateSQL)

If Err.Number <> 0 Then
    db.RollbackTrans
    Response.Write "<script>alert('문서 수정 중 오류가 발생했습니다: " & Err.Description & "'); history.back();</script>"
    Response.End
End If

' 2. ApprovalLogs 테이블의 모든 결재 기록 초기화
Dim resetLogsSQL
resetLogsSQL = "UPDATE " & dbSchema & ".ApprovalLogs SET " & _
               "status = '대기', " & _
               "approved_at = NULL, " & _
               "comments = NULL " & _
               "WHERE target_table_name = 'VehicleRequests' " & _
               "AND target_id = " & requestId

db.Execute(resetLogsSQL)

If Err.Number <> 0 Then
    db.RollbackTrans
    Response.Write "<script>alert('결재선 초기화 중 오류가 발생했습니다: " & Err.Description & "'); history.back();</script>"
    Response.End
End If

' 트랜잭션 커밋
db.CommitTrans

' 성공 메시지와 함께 상세 페이지로 리다이렉트
Response.Write "<script>alert('문서가 수정되었습니다.'); location.href='approval_detail.asp?id=" & requestId & "&type=VehicleRequests';</script>"
Response.End
%> 