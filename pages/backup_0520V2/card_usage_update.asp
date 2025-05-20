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
Dim usageId, cardId, usageDate, amount, purpose, storeName
usageId = Request.Form("usage_id")
cardId = Request.Form("card_id")
usageDate = Request.Form("usage_date")
amount = Request.Form("amount")
purpose = Request.Form("purpose")
storeName = Request.Form("store_name")

' 필수 값 검증
If usageId = "" Or cardId = "" Or usageDate = "" Or amount = "" Or purpose = "" Or storeName = "" Then
    Response.Write "<script>alert('필수 항목이 누락되었습니다.'); history.back();</script>"
    Response.End
End If

' 현재 사용자가 문서 작성자인지 확인
Dim checkOwnerSQL, checkOwnerRS
checkOwnerSQL = "SELECT user_id, approval_status FROM " & dbSchema & ".CardUsage WHERE usage_id = " & usageId
Set checkOwnerRS = db.Execute(checkOwnerSQL)

If checkOwnerRS.EOF Then
    Response.Write "<script>alert('존재하지 않는 문서입니다.'); history.back();</script>"
    Response.End
End If

If checkOwnerRS("user_id") <> Session("user_id") Then
    Response.Write "<script>alert('수정 권한이 없습니다.'); history.back();</script>"
    Response.End
End If

If checkOwnerRS("approval_status") <> "반려" And checkOwnerRS("approval_status") <> "대기" Then
    Response.Write "<script>alert('대기 또는 반려 상태의 문서만 수정할 수 있습니다.'); history.back();</script>"
    Response.End
End If

' 트랜잭션 시작
db.BeginTrans

' 1. CardUsage 테이블 업데이트
Dim updateCardSQL
updateCardSQL = "UPDATE " & dbSchema & ".CardUsage SET " & _
                "card_id = " & cardId & ", " & _
                "usage_date = '" & usageDate & "', " & _
                "amount = " & amount & ", " & _
                "purpose = '" & PreventSQLInjection(purpose) & "', " & _
                "store_name = '" & PreventSQLInjection(storeName) & "', " & _
                "approval_status = '대기' " & _
                "WHERE usage_id = " & usageId

db.Execute(updateCardSQL)

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
               "WHERE target_table_name = 'CardUsage' " & _
               "AND target_id = " & usageId

db.Execute(resetLogsSQL)

If Err.Number <> 0 Then
    db.RollbackTrans
    Response.Write "<script>alert('결재선 초기화 중 오류가 발생했습니다: " & Err.Description & "'); history.back();</script>"
    Response.End
End If

' 트랜잭션 커밋
db.CommitTrans

' 성공 메시지와 함께 상세 페이지로 리다이렉트
Response.Write "<script>alert('문서가 수정되었습니다.'); location.href='approval_detail.asp?id=" & usageId & "';</script>"
Response.End
%> 