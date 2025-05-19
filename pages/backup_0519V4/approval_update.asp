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

' 파라미터 검증
Dim usageId, approvalStep, action, comment
usageId = Request.Form("usage_id")
approvalStep = Request.Form("approval_step")
action = Request.Form("action")
comment = Request.Form("comment")

If usageId = "" Or approvalStep = "" Or action = "" Then
    Response.Write "<script>alert('필수 파라미터가 누락되었습니다.'); history.back();</script>"
    Response.End
End If

' 현재 로그인한 사용자 정보
Dim currentUserId
currentUserId = Session("user_id")

' 트랜잭션 시작
db.BeginTrans

' 1. 현재 결재 단계의 상태 업데이트
Dim updateLogSQL
updateLogSQL = "UPDATE " & dbSchema & ".ApprovalLogs SET " & _
               "status = '" & IIf(action = "approve", "승인", "반려") & "', " & _
               "comment = '" & PreventSQLInjection(comment) & "', " & _
               "processed_at = GETDATE() " & _
               "WHERE target_table_name = 'CardUsage' " & _
               "AND target_id = " & usageId & " " & _
               "AND approval_step = " & approvalStep & " " & _
               "AND approver_id = '" & PreventSQLInjection(currentUserId) & "'"

db.Execute(updateLogSQL)

If Err.Number <> 0 Then
    db.RollbackTrans
    Response.Write "<script>alert('결재 상태 업데이트 중 오류가 발생했습니다: " & Err.Description & "'); history.back();</script>"
    Response.End
End If

' 2. CardUsage 테이블의 approval_status 업데이트
Dim cardStatus
If action = "reject" Then
    cardStatus = "반려"
Else
    ' 다음 결재자가 있는지 확인
    Dim nextApproverSQL, nextApproverRS
    nextApproverSQL = "SELECT TOP 1 approval_step FROM " & dbSchema & ".ApprovalLogs " & _
                      "WHERE target_table_name = 'CardUsage' " & _
                      "AND target_id = " & usageId & " " & _
                      "AND approval_step > " & approvalStep & " " & _
                      "ORDER BY approval_step"
    
    Set nextApproverRS = db.Execute(nextApproverSQL)
    
    If nextApproverRS.EOF Then
        ' 다음 결재자가 없으면 최종 승인
        cardStatus = "승인"
    Else
        ' 다음 결재자가 있으면 진행중
        cardStatus = "진행중"
    End If
End If

Dim updateCardSQL
updateCardSQL = "UPDATE " & dbSchema & ".CardUsage SET " & _
                "approval_status = '" & cardStatus & "' " & _
                "WHERE usage_id = " & usageId

db.Execute(updateCardSQL)

If Err.Number <> 0 Then
    db.RollbackTrans
    Response.Write "<script>alert('카드 사용내역 상태 업데이트 중 오류가 발생했습니다: " & Err.Description & "'); history.back();</script>"
    Response.End
End If

' 트랜잭션 커밋
db.CommitTrans

' 성공 메시지와 함께 대시보드로 리다이렉트
Response.Write "<script>alert('결재가 처리되었습니다.'); location.href='dashboard.asp';</script>"
Response.End
%> 