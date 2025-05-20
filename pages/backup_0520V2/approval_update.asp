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

If action = "reject" Then
    ' 반려 처리
    ' 1. 현재 결재 단계를 반려로 변경하고 나머지 결재선은 초기화
    Dim updateLogSQL
    updateLogSQL = "UPDATE " & dbSchema & ".ApprovalLogs SET " & _
                   "status = CASE " & _
                   "   WHEN approval_step <= " & approvalStep & " THEN '반려' " & _
                   "   ELSE '대기' END, " & _
                   "processed_at = CASE " & _
                   "   WHEN approval_step <= " & approvalStep & " THEN NULL " & _
                   "   ELSE NULL END, " & _
                   "comment = CASE " & _
                   "   WHEN approval_step = " & approvalStep & " THEN '" & PreventSQLInjection(comment) & "' " & _
                   "   ELSE NULL END " & _
                   "WHERE target_table_name = 'CardUsage' " & _
                   "AND target_id = " & usageId

    db.Execute(updateLogSQL)
    
    If Err.Number <> 0 Then
        db.RollbackTrans
        Response.Write "<script>alert('결재 상태 업데이트 중 오류가 발생했습니다: " & Err.Description & "'); history.back();</script>"
        Response.End
    End If
    
    ' 2. CardUsage 상태를 '반려'로 변경
    Dim updateCardSQL
    updateCardSQL = "UPDATE " & dbSchema & ".CardUsage SET " & _
                   "approval_status = '반려' " & _
                   "WHERE usage_id = " & usageId
    
    db.Execute(updateCardSQL)
    
Else
    ' 승인 처리
    ' 1. 현재 결재 단계를 승인으로 변경
    Dim updateCurrentSQL
    updateCurrentSQL = "UPDATE " & dbSchema & ".ApprovalLogs SET " & _
                      "status = '승인', " & _
                      "comment = '" & PreventSQLInjection(comment) & "', " & _
                      "processed_at = GETDATE() " & _
                      "WHERE target_table_name = 'CardUsage' " & _
                      "AND target_id = " & usageId & " " & _
                      "AND approval_step = " & approvalStep
    
    db.Execute(updateCurrentSQL)
    
    If Err.Number <> 0 Then
        db.RollbackTrans
        Response.Write "<script>alert('결재 상태 업데이트 중 오류가 발생했습니다: " & Err.Description & "'); history.back();</script>"
        Response.End
    End If
    
    ' 2. 다음 결재자 확인
    Dim nextApproverSQL, nextApproverRS
    nextApproverSQL = "SELECT TOP 1 approval_step FROM " & dbSchema & ".ApprovalLogs " & _
                      "WHERE target_table_name = 'CardUsage' " & _
                      "AND target_id = " & usageId & " " & _
                      "AND approval_step > " & approvalStep & " " & _
                      "AND status = '대기' " & _
                      "ORDER BY approval_step"
    
    Set nextApproverRS = db.Execute(nextApproverSQL)
    
    ' 3. CardUsage 상태 업데이트
    Dim cardStatus
    If nextApproverRS.EOF Then
        ' 다음 결재자가 없으면 최종 승인
        cardStatus = "승인"
    Else
        ' 다음 결재자가 있으면 진행중
        cardStatus = "진행중"
    End If
    
    Dim updateStatusSQL
    updateStatusSQL = "UPDATE " & dbSchema & ".CardUsage SET " & _
                     "approval_status = '" & cardStatus & "' " & _
                     "WHERE usage_id = " & usageId
    
    db.Execute(updateStatusSQL)
End If

If Err.Number <> 0 Then
    db.RollbackTrans
    Response.Write "<script>alert('처리 중 오류가 발생했습니다: " & Err.Description & "'); history.back();</script>"
    Response.End
End If

' 트랜잭션 커밋
db.CommitTrans

' 성공 메시지와 함께 대시보드로 리다이렉트
Response.Write "<script>alert('결재가 처리되었습니다.'); location.href='dashboard.asp';</script>"
Response.End
%> 