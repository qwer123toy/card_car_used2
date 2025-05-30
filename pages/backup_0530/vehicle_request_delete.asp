<%@ Language="VBScript" CodePage="65001" %>
<% 
Response.CodePage = 65001
Response.CharSet = "utf-8"
%>

<!--#include file="../db.asp"-->
<!--#include file="../includes/functions.asp"-->
<%
' 로그인 체크
If Not IsAuthenticated() Then
    RedirectTo("../index.asp")
End If

On Error Resume Next

' URL 파라미터에서 신청서 ID 추출
Dim requestId, errorMsg, successMsg
requestId = PreventSQLInjection(Request.QueryString("id"))

If requestId = "" Then
    errorMsg = "잘못된 접근입니다. 신청서 ID가 필요합니다."
    Response.Redirect("vehicle_request.asp")
End If

' 사용자 권한 확인: 본인 신청서이고 상태가 '작성중'인지 확인
Dim checkCmd, checkRS
Set checkCmd = Server.CreateObject("ADODB.Command")
checkCmd.ActiveConnection = db
checkCmd.CommandText = "SELECT user_id, approval_status FROM VehicleRequests WHERE request_id = ? AND is_deleted = 0"
checkCmd.Parameters.Append checkCmd.CreateParameter("@request_id", 3, 1, , CLng(requestId))

Set checkRS = checkCmd.Execute()

If Err.Number <> 0 Or checkRS.EOF Then
    errorMsg = "요청하신 신청서를 찾을 수 없습니다."
    RedirectTo("vehicle_request.asp")
ElseIf checkRS("user_id") <> Session("user_id") Then
    errorMsg = "본인이 작성한 신청서만 삭제할 수 있습니다."
    RedirectTo("vehicle_request.asp")
ElseIf checkRS("approval_status") <> "작성중" And checkRS("approval_status") <> "반려" And checkRS("approval_status") <> "대기" Then
    errorMsg = "작성중, 반려, 대기 상태의 신청서만 삭제할 수 있습니다."
    RedirectTo("vehicle_request.asp")
End If

If Not checkRS Is Nothing Then
    If checkRS.State = 1 Then ' adStateOpen
        checkRS.Close
    End If
    Set checkRS = Nothing
End If

If errorMsg = "" Then
    ' 신청서 삭제 (is_deleted 필드 업데이트)
    Dim cmd
    Set cmd = Server.CreateObject("ADODB.Command")
    cmd.ActiveConnection = db
    cmd.CommandText = "UPDATE VehicleRequests SET is_deleted = 1 WHERE request_id = ?"
    cmd.Parameters.Append cmd.CreateParameter("@request_id", 3, 1, , CLng(requestId))
    
    ' 명령 실행
    cmd.Execute
    
    If Err.Number <> 0 Then
        errorMsg = "신청서 삭제 중 오류가 발생했습니다: " & Err.Description
    Else
        ' 활동 로그 기록
        LogActivity Session("user_id"), "차량이용신청삭제", "개인차량 이용 신청서 삭제 (ID: " & requestId & ")"
        
        ' 목록 페이지로 리디렉션 (성공 메시지 포함)
        successMsg = "신청서가 성공적으로 삭제되었습니다."
        Session("success_msg") = successMsg
        RedirectTo("vehicle_request.asp")
    End If
End If

On Error GoTo 0

' 오류 발생 시 목록 페이지로 리디렉션
If errorMsg <> "" Then
    Session("error_msg") = errorMsg
    RedirectTo("vehicle_request.asp")
End If
%> 