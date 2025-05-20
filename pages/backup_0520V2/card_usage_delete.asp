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

' URL 파라미터에서 ID 가져오기
Dim usageId, errorMsg, successMsg
usageId = PreventSQLInjection(Request.QueryString("id"))

If usageId = "" Then
    Session("error_msg") = "잘못된 접근입니다. 카드 사용 내역 ID가 필요합니다."
    Response.Redirect("card_usage.asp")
    Response.End
End If

' dbSchema가 설정되지 않은 경우를 대비해 기본값 설정
If Not(IsObject(dbSchema)) And (TypeName(dbSchema) <> "String" Or Len(dbSchema) = 0) Then
    dbSchema = "dbo"
End If

' 삭제할 데이터 존재 여부 확인
Dim checkSQL, checkRS
checkSQL = "SELECT usage_id FROM " & dbSchema & ".CardUsage WHERE usage_id = " & usageId
Set checkRS = db.Execute(checkSQL)

If Err.Number <> 0 Then
    Session("error_msg") = "데이터 조회 중 오류가 발생했습니다: " & Err.Description
    Response.Redirect("card_usage.asp")
    Response.End
End If

If checkRS.EOF Then
    Session("error_msg") = "삭제할 카드 사용 내역을 찾을 수 없습니다."
    Response.Redirect("card_usage.asp")
    Response.End
End If

' 카드 사용 내역 삭제
Dim deleteSQL
deleteSQL = "DELETE FROM " & dbSchema & ".CardUsage WHERE usage_id = " & usageId

On Error Resume Next
db.Execute(deleteSQL)

If Err.Number <> 0 Then
    Session("error_msg") = "카드 사용 내역 삭제 중 오류가 발생했습니다: " & Err.Description
    LogActivity Session("user_id"), "오류발생", "카드 사용 내역 삭제 오류: " & Err.Description
Else
    Session("success_msg") = "카드 사용 내역이 성공적으로 삭제되었습니다."
    LogActivity Session("user_id"), "카드사용삭제", "카드 사용 내역 삭제 (ID: " & usageId & ")"
End If

Response.Redirect("card_usage.asp")
%> 