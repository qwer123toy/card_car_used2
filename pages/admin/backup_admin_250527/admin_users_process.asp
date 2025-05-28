<%@ Language="VBScript" CodePage="65001" %>
<% 
Response.CodePage = 65001
Response.CharSet = "utf-8"
%>

<!--#include file="../../db.asp"-->
<!--#include file="../../includes/functions.asp"-->

<%
' 로그인 및 관리자 체크
If Not IsAuthenticated() Then
    RedirectTo("../../index.asp")
End If

If Not IsAdmin() Then
    Response.Write("<script>alert('관리자 권한이 필요합니다.'); window.location.href='../dashboard.asp';</script>")
    Response.End
End If

' POST 요청 확인
If Request.ServerVariables("REQUEST_METHOD") <> "POST" Then
    Response.Write("<script>alert('잘못된 접근입니다.'); window.location.href='admin_users.asp';</script>")
    Response.End
End If

' 폼 값 수집
Dim userId, name, email, deptId, gradeId
userId = PreventSQLInjection(Request.Form("user_id"))
name = PreventSQLInjection(Request.Form("name"))
email = PreventSQLInjection(Request.Form("email"))
deptId = PreventSQLInjection(Request.Form("department_id"))
gradeId = PreventSQLInjection(Request.Form("job_grade_id"))

' 필수값 확인
If userId = "" Or name = "" Then
    Response.Write("<script>alert('사용자 ID 또는 이름이 누락되었습니다.'); history.back();</script>")
    Response.End
End If

' 업데이트 수행
On Error Resume Next

Dim sql
sql = "UPDATE " & dbSchema & ".Users SET " & _
        "name = '" & name & "', " & _
        "email = " & IIf(email = "", "NULL", "'" & email & "'") & ", " & _
        "department_id = " & IIf(IsNumeric(deptId), deptId, "NULL") & ", " & _
        "job_grade = " & IIf(IsNumeric(gradeId), gradeId, "NULL") & " " & _
      "WHERE user_id = '" & userId & "'"

On Error Resume Next
db99.Execute sql

If Err.Number <> 0 Then
    Response.Write("<script>alert('사용자 수정 중 오류 발생: " & Server.HTMLEncode(Err.Description) & "'); history.back();</script>")
    Response.End
Else
    LogActivity Session("user_id"), "사용자수정", "사용자 정보 수정 (ID: " & userId & ", 이름: " & name & ")"
    Response.Write("<script>alert('사용자 정보가 수정되었습니다.'); window.location.href='admin_user_view.asp?id=" & userId & "';</script>")
    Response.End
End If

On Error GoTo 0
%>
