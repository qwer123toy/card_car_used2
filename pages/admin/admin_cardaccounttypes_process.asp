<%@ Language="VBScript" CodePage="65001" %>
<% 
Response.CodePage = 65001
Response.CharSet = "utf-8"
%>

<!--#include file="../../db.asp"-->
<!--#include file="../../includes/functions.asp"-->

<%
If Not IsAuthenticated() Then
    RedirectTo("../../index.asp")
End If

If Not IsAdmin() Then
    Response.Write("<script>alert('관리자 권한이 필요합니다.'); window.location.href='../dashboard.asp';</script>")
    Response.End
End If

If Request.ServerVariables("REQUEST_METHOD") <> "POST" Then
    Response.Write("<script>alert('잘못된 접근입니다.'); window.location.href='admin_cardaccounttypes.asp';</script>")
    Response.End
End If

Dim action, typeName, typeId
action = Request.Form("action")
typeName = PreventSQLInjection(Request.Form("type_name"))

If typeName = "" Then
    Response.Write("<script>alert('유형명을 입력해주세요.'); history.back();</script>")
    Response.End
End If

On Error Resume Next

If action = "add" Then
    Dim addSQL
    addSQL = "INSERT INTO " & dbSchema & ".CardAccountTypes (type_name) VALUES ('" & typeName & "')"
    db99.Execute(addSQL)

    If Err.Number <> 0 Then
        Response.Write("<script>alert('계정 유형 추가 중 오류가 발생했습니다: " & Server.HTMLEncode(Err.Description) & "'); history.back();</script>")
    Else
        LogActivity Session("user_id"), "카드계정유형추가", "카드 계정 유형 추가 (이름: " & typeName & ")"
        Response.Write("<script>alert('계정 유형이 추가되었습니다.'); window.location.href='admin_cardaccounttypes.asp';</script>")
    End If

ElseIf action = "edit" Then
    typeId = PreventSQLInjection(Request.Form("account_type_id"))

    If typeId = "" Then
        Response.Write("<script>alert('수정할 항목 ID가 누락되었습니다.'); window.location.href='admin_cardaccounttypes.asp';</script>")
        Response.End
    End If

    Dim editSQL
    editSQL = "UPDATE " & dbSchema & ".CardAccountTypes SET type_name = '" & typeName & "' WHERE account_type_id = " & typeId
    db99.Execute(editSQL)

    If Err.Number <> 0 Then
        Response.Write("<script>alert('계정 유형 수정 중 오류가 발생했습니다: " & Server.HTMLEncode(Err.Description) & "'); history.back();</script>")
    Else
        LogActivity Session("user_id"), "카드계정유형수정", "카드 계정 유형 수정 (ID: " & typeId & ", 이름: " & typeName & ")"
        Response.Write("<script>alert('계정 유형이 수정되었습니다.'); window.location.href='admin_cardaccounttypes.asp';</script>")
    End If

Else
    Response.Write("<script>alert('잘못된 요청입니다.'); window.location.href='admin_cardaccounttypes.asp';</script>")
End If

On Error GoTo 0
%>

