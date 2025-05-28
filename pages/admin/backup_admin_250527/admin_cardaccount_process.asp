<%@ Language="VBScript" CodePage="65001" %>
<%
Response.CodePage = 65001
Response.CharSet = "utf-8"
%>

<!--#include file="../../db.asp"-->
<!--#include file="../../includes/functions.asp"-->

<%
' 로그인 및 권한 확인
If Not IsAuthenticated() Then
    RedirectTo("../../index.asp")
End If

If Not IsAdmin() Then
    Response.Write("<script>alert('관리자 권한이 필요합니다.'); window.location.href='../dashboard.asp';</script>")
    Response.End
End If

' POST 확인
If Request.ServerVariables("REQUEST_METHOD") <> "POST" Then
    Response.Write("<script>alert('잘못된 접근입니다.'); window.location.href='admin_cardaccount.asp';</script>")
    Response.End
End If

' 폼 데이터
Dim action, accountName, issuer
action = Request.Form("action")
accountName = PreventSQLInjection(Request.Form("account_name"))
issuer = PreventSQLInjection(Request.Form("issuer"))

' 유효성 검사
If accountName = "" Or issuer = "" Then
    Response.Write("<script>alert('카드명과 카드회사를 모두 입력해주세요.'); history.back();</script>")
    Response.End
End If

On Error Resume Next

If action = "add" Then
    Dim addSQL
    addSQL = "INSERT INTO " & dbSchema & ".CardAccount (account_name, issuer) " & _
             "VALUES ('" & accountName & "', '" & issuer & "')"
    db99.Execute(addSQL)

    If Err.Number <> 0 Then
        Dim msg
        msg = Replace(Server.HTMLEncode(Err.Description), "'", "\\'")
        Response.Write("<script>alert('카드 계정 추가 중 오류가 발생했습니다: " & msg & "'); history.back();</script>")
        Response.End
    Else
        LogActivity Session("user_id"), "카드계정추가", "카드 계정 추가 (카드명: " & accountName & ", 카드회사: " & issuer & ")"
        Response.Write("<script>alert('카드 계정이 추가되었습니다.'); window.location.href='admin_cardaccount.asp';</script>")
    End If

Else
    Response.Write("<script>alert('잘못된 요청입니다.'); window.location.href='admin_cardaccount.asp';</script>")
End If

On Error GoTo 0
%>
