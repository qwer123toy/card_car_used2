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
    Response.Write "<script>alert('관리자 권한이 필요합니다.'); history.back();</script>"
    Response.End
End If

If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    Dim usageId, usageDate, amount, storeName, categoryId, purpose, cardId
    usageId     = PreventSQLInjection(Request.Form("usage_id"))
    usageDate   = PreventSQLInjection(Request.Form("usage_date"))
    amount      = PreventSQLInjection(Request.Form("amount"))
    storeName   = PreventSQLInjection(Request.Form("store_name"))
    categoryId  = PreventSQLInjection(Request.Form("category_id"))
    purpose     = PreventSQLInjection(Request.Form("purpose"))
    cardId      = PreventSQLInjection(Request.Form("card_id"))

    If usageId = "" Then
        Response.Write "<script>alert('수정 대상 ID가 없습니다.'); history.back();</script>"
        Response.End
    End If

    Dim sql
    sql = "UPDATE " & dbSchema & ".CardUsage SET " & _
          "usage_date = '" & usageDate & "', " & _
          "amount = " & amount & ", " & _
          "store_name = '" & storeName & "', " & _
          "expense_category_id = " & categoryId & ", " & _
          "purpose = '" & purpose & "', " & _
          "card_id = " & cardId & " " & _
          "WHERE usage_id = " & usageId

    On Error Resume Next
    db99.Execute sql

    If Err.Number <> 0 Then
    Response.Write "<script>alert('수정 중 오류 발생: " & Server.HTMLEncode(Err.Description) & "'); history.back();</script>"
    Response.End
Else
    Response.Write "<script>"
    Response.Write "alert('수정이 완료되었습니다.');"
    Response.Write "window.location.href='admin_card_usage_view.asp?id=" & Server.HTMLEncode(Server.URLEncode(usageId)) & "';"
    Response.Write "</script>"
    Response.End
End If

    On Error GoTo 0
ElseIf Request.QueryString("action") = "delete" Then
    ' 삭제 처리 (이미 구현된 부분)
End If
%>
