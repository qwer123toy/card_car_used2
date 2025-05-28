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
    Dim request_id, request_date, start_date, end_date, title, destination, purpose
    request_id = Request.Form("request_id")
    request_date = Request.Form("request_date")
    start_date = Request.Form("start_date")
    end_date = Request.Form("end_date")
    title = Replace(Request.Form("title"), "'", "''")
    destination = Replace(Request.Form("destination"), "'", "''")
    purpose = Replace(Request.Form("purpose"), "'", "''")

    Dim updateSQL
    updateSQL = "UPDATE VehicleRequests SET " & _
                "request_date = '" & request_date & "', " & _
                "start_date = '" & start_date & "', " & _
                "end_date = '" & end_date & "', " & _
                "title = '" & title & "', " & _
                "destination = '" & destination & "', " & _
                "purpose = '" & purpose & "' " & _
                "WHERE request_id = " & request_id

    On Error Resume Next
    db.Execute(updateSQL)

    If Err.Number <> 0 Then
        Response.Write "<script>alert('수정 중 오류가 발생했습니다.'); history.back();</script>"
        Response.End
    Else
        Response.Write "<script>alert('수정이 완료되었습니다.'); location.href='admin_vehicle_requests.asp';</script>"
    End If
End If
%>
