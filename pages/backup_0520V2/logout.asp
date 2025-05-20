<%@ Language="VBScript" CodePage="65001" %>
<% 
Response.CodePage = 65001
Response.CharSet = "utf-8"
%>
<!--#include file="../db.asp"-->
<!--#include file="../includes/functions.asp"-->
<%
' 활동 로그 기록 (로그아웃 전에 기록)
If IsAuthenticated() Then
    LogActivity Session("user_id"), "로그아웃", "사용자 로그아웃"
End If

' 세션 변수 제거
Session.Abandon

' 로그 기록
LogActivity Session("user_id"), "로그아웃", "사용자 로그아웃"

' 로그인 페이지로 이동
RedirectTo("/contents/card_car_used/index.asp")
%> 