<!--#include file="../includes/connection.asp"-->
<!--#include file="../includes/functions.asp"-->
<%
' 활동 로그 기록 (로그아웃 전에 기록)
If IsAuthenticated() Then
    LogActivity Session("user_id"), "로그아웃", "사용자 로그아웃"
End If

' 세션 변수 제거
Session.Contents.RemoveAll()
Session.Abandon

' 로그인 페이지로 리디렉션
RedirectTo("../index.asp")
%> 