<%@ Language="VBScript" CodePage="65001" %>
<!-- METADATA TYPE="typelib" NAME="ADODB Type Library"
File="C:\Program Files\Common Files\System\ado\msado15.dll" -->
<% Option Explicit %>
<% Response.Expires=-1 %>
<!--#include file="db.asp"-->
<!--#include file="includes/functions.asp"-->

<% Response.CodePage = 65001
   Response.CharSet = "utf-8" %>
<%
' 세션 디버깅
' Response.Write("<!-- 세션ID: " & Session.SessionID & " | 인증상태: " & IsAuthenticated() & " -->")

' 이미 로그인한 경우 메인 페이지로 리디렉션
If IsAuthenticated() Then
    RedirectTo("/pages/dashboard.asp")
End If

Dim errorMsg : errorMsg = ""

' 로그인 처리
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    Dim userId, password, loginResult
    
    userId = PreventSQLInjection(Request.Form("user_id"))
    password = PreventSQLInjection(Request.Form("password"))
    
    If userId = "" Or password = "" Then
        errorMsg = "아이디와 비밀번호를 모두 입력해주세요."
    Else
        ' 하드코딩된 계정 인증 먼저 시도 (DB 오류 우회)
        If (userId = "admin" And password = "admin123") Or (userId = "user1" And password = "user123") Then
                ' 세션 설정
            Session("user_id") = userId
            
            ' 관리자 여부 설정
            If userId = "admin" Then
                Session("name") = "관리자"
                Session("department_id") = 1
                Session("is_admin") = "Y"
                
                ' 관리자인 경우 관리자 대시보드로 이동
                Response.Redirect("/pages/admin/admin_dashboard.asp")
                Response.End
            Else
                Session("name") = "사용자"
                Session("department_id") = 2
                Session("is_admin") = "N"
            End If
            
            ' 로그 기록
            LogActivity userId, "로그인", "하드코딩된 계정으로 로그인"
            
            ' 페이지 이동
            Response.Redirect("/pages/dashboard.asp")
            Response.End
        End If
        
        ' 데이터베이스 인증 처리 - 하드코딩된 계정이 아닌 경우 시도
        If dbConnected Then  ' 데이터베이스 연결 상태 확인
            Dim loginSQL, loginRS
            
            ' SQL Injection 방지를 위한 파라미터화된 쿼리 사용
            loginSQL = "SELECT user_id, name, department_id FROM Users WHERE user_id = ? AND password = ?"
            
            On Error Resume Next
            
            ' 준비된 명령 객체 생성
            Dim cmd
            Set cmd = Server.CreateObject("ADODB.Command")
            cmd.ActiveConnection = db99
            cmd.CommandText = loginSQL
            
            ' 파라미터 추가
            cmd.Parameters.Append cmd.CreateParameter("@user_id", 200, 1, 30, userId)
            cmd.Parameters.Append cmd.CreateParameter("@password", 200, 1, 50, password)
            
            ' 명령 실행
            Set loginRS = cmd.Execute()
                
            If Err.Number <> 0 Then
                ' 데이터베이스 오류가 발생하면 로그 기록
                LogActivity "SYSTEM", "오류", "로그인 중 DB 오류 발생: " & Err.Description
                errorMsg = "데이터베이스 연결 오류가 발생했습니다. 관리자에게 문의하세요."
            Else
                ' DB 오류 없이 정상적으로 실행된 경우
                If Not loginRS.EOF Then
                    ' 로그인 성공
                    Session("user_id") = loginRS("user_id")
                    Session("name") = loginRS("name")
                    Session("department_id") = loginRS("department_id")
                    
                    ' 관리자 체크
                    If userId = "admin" Then
                        Session("is_admin") = "Y"
                        
                        ' 로그 기록
                        LogActivity userId, "로그인", "관리자 로그인"
                        
                        ' 관리자인 경우 관리자 대시보드로 이동
                        Response.Redirect("/pages/admin/admin_dashboard.asp")
                        Response.End
                    Else
                        Session("is_admin") = "N"
                        
                        ' 로그 기록
                        LogActivity userId, "로그인", "사용자 로그인"
                        
                        ' 일반 사용자는 일반 대시보드로 이동
                        Response.Redirect("/pages/dashboard.asp")
                        Response.End
                    End If
                Else
                    errorMsg = "아이디 또는 비밀번호가 일치하지 않습니다."
                End If
                
                ' 레코드셋 닫기
                If Not loginRS Is Nothing Then
                    If loginRS.State <> 0 Then loginRS.Close
                    Set loginRS = Nothing
                End If
            End If
            On Error GoTo 0
        Else
            ' DB 연결이 안 된 경우
            errorMsg = "시스템 오류가 발생했습니다. 잠시 후 다시 시도해주세요."
        End If
    End If
End If

' 로컬 디버깅용 코드 (필요 시)
Sub ForceLogin(id)
    Session("user_id") = id
    
    ' 관리자 여부 설정
    If id = "admin" Then
        Session("name") = "관리자"
        Session("department_id") = 1
        Session("is_admin") = "Y"
        
        ' 관리자 대시보드로 이동
        Response.Redirect("/pages/admin/admin_dashboard.asp")
    Else
        Session("name") = "사용자"
        Session("department_id") = 2
        Session("is_admin") = "N"
        
        ' 일반 대시보드로 이동
        Response.Redirect("/pages/dashboard.asp")
    End If
    
    ' 로그 기록
    LogActivity id, "로그인", "개발 모드 강제 로그인"
End Sub
%>
<!DOCTYPE html>
<html>
    <head>
        <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>카드 지출 결의/개인차량 이용 관리</title>
    <style>
    /* 기본 스타일 */
    body {
        font-family: 'Pretendard', 'Noto Sans KR', sans-serif;
        line-height: 1.6;
        color: #333;
        background-color: #f5f5f5;
        margin: 0;
        padding: 0;
    }
    
    .container {
        width: 100%;
        max-width: 1200px;
        margin: 0 auto;
        padding: 0 15px;
    }
    
    /* 헤더 */
    header {
        background-color: #fff;
        box-shadow: 0 2px 4px rgba(0, 0, 0, 0.1);
        padding: 1rem 0;
    }
    
    header .container {
        display: flex;
        justify-content: space-between;
        align-items: center;
    }
    
    .logo a {
        font-size: 1.5rem;
        font-weight: bold;
        color: #333;
        text-decoration: none;
    }
    
    nav ul {
        display: flex;
        list-style: none;
        margin: 0;
        padding: 0;
    }
    
    nav ul li {
        margin-left: 1.5rem;
    }
    
    nav ul li a {
        text-decoration: none;
        color: #555;
        font-weight: 500;
    }
    
    /* 로그인 컨테이너 */
    .login-container {
        display: flex;
        justify-content: center;
        align-items: center;
        min-height: calc(100vh - 200px);
    }
    
    .shadcn-card {
        background-color: #fff;
        border-radius: 8px;
        box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
        width: 100%;
        padding: 2rem;
    }
    
    .shadcn-card-header {
        margin-bottom: 1.5rem;
    }
    
    .shadcn-card-title {
        margin: 0 0 0.5rem;
        font-size: 1.5rem;
        font-weight: 600;
    }
    
    .shadcn-card-description {
        margin: 0;
        color: #666;
    }
    
    .shadcn-card-content {
        margin-bottom: 1.5rem;
    }
    
    .form-group {
        margin-bottom: 1rem;
    }
    
    .shadcn-input-label {
        display: block;
        margin-bottom: 0.5rem;
        font-weight: 500;
    }
    
    .shadcn-input {
        width: 100%;
        padding: 0.75rem;
        border: 1px solid #ddd;
        border-radius: 4px;
        font-size: 1rem;
    }
    
    .shadcn-card-footer {
        display: flex;
        justify-content: flex-end;
        gap: 1rem;
    }
    
    .shadcn-btn {
        padding: 0.75rem 1.5rem;
        border-radius: 4px;
        font-weight: 500;
        cursor: pointer;
        text-decoration: none;
        display: inline-block;
        text-align: center;
    }
    
    .shadcn-btn-primary {
        background-color: #0070f3;
        color: white;
        border: none;
    }
    
    .shadcn-btn-outline {
        background-color: transparent;
        color: #0070f3;
        border: 1px solid #0070f3;
    }
    
    .shadcn-alert {
        padding: 1rem;
        border-radius: 4px;
        margin-bottom: 1rem;
    }
    
    .shadcn-alert-error {
        background-color: #fee2e2;
        border: 1px solid #ef4444;
    }
    
    .shadcn-alert-title {
        display: block;
        font-weight: 600;
        margin-bottom: 0.25rem;
    }
    
    /* 푸터 */
    footer {
        background-color: #fff;
        padding: 1rem 0;
        text-align: center;
        margin-top: 2rem;
        border-top: 1px solid #eee;
    }
    </style>
    <!-- 추가 스타일 -->
    <style>
    .error-message {
        color: #e11d48;
        font-size: 0.9rem;
        margin-top: 0.5rem;
        margin-bottom: 1rem;
        padding: 0.5rem;
        background-color: #fee2e2;
        border-radius: 4px;
    }
    </style>
    </head>
    <body>
    <header>
        <div class="container">
            <div class="logo">
                <a href="/index.asp">카드지출/차량이용관리</a>
            </div>
            <nav>
                <ul>
                    <%
                    If Session("user_id") <> "" Then
                    %>
                        <li><a href="/pages/card_usage.asp">카드사용 내역</a></li>
                        <li><a href="/pages/vehicle_request.asp">개인차량이용 신청</a></li>
                        <%
                        If Session("is_admin") = "Y" Then
                        %>
                            <li><a href="/pages/admin/admin_dashboard.asp">관리자</a></li>
                        <%
                        End If
                        %>
                        <li><a href="/pages/logout.asp">로그아웃</a></li>
                    <%
                    Else
                    %>
                        <li><a href="/index.asp">로그인</a></li>
                        <li><a href="/pages/register.asp">회원가입</a></li>
                    <%
                    End If
                    %>
                </ul>
            </nav>
        </div>
    </header>
    <main class="container">
<div class="login-container">
            <div class="shadcn-card" style="max-width: 400px;">
        <div class="shadcn-card-header">
            <h2 class="shadcn-card-title">로그인</h2>
                    <p class="shadcn-card-description">계정 정보를 입력하여 로그인하세요.</p>
                </div>
                <form method="post" action="index.asp">
                    <div class="shadcn-card-content">
                        <div class="form-group">
                            <label for="user_id">아이디</label>
                            <input type="text" id="user_id" name="user_id" class="shadcn-input" required>
                        </div>
                        <div class="form-group">
                            <label for="password">비밀번호</label>
                            <input type="password" id="password" name="password" class="shadcn-input" required>
        </div>
        
        <% If errorMsg <> "" Then %>
                        <div class="error-message">
                            <%= errorMsg %>
        </div>
        <% End If %>
        
                        <button type="submit" class="shadcn-btn shadcn-btn-primary" style="width: 100%;">로그인</button>
                </div>
                </form>
                <div class="shadcn-card-footer" style="text-align: center;">
                    <p>계정이 없으신가요? <a href="/pages/register.asp">회원가입</a></p>
                </div>
                </div>
        </div>
    </main>
    <footer>
        <div class="container">
            <p>&copy; <%= Year(Now) %> 카드 지출 결의/개인차량 이용 관리 시스템</p>
    </div>
    </footer>
</body>
</html>