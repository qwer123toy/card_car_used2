<%@ Language="VBScript" CodePage="65001" %>
<%
Response.CodePage = 65001
Response.CharSet = "utf-8"
%>

<!--#include file="../db.asp"-->
<!--#include file="../includes/functions.asp"-->

<%
' 로그인 상태 확인
If IsAuthenticated() Then
    RedirectTo("/card_car_used/pages/dashboard.asp")
End If

' 로그인 처리
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    Dim username, password, errorMsg
    username = Trim(Request.Form("username"))
    password = Trim(Request.Form("password"))
    
    If username = "" Or password = "" Then
        errorMsg = "사용자명과 비밀번호를 모두 입력해주세요."
    Else
        ' 하드코딩된 계정 인증 먼저 시도 (DB 오류 우회)
        If (username = "admin" And password = "admin123") Or (username = "user1" And password = "user123") Then
            ' 세션 설정
            Session("user_id") = username
            
            ' 관리자 여부 설정
            If username = "admin" Then
                Session("name") = "관리자"
                Session("department_id") = 1
                Session("is_admin") = "Y"
            Else
                Session("name") = "사용자"
                Session("department_id") = 2
                Session("is_admin") = "N"
            End If
            
            ' 로그 기록
            LogActivity username, "로그인", "하드코딩된 계정으로 로그인"
            
            ' 페이지 이동
            RedirectTo("/card_car_used/pages/dashboard.asp")
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
            cmd.Parameters.Append cmd.CreateParameter("@user_id", 200, 1, 30, username)
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
                    If username = "admin" Then
                        Session("is_admin") = "Y"
                        
                        ' 로그 기록
                        LogActivity username, "로그인", "관리자 로그인"
                    Else
                        Session("is_admin") = "N"
                        
                        ' 로그 기록
                        LogActivity username, "로그인", "사용자 로그인"
                    End If
                    
                    ' 페이지 이동
                    RedirectTo("/card_car_used/pages/dashboard.asp")
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
%>

<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>카드 지출 결의/개인차량 이용 관리</title>
    <link rel="stylesheet" href="../css/shadcn.css">
</head>
<body>
    <div class="shadcn-container">
        <div class="shadcn-card" style="max-width: 400px; margin: 50px auto;">
            <div class="shadcn-card-header">
                <h2 class="shadcn-card-title">로그인</h2>
            </div>
            
            <% If errorMsg <> "" Then %>
            <div class="shadcn-alert shadcn-alert-destructive">
                <%= errorMsg %>
            </div>
            <% End If %>
            
            <form id="loginForm" method="post" action="/card_car_used/index.asp">
                <div class="shadcn-card-content">
                    <div class="form-group">
                        <label for="username">사용자명</label>
                        <input type="text" id="username" name="username" class="shadcn-input" required>
                    </div>
                    <div class="form-group">
                        <label for="password">비밀번호</label>
                        <input type="password" id="password" name="password" class="shadcn-input" required>
                    </div>
                    <button type="submit" class="shadcn-btn shadcn-btn-primary">로그인</button>
                </div>
            </form>
            <div class="shadcn-card-footer" style="text-align: center;">
                <a href="/card_car_used/pages/register.asp" class="shadcn-btn shadcn-btn-outline">회원가입</a>
            </div>
        </div>
    </div>
</body>
</html>


<script>
    const loginRules = {
        user_id: {
            required: true,
            message: '?꾩씠?붾? ?낅젰?댁＜?몄슂.'
        },
        password: {
            required: true,
            message: '鍮꾨?踰덊샇瑜??낅젰?댁＜?몄슂.'
        }
    };
</script>

<!--#include file="includes/footer.asp"--> 
