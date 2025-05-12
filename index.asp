<!--#include file="includes/connection.asp"-->
<!--#include file="includes/functions.asp"-->
<%
' 이미 로그인한 경우 메인 페이지로 리디렉션
If IsAuthenticated() Then
    RedirectTo("pages/dashboard.asp")
End If

' 로그인 처리
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    Dim userId, password, errorMsg, SQL, rs
    
    userId = PreventSQLInjection(Request.Form("user_id"))
    password = PreventSQLInjection(Request.Form("password"))
    
    If userId = "" Or password = "" Then
        errorMsg = "아이디와 비밀번호를 모두 입력해주세요."
    Else
        ' 사용자 확인
        SQL = "SELECT * FROM Users WHERE user_id = '" & userId & "'"
        Set rs = dbConn.Execute(SQL)
        
        If rs.EOF Then
            errorMsg = "존재하지 않는 아이디입니다."
        Else
            ' 비밀번호 확인 (실제 환경에서는 해시된 비밀번호를 비교해야 함)
            If rs("password") = password Then
                ' 세션 설정
                Session("user_id") = rs("user_id")
                Session("name") = rs("name")
                Session("department_id") = rs("department_id")
                
                ' 관리자 여부 확인
                Dim sqlAdmin, rsAdmin
                sqlAdmin = "SELECT 1 FROM Administrators WHERE user_id = '" & userId & "'"
                Set rsAdmin = dbConn.Execute(sqlAdmin)
                
                If Not rsAdmin.EOF Then
                    Session("is_admin") = "Y"
                Else
                    Session("is_admin") = "N"
                End If
                
                ' 로그인 기록
                LogActivity userId, "로그인", "사용자 로그인 성공"
                
                ' 페이지 이동
                RedirectTo("pages/dashboard.asp")
            Else
                errorMsg = "비밀번호가 일치하지 않습니다."
            End If
        End If
        
        rs.Close
        Set rs = Nothing
    End If
End If
%>
<!--#include file="includes/header.asp"-->

<div class="login-container">
    <div class="shadcn-card" style="max-width: 450px; margin: 80px auto;">
        <div class="shadcn-card-header">
            <h2 class="shadcn-card-title">로그인</h2>
            <p class="shadcn-card-description">카드 지출 결의 및 개인차량 이용 내력 관리 시스템에 오신 것을 환영합니다.</p>
        </div>
        
        <% If errorMsg <> "" Then %>
        <div class="shadcn-alert shadcn-alert-error">
            <div>
                <span class="shadcn-alert-title">오류</span>
                <span class="shadcn-alert-description"><%= errorMsg %></span>
            </div>
        </div>
        <% End If %>
        
        <div class="shadcn-card-content">
            <form id="loginForm" method="post" action="index.asp" onsubmit="return validateForm('loginForm', loginRules)">
                <div class="form-group">
                    <label class="shadcn-input-label" for="user_id">아이디</label>
                    <input class="shadcn-input" type="text" id="user_id" name="user_id" placeholder="아이디를 입력하세요">
                </div>
                
                <div class="form-group">
                    <label class="shadcn-input-label" for="password">비밀번호</label>
                    <input class="shadcn-input" type="password" id="password" name="password" placeholder="비밀번호를 입력하세요">
                </div>
                
                <div class="shadcn-card-footer" style="margin-top: 1.5rem;">
                    <button type="submit" class="shadcn-btn shadcn-btn-primary">로그인</button>
                    <a href="pages/register.asp" class="shadcn-btn shadcn-btn-outline">회원가입</a>
                </div>
            </form>
        </div>
    </div>
</div>

<script>
    const loginRules = {
        user_id: {
            required: true,
            message: '아이디를 입력해주세요.'
        },
        password: {
            required: true,
            message: '비밀번호를 입력해주세요.'
        }
    };
</script>

<!--#include file="includes/footer.asp"--> 