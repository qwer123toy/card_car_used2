<%@ Language="VBScript" CodePage="65001" %>
<% 
Response.CodePage = 65001
Response.CharSet = "utf-8"
%>

<!--#include file="../db.asp"-->
<!--#include file="../includes/functions.asp"-->
<%
' 이미 로그인한 경우 메인 페이지로 리디렉션
If IsAuthenticated() Then
    RedirectTo("/contents/card_car_used/pages/dashboard.asp")
End If

' 부서 목록을 직접 정의
Dim deptRS
Set deptRS = Server.CreateObject("ADODB.Recordset")
deptRS.Fields.Append "department_id", 3 ' adInteger
deptRS.Fields.Append "name", 200, 100 ' adVarChar
deptRS.Open
deptRS.AddNew
deptRS("department_id") = 1
deptRS("name") = "관리부"
deptRS.Update
deptRS.AddNew
deptRS("department_id") = 2
deptRS("name") = "영업부"
deptRS.Update
deptRS.AddNew
deptRS("department_id") = 3
deptRS("name") = "기술부"
deptRS.Update
deptRS.MoveFirst

Dim errorMsg, successMsg
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    Dim userId, password, confirmPassword, name, email, departmentId
    
    userId = PreventSQLInjection(Request.Form("user_id"))
    password = PreventSQLInjection(Request.Form("password"))
    confirmPassword = PreventSQLInjection(Request.Form("confirm_password"))
    name = PreventSQLInjection(Request.Form("name"))
    email = PreventSQLInjection(Request.Form("email"))
    departmentId = PreventSQLInjection(Request.Form("department_id"))
    
    
    If userId = "" Or password = "" Or confirmPassword = "" Or name = "" Then
        errorMsg = "필수 항목을 모두 입력해주세요."
    ElseIf password <> confirmPassword Then
        errorMsg = "비밀번호가 일치하지 않습니다."
    ElseIf Len(password) < 6 Then
        errorMsg = "비밀번호는 최소 6자리 이상이어야 합니다."
    Else
        ' 아이디 중복 확인은 우선 생략
        ' successMsg = "회원가입이 완료되었습니다. 로그인해주세요."
        
        ' 실제 운영 환경에서는 아래 주석을 해제하여 사용
        
        Dim checkSQL, checkRS
        checkSQL = "SELECT 1 FROM Users WHERE user_id = '" & userId & "'"
        Set checkRS = db99.Execute(checkSQL)
        
        If Not checkRS.EOF Then
            errorMsg = "이미 사용 중인 아이디입니다."
        Else
            ' 사용자 등록
            Dim SQL
            
            ' 이메일 및 부서 ID 처리
            Dim emailValue, deptIdValue
            
            If email = "" Then
                emailValue = "NULL"
            Else
                emailValue = "'" & email & "'"
            End If
            
            If departmentId = "" Then
                deptIdValue = "NULL"
            Else
                deptIdValue = departmentId
            End If
            
            SQL = "INSERT INTO Users (user_id, password, name, email, department_id, created_at) " & _
                  "VALUES ('" & userId & "', '" & password & "', '" & name & "', " & _
                  emailValue & ", " & deptIdValue & ", GETDATE())"
            
            On Error Resume Next
            db99.Execute SQL
            
            If Err.Number <> 0 Then
                errorMsg = "등록 중 오류가 발생했습니다: " & Err.Description
            Else
                successMsg = "회원가입이 완료되었습니다. 로그인해주세요."
            End If
            On Error GoTo 0
        End If
        
        checkRS.Close
        Set checkRS = Nothing
    End If
End If
%>
<!--#include file="../includes/header.asp"-->

<div class="login-container">
    <div class="shadcn-card" style="max-width: 600px; margin: 50px auto;">
        <div class="shadcn-card-header">
            <h2 class="shadcn-card-title">회원가입</h2>
            <p class="shadcn-card-description">카드 지출 결의 및 개인차량 이용 내력 관리 시스템에 가입하세요.</p>
        </div>
        
        <% If errorMsg <> "" Then %>
        <div class="shadcn-alert shadcn-alert-error">
            <div>
                <span class="shadcn-alert-title">오류</span>
                <span class="shadcn-alert-description"><%= errorMsg %></span>
            </div>
        </div>
        <% End If %>
        
        <% If successMsg <> "" Then %>
        <div class="shadcn-alert shadcn-alert-success">
            <div>
                <span class="shadcn-alert-title">성공</span>
                <span class="shadcn-alert-description"><%= successMsg %></span>
            </div>
        </div>
        <script>
            setTimeout(function() {
                window.location.href = "/contents/card_car_used/index.asp";
            }, 3000);
        </script>
        <% End If %>
        
        <div class="shadcn-card-content">
            <form id="registerForm" method="post" action="/contents/card_car_used/pages/register.asp">
                <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(250px, 1fr)); gap: 16px;">
                    <div class="form-group">
                        <label class="shadcn-input-label" for="user_id">아이디 <span class="required">*</span></label>
                        <input class="shadcn-input" type="text" id="user_id" name="user_id" placeholder="아이디를 입력하세요">
                    </div>
                    
                    <div class="form-group">
                        <label class="shadcn-input-label" for="name">이름 <span class="required">*</span></label>
                        <input class="shadcn-input" type="text" id="name" name="name" placeholder="이름을 입력하세요">
                    </div>
                    
                    <div class="form-group">
                        <label class="shadcn-input-label" for="password">비밀번호 <span class="required">*</span></label>
                        <input class="shadcn-input" type="password" id="password" name="password" placeholder="비밀번호를 입력하세요">
                    </div>
                    
                    <div class="form-group">
                        <label class="shadcn-input-label" for="confirm_password">비밀번호 확인 <span class="required">*</span></label>
                        <input class="shadcn-input" type="password" id="confirm_password" name="confirm_password" placeholder="비밀번호를 다시 입력하세요">
                    </div>
                    
                    <div class="form-group">
                        <label class="shadcn-input-label" for="email">이메일</label>
                        <input class="shadcn-input" type="email" id="email" name="email" placeholder="이메일을 입력하세요">
                    </div>
                    
                    <div class="form-group">
                        <label class="shadcn-input-label" for="department_id">부서</label>
                        <select class="shadcn-select" id="department_id" name="department_id">
                            <option value="">선택하세요</option>
                            <% 
                            Do While Not deptRS.EOF 
                            %>
                            <option value="<%= deptRS("department_id") %>"><%= deptRS("name") %></option>
                            <% 
                                deptRS.MoveNext
                                Loop 
                            %>
                        </select>
                    </div>
                </div>
                
                <div class="shadcn-card-footer" style="margin-top: 1.5rem;">
                    <button type="submit" class="shadcn-btn shadcn-btn-primary">회원가입</button>
                    <a href="/contents/card_car_used/index.asp" class="shadcn-btn shadcn-btn-outline">로그인 페이지로</a>
                </div>
            </form>
        </div>
    </div>
</div>

<!--#include file="../includes/footer.asp"-->