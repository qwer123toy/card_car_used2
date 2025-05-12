<!--#include file="../includes/connection.asp"-->
<!--#include file="../includes/functions.asp"-->
<%
' 이미 로그인한 경우 메인 페이지로 리디렉션
If IsAuthenticated() Then
    RedirectTo("dashboard.asp")
End If

' 회원가입 처리
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    Dim userId, password, confirmPassword, name, departmentId, jobGrade, errorMsg, successMsg, SQL, rs
    
    userId = PreventSQLInjection(Request.Form("user_id"))
    password = PreventSQLInjection(Request.Form("password"))
    confirmPassword = PreventSQLInjection(Request.Form("confirm_password"))
    name = PreventSQLInjection(Request.Form("name"))
    departmentId = PreventSQLInjection(Request.Form("department_id"))
    jobGrade = PreventSQLInjection(Request.Form("job_grade"))
    
    ' 입력값 검증
    If userId = "" Or password = "" Or confirmPassword = "" Or name = "" Or departmentId = "" Or jobGrade = "" Then
        errorMsg = "모든 필드를 입력해주세요."
    ElseIf password <> confirmPassword Then
        errorMsg = "비밀번호가 일치하지 않습니다."
    Else
        ' 아이디 중복 확인
        SQL = "SELECT 1 FROM Users WHERE user_id = '" & userId & "'"
        Set rs = dbConn.Execute(SQL)
        
        If Not rs.EOF Then
            errorMsg = "이미 사용 중인 아이디입니다."
        Else
            ' 사용자 등록
            SQL = "INSERT INTO Users (user_id, password, name, department_id, job_grade) VALUES ('" & _
                  userId & "', '" & password & "', '" & name & "', " & departmentId & ", '" & jobGrade & "')"
            
            On Error Resume Next
            dbConn.Execute SQL
            
            If Err.Number <> 0 Then
                errorMsg = "회원가입 중 오류가 발생했습니다: " & Err.Description
            Else
                successMsg = "회원가입이 완료되었습니다. 로그인해주세요."
                
                ' 활동 로그 기록
                LogActivity userId, "회원가입", "새 사용자 등록"
            End If
            On Error GoTo 0
        End If
        
        rs.Close
        Set rs = Nothing
    End If
End If

' 부서 목록 가져오기
Dim departmentSQL, departmentRS
departmentSQL = "SELECT department_id, name FROM Department ORDER BY name"
Set departmentRS = dbConn.Execute(departmentSQL)
%>
<!--#include file="../includes/header.asp"-->

<div class="register-container">
    <div class="shadcn-card" style="max-width: 550px; margin: 50px auto;">
        <div class="shadcn-card-header">
            <h2 class="shadcn-card-title">회원가입</h2>
            <p class="shadcn-card-description">카드 지출 결의 및 개인차량 이용 내력 관리 시스템에 가입합니다.</p>
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
        <% End If %>
        
        <div class="shadcn-card-content">
            <form id="registerForm" method="post" action="register.asp" onsubmit="return validateForm('registerForm', registerRules)">
                <div class="form-group">
                    <label class="shadcn-input-label" for="user_id">아이디</label>
                    <input class="shadcn-input" type="text" id="user_id" name="user_id" placeholder="아이디를 입력하세요">
                </div>
                
                <div class="form-group">
                    <label class="shadcn-input-label" for="password">비밀번호</label>
                    <input class="shadcn-input" type="password" id="password" name="password" placeholder="비밀번호를 입력하세요">
                </div>
                
                <div class="form-group">
                    <label class="shadcn-input-label" for="confirm_password">비밀번호 확인</label>
                    <input class="shadcn-input" type="password" id="confirm_password" name="confirm_password" placeholder="비밀번호를 다시 입력하세요">
                </div>
                
                <div class="form-group">
                    <label class="shadcn-input-label" for="name">이름</label>
                    <input class="shadcn-input" type="text" id="name" name="name" placeholder="이름을 입력하세요">
                </div>
                
                <div class="form-group">
                    <label class="shadcn-input-label" for="department_id">부서</label>
                    <select class="shadcn-select" id="department_id" name="department_id">
                        <option value="">부서를 선택하세요</option>
                        <% 
                        If Not departmentRS.EOF Then
                            Do While Not departmentRS.EOF 
                        %>
                            <option value="<%= departmentRS("department_id") %>"><%= departmentRS("name") %></option>
                        <% 
                                departmentRS.MoveNext
                            Loop
                        End If
                        departmentRS.Close
                        %>
                    </select>
                </div>
                
                <div class="form-group">
                    <label class="shadcn-input-label" for="job_grade">직급</label>
                    <input class="shadcn-input" type="text" id="job_grade" name="job_grade" placeholder="직급을 입력하세요">
                </div>
                
                <div class="shadcn-card-footer" style="margin-top: 1.5rem;">
                    <button type="submit" class="shadcn-btn shadcn-btn-primary">회원가입</button>
                    <a href="../index.asp" class="shadcn-btn shadcn-btn-outline">로그인 화면으로</a>
                </div>
            </form>
        </div>
    </div>
</div>

<script>
    const registerRules = {
        user_id: {
            required: true,
            minLength: 4,
            message: '아이디는 4자 이상 입력해주세요.'
        },
        password: {
            required: true,
            minLength: 6,
            message: '비밀번호는 6자 이상 입력해주세요.'
        },
        confirm_password: {
            required: true,
            message: '비밀번호 확인을 입력해주세요.'
        },
        name: {
            required: true,
            message: '이름을 입력해주세요.'
        },
        department_id: {
            required: true,
            message: '부서를 선택해주세요.'
        },
        job_grade: {
            required: true,
            message: '직급을 입력해주세요.'
        }
    };
    
    // 비밀번호 일치 확인
    document.getElementById('registerForm').addEventListener('submit', function(e) {
        const password = document.getElementById('password').value;
        const confirmPassword = document.getElementById('confirm_password').value;
        
        if (password !== confirmPassword) {
            e.preventDefault();
            showError(document.getElementById('confirm_password'), '비밀번호가 일치하지 않습니다.');
        }
    });
</script>

<!--#include file="../includes/footer.asp"-->