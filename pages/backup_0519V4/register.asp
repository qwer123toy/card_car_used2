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

' 부서 목록 가져오기
Dim deptRS, deptSQL
deptSQL = "SELECT department_id, name FROM " & dbSchema & ".Department ORDER BY name"
On Error Resume Next
Set deptRS = db.Execute(deptSQL)

' 부서 테이블이 없는 경우 대체 테이블 시도
If Err.Number <> 0 Then
    Err.Clear
    deptSQL = "SELECT department_id, name FROM " & dbSchema & ".Departments ORDER BY name"
    Set deptRS = db.Execute(deptSQL)
End If

' 직급 목록 가져오기
Dim gradeRS, gradeSQL
gradeSQL = "SELECT job_grade_id, name FROM " & dbSchema & ".job_grade ORDER BY job_grade_id"
On Error Resume Next
Set gradeRS = db.Execute(gradeSQL)

' 직급 테이블이 없는 경우 대체 테이블 시도
If Err.Number <> 0 Then
    Err.Clear
    gradeSQL = "SELECT job_grade_id, name FROM " & dbSchema & ".job_grades ORDER BY job_grade_id"
    Set gradeRS = db.Execute(gradeSQL)
End If



Dim errorMsg, successMsg
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    Dim userId, password, confirmPassword, name, email, departmentId, job_gradeId
    
    userId = PreventSQLInjection(Request.Form("user_id"))
    password = PreventSQLInjection(Request.Form("password"))
    confirmPassword = PreventSQLInjection(Request.Form("confirm_password"))
    name = PreventSQLInjection(Request.Form("name"))
    email = PreventSQLInjection(Request.Form("email"))
    departmentId = PreventSQLInjection(Request.Form("department_id"))
    job_gradeId = PreventSQLInjection(Request.Form("job_grade_id"))
    
    If userId = "" Or password = "" Or confirmPassword = "" Or name = "" Then
        errorMsg = "필수 항목을 모두 입력해주세요."
    ElseIf password <> confirmPassword Then
        errorMsg = "비밀번호가 일치하지 않습니다."
    ElseIf Len(password) < 6 Then
        errorMsg = "비밀번호는 최소 6자리 이상이어야 합니다."
    Else
        ' 아이디 중복 확인
        Dim checkSQL, checkRS
        checkSQL = "SELECT 1 FROM Users WHERE user_id = ?"
        
        Dim cmd
        Set cmd = Server.CreateObject("ADODB.Command")
        cmd.ActiveConnection = db
        cmd.CommandText = checkSQL
        cmd.Parameters.Append cmd.CreateParameter("@user_id", 200, 1, 30, userId)
        
        Set checkRS = cmd.Execute()
        
        If Not checkRS.EOF Then
            errorMsg = "이미 사용 중인 아이디입니다."
        Else
            ' 사용자 등록
            Dim SQL
            
            ' 이메일, 부서 ID, 직급 ID 처리
            Dim emailValue, deptIdValue, gradeIdValue
            
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
            
            If job_gradeId = "" Then
                gradeIdValue = "NULL"
            Else
                gradeIdValue = job_gradeId
            End If
            
            SQL = "INSERT INTO Users (user_id, password, name, email, department_id, job_grade, created_at) " & _
                  "VALUES (?, ?, ?, " & emailValue & ", " & deptIdValue & ", " & gradeIdValue & ", GETDATE())"
            
            Set cmd = Server.CreateObject("ADODB.Command")
            cmd.ActiveConnection = db
            cmd.CommandText = SQL
            cmd.Parameters.Append cmd.CreateParameter("@user_id", 200, 1, 30, userId)
            cmd.Parameters.Append cmd.CreateParameter("@password", 200, 1, 100, password)
            cmd.Parameters.Append cmd.CreateParameter("@name", 200, 1, 50, name)
            
            On Error Resume Next
            cmd.Execute
            
            If Err.Number <> 0 Then
                errorMsg = "등록 중 오류가 발생했습니다: " & Err.Description
            Else
                successMsg = "회원가입이 완료되었습니다. 로그인해주세요."
            End If
            On Error GoTo 0
        End If
        
        If IsObject(checkRS) Then
            If checkRS.State = 1 Then
                checkRS.Close
            End If
            Set checkRS = Nothing
        End If
    End If
End If
%>
<!--#include file="../includes/header.asp"-->

<style>
.container { max-width: 900px; }
.card {
    border: none;
    box-shadow: 0 4px 6px rgba(0,0,0,0.1);
    border-radius: 12px;
    margin-top: 2rem;
}
.card-header {
    background-color: #f8f9fa;
    border-bottom: 1px solid #eee;
    padding: 1.5rem;
    border-radius: 12px 12px 0 0 !important;
}
.card-body { padding: 2rem; }
.form-group { margin-bottom: 1.5rem; }
.form-control {
    border-radius: 6px;
    border: 1px solid #ced4da;
    padding: 0.75rem 1rem;
    font-size: 1rem;
}
.form-control:focus {
    border-color: #80bdff;
    box-shadow: 0 0 0 0.2rem rgba(0,123,255,.25);
}
label { 
    font-weight: 600;
    margin-bottom: 0.5rem;
    color: #495057;
}
.btn {
    padding: 0.75rem 1.5rem;
    font-weight: 600;
    border-radius: 6px;
    transition: all 0.2s;
}
.btn-primary {
    background-color: #007bff;
    border-color: #007bff;
}
.btn-primary:hover {
    background-color: #0069d9;
    border-color: #0062cc;
    transform: translateY(-1px);
}
.btn-secondary {
    background-color: #6c757d;
    border-color: #6c757d;
}
.btn-secondary:hover {
    background-color: #5a6268;
    border-color: #545b62;
    transform: translateY(-1px);
}
.required-mark {
    color: #dc3545;
    margin-left: 2px;
}
.form-text {
    font-size: 0.875rem;
    color: #6c757d;
    margin-top: 0.25rem;
}
.alert {
    border-radius: 6px;
    padding: 1rem 1.25rem;
    margin-bottom: 1.5rem;
}
.password-section {
    background-color: #f8f9fa;
    padding: 1.5rem;
    border-radius: 8px;
    margin-bottom: 1.5rem;
}
</style>

<div class="container">
    <div class="row justify-content-center">
        <div class="col-md-8">
            <div class="card">
                <div class="card-header">
                    <h2 class="text-center mb-0">회원가입</h2>
                </div>
                <div class="card-body">
                    <% If errorMsg <> "" Then %>
                    <div class="alert alert-danger" role="alert">
                        <%= errorMsg %>
                    </div>
                    <% End If %>
                    
                    <% If successMsg <> "" Then %>
                    <div class="alert alert-success" role="alert">
                        <%= successMsg %>
                    </div>
                    <% End If %>
                    
                    <form method="post" action="register.asp">
                        <div class="form-group">
                            <label for="user_id">아이디<span class="required-mark">*</span></label>
                            <input type="text" class="form-control" id="user_id" name="user_id" required>
                        </div>
                        
                        <div class="password-section">
                            <div class="form-group">
                                <label for="password">비밀번호<span class="required-mark">*</span></label>
                                <input type="password" class="form-control" id="password" name="password" required>
                                <small class="form-text text-muted">비밀번호는 최소 6자리 이상이어야 합니다.</small>
                            </div>
                            
                            <div class="form-group mb-0">
                                <label for="confirm_password">비밀번호 확인<span class="required-mark">*</span></label>
                                <input type="password" class="form-control" id="confirm_password" name="confirm_password" required>
                            </div>
                        </div>
                        
                        <div class="form-group">
                            <label for="name">이름<span class="required-mark">*</span></label>
                            <input type="text" class="form-control" id="name" name="name" required>
                        </div>
                        
                        <div class="form-group">
                            <label for="email">이메일</label>
                            <input type="email" class="form-control" id="email" name="email">
                        </div>
                        
                        <div class="form-group">
                            <label for="department_id">부서</label>
                            <select class="form-control" id="department_id" name="department_id">
                                <option value="">선택해주세요</option>
                                <% 
                                If Not deptRS.BOF Then
                                    deptRS.MoveFirst
                                    Do Until deptRS.EOF 
                                %>
                                    <option value="<%= deptRS("department_id") %>"><%= deptRS("name") %></option>
                                <% 
                                    deptRS.MoveNext
                                    Loop
                                End If
                                %>
                            </select>
                        </div>
                        
                        <div class="form-group">
                            <label for="job_grade_id">직급</label>
                            <select class="form-control" id="job_grade_id" name="job_grade_id">
                                <option value="">선택해주세요</option>
                                <% 
                                If Not gradeRS.BOF Then
                                    gradeRS.MoveFirst
                                    Do Until gradeRS.EOF 
                                %>
                                    <option value="<%= gradeRS("job_grade_id") %>"><%= gradeRS("name") %></option>
                                <% 
                                    gradeRS.MoveNext
                                    Loop
                                End If
                                %>
                            </select>
                        </div>
                        
                        <div class="form-group text-center mt-4">
                            <button type="submit" class="btn btn-primary">회원가입</button>
                            <a href="index.asp" class="btn btn-secondary ml-2">취소</a>
                        </div>
                    </form>
                </div>
            </div>
        </div>
    </div>
</div>

<!--#include file="../includes/footer.asp"-->