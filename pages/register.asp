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
gradeSQL = "SELECT job_grade_id, name FROM " & dbSchema & ".job_grade ORDER BY sort_order"
On Error Resume Next
Set gradeRS = db.Execute(gradeSQL)




Dim errorMsg, successMsg
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    Dim userId, password, confirmPassword, name, email, phone, departmentId, job_gradeId
    
    userId = PreventSQLInjection(Request.Form("user_id"))
    password = PreventSQLInjection(Request.Form("password"))
    confirmPassword = PreventSQLInjection(Request.Form("confirm_password"))
    name = PreventSQLInjection(Request.Form("name"))
    email = PreventSQLInjection(Request.Form("email"))
    departmentId = PreventSQLInjection(Request.Form("department_id"))
    job_gradeId = PreventSQLInjection(Request.Form("job_grade_id"))
    phone = PreventSQLInjection(Request.Form("phone"))
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
            Dim emailValue, deptIdValue, gradeIdValue, phoneValue
            If phone = "" Then
                phoneValue = "NULL"
            Else
                phoneValue = "'" & phone & "'"
            End If
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
            
            SQL = "INSERT INTO Users (user_id, password, name, email, phone, department_id, job_grade, created_at) " & _
            "VALUES (?, ?, ?, ?, ?, ?, ?, GETDATE())"
            
            Set cmd = Server.CreateObject("ADODB.Command")
            cmd.ActiveConnection = db
            cmd.CommandText = SQL
            cmd.CommandType = 1 ' adCmdText
            
            cmd.Parameters.Append cmd.CreateParameter("@user_id", 200, 1, 30, userId)
            cmd.Parameters.Append cmd.CreateParameter("@password", 200, 1, 100, password)
            cmd.Parameters.Append cmd.CreateParameter("@name", 200, 1, 50, name)
            cmd.Parameters.Append cmd.CreateParameter("@email", 200, 1, 100, email)
            cmd.Parameters.Append cmd.CreateParameter("@phone", 200, 1, 20, phone)
            If departmentId = "" Then departmentId = Null
            cmd.Parameters.Append cmd.CreateParameter("@department_id", 3, 1, , departmentId) ' adInteger
            If job_gradeId = "" Then job_gradeId = Null
            cmd.Parameters.Append cmd.CreateParameter("@job_grade", 3, 1, , job_gradeId) ' adInteger
            
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
.container {
    max-width: 900px;
    margin: 2rem auto;
    padding: 0 1rem;
}

.page-header {
    display: flex;
    justify-content: space-between;
    align-items: center;
    margin-bottom: 2rem;
    padding: 1rem;
    background: white;
    border-radius: 12px;
    box-shadow: 0 2px 4px rgba(0,0,0,0.05);
}

.page-title {
    font-size: 1.5rem;
    font-weight: 600;
    color: #2C3E50;
    margin: 0;
}

.card {
    background: #fff;
    border: none;
    border-radius: 16px;
    box-shadow: 0 0 20px rgba(0,0,0,0.05);
    margin-bottom: 2rem;
    overflow: hidden;
}

.card-header {
    background: linear-gradient(to right, #4A90E2, #5A9EEA);
    border-bottom: none;
    padding: 1.5rem;
}

.card-header h5 {
    color: #fff;
    font-weight: 600;
    margin: 0;
    font-size: 1.25rem;
}

.card-body {
    padding: 2rem;
}

.form-group {
    margin-bottom: 1.5rem;
    position: relative;
}

.form-label {
    font-weight: 600;
    color: #2C3E50;
    margin-bottom: 0.5rem;
    display: block;
}

.form-control {
    border: 2px solid #E9ECEF;
    border-radius: 8px;
    padding: 0.875rem 1rem;
    font-size: 1rem;
    transition: all 0.2s ease;
    width: 100%;
}

.form-control:focus {
    border-color: #4A90E2;
    box-shadow: 0 0 0 4px rgba(74,144,226,0.1);
}

.form-select {
    border: 2px solid #E9ECEF;
    border-radius: 8px;
    padding: 0.875rem 1rem;
    font-size: 1rem;
    width: 100%;
}

.required-mark {
    color: #E74C3C;
    margin-left: 4px;
}

.btn {
    padding: 0.875rem 1.5rem;
    font-weight: 600;
    border-radius: 8px;
    transition: all 0.2s ease;
}

.btn-primary {
    background: linear-gradient(to right, #4A90E2, #5A9EEA);
    border: none;
    color: white;
}

.btn-primary:hover {
    transform: translateY(-2px);
    box-shadow: 0 4px 12px rgba(74,144,226,0.2);
}

.btn-secondary {
    background: #F8FAFC;
    border: 2px solid #E9ECEF;
    color: #2C3E50;
}

.btn-secondary:hover {
    background: #E9ECEF;
    transform: translateY(-2px);
}

.alert {
    border: none;
    border-radius: 12px;
    padding: 1.25rem 1.5rem;
    margin-bottom: 2rem;
    font-weight: 500;
}

.alert-danger {
    background: #FDF1F1;
    color: #E74C3C;
}

.alert-success {
    background: #EDF9F0;
    color: #2ECC71;
}

.row {
    display: flex;
    flex-wrap: wrap;
    margin: -0.75rem;
}

.col-md-6 {
    flex: 0 0 50%;
    max-width: 50%;
    padding: 0.75rem;
}

@media (max-width: 768px) {
    .col-md-6 {
        flex: 0 0 100%;
        max-width: 100%;
    }
}
</style>

<script>
    function formatPhoneNumber(input) {
        let value = input.value.replace(/\D/g, ""); // 숫자 이외 제거
        if (value.length > 11) value = value.slice(0, 11); // 최대 11자리 제한
    
        let result = "";
    
        if (value.startsWith("02")) {
            // 서울번호(예외 케이스)
            if (value.length > 2) {
                result += value.substr(0, 2) + "-";
                if (value.length > 5) {
                    result += value.substr(2, 3) + "-" + value.substr(5);
                } else {
                    result += value.substr(2);
                }
            } else {
                result += value;
            }
        } else {
            // 일반적인 휴대폰 번호
            if (value.length > 3) {
                result += value.substr(0, 3) + "-";
                if (value.length > 7) {
                    result += value.substr(3, 4) + "-" + value.substr(7);
                } else {
                    result += value.substr(3);
                }
            } else {
                result = value;
            }
        }
    
        input.value = result;
    }
    </script>

<div class="container">
    <div class="page-header">
        <h2 class="page-title">회원가입</h2>
    </div>

    <% If errorMsg <> "" Then %>
        <div class="alert alert-danger" role="alert">
            <i class="fas fa-exclamation-circle me-2"></i><%= errorMsg %>
        </div>
    <% End If %>
    
    <% If successMsg <> "" Then %>
        <div class="alert alert-success" role="alert">
            <i class="fas fa-check-circle me-2"></i><%= successMsg %>
        </div>
    <% End If %>

    <div class="card">
        <div class="card-header">
            <h5 class="card-title">사용자 정보 입력</h5>
        </div>
        <div class="card-body">
            <form method="post" action="register.asp">
                <div class="row">
                    <div class="col-md-6">
                        <div class="form-group">
                            <label class="form-label">아이디<span class="required-mark">*</span></label>
                            <input type="text" name="user_id" class="form-control" required>
                        </div>
                    </div>
                    <div class="col-md-6">
                        <div class="form-group">
                            <label class="form-label">이름<span class="required-mark">*</span></label>
                            <input type="text" name="name" class="form-control" required>
                        </div>
                    </div>
                    <div class="col-md-6">
                        <div class="form-group">
                            <label class="form-label">비밀번호<span class="required-mark">*</span></label>
                            <input type="password" name="password" class="form-control" required>
                        </div>
                    </div>
                    <div class="col-md-6">
                        <div class="form-group">
                            <label class="form-label">비밀번호 확인<span class="required-mark">*</span></label>
                            <input type="password" name="confirm_password" class="form-control" required>
                        </div>
                    </div>
                    <div class="col-md-6">
                        <div class="form-group">
                            <label class="form-label">이메일</label>
                            <input type="email" name="email" class="form-control">
                        </div>
                    </div>
                    <div class="col-md-6">
                        <div class="form-group">
                            <label class="form-label">전화번호</label>
                            <input type="text" name="phone" class="form-control" id="phone" maxlength="13" oninput="formatPhoneNumber(this);">
                        </div>
                    </div>
                    <div class="col-md-6">
                        <div class="form-group">
                            <label class="form-label">부서</label>
                            <select name="department_id" class="form-select">
                                <option value="">선택해주세요</option>
                                <% 
                                If Not deptRS.EOF Then
                                    Do While Not deptRS.EOF 
                                %>
                                    <option value="<%= deptRS("department_id") %>">
                                        <%= deptRS("name") %>
                                    </option>
                                <% 
                                        deptRS.MoveNext
                                    Loop
                                End If
                                %>
                            </select>
                        </div>
                    </div>
                    <div class="col-md-6">
                        <div class="form-group">
                            <label class="form-label">직급</label>
                            <select name="job_grade_id" class="form-select">
                                <option value="">선택해주세요</option>
                                <% 
                                If Not gradeRS.EOF Then
                                    Do While Not gradeRS.EOF 
                                %>
                                    <option value="<%= gradeRS("job_grade_id") %>">
                                        <%= gradeRS("name") %>
                                    </option>
                                <% 
                                        gradeRS.MoveNext
                                    Loop
                                End If
                                %>
                            </select>
                        </div>
                    </div>
                </div>

                <div class="text-center mt-4">
                    <button type="submit" class="btn btn-primary me-2">
                        <i class="fas fa-user-plus me-1"></i> 회원가입
                    </button>
                    <a href="../index.asp" class="btn btn-secondary">
                        <i class="fas fa-times me-1"></i> 취소
                    </a>
                </div>
            </form>
        </div>
    </div>
</div>

<!--#include file="../includes/footer.asp"-->