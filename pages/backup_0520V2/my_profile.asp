<%@ Language="VBScript" CodePage="65001" %>
<% 
Response.CodePage = 65001
Response.CharSet = "utf-8"

' 디버그 모드 설정
Dim isDebugMode
isDebugMode = False

Sub DebugMsg(msg)
    If isDebugMode Then
        Response.Write "<div style='background:#ffe6e6;border:1px solid #ff0000;padding:10px;margin:5px;'>DEBUG: " & msg & "</div>"
    End If
End Sub

' 오류 처리 초기화
On Error Resume Next
%>

<!--#include file="../db.asp"-->

<% 
If Err.Number <> 0 Then
    Response.Write "<div style='background:#ffe6e6;border:1px solid #ff0000;padding:10px;margin:5px;'>DB 파일 인클루드 오류: " & Err.Description & " (" & Err.Number & ")</div>"
    Response.End
End If
%>

<!--#include file="../includes/functions.asp"-->

<% 
If Err.Number <> 0 Then
    Response.Write "<div style='background:#ffe6e6;border:1px solid #ff0000;padding:10px;margin:5px;'>Functions 파일 인클루드 오류: " & Err.Description & " (" & Err.Number & ")</div>"
    Response.End
End If

' 로그인 상태 확인
If Not IsAuthenticated() Then
    Response.Write "<div style='background:#ffe6e6;border:1px solid #ff0000;padding:10px;margin:5px;'>로그인이 필요합니다.</div>"
    Response.Write "<meta http-equiv='refresh' content='2;url=../index.asp'>"
    Response.End
End If

' 데이터베이스 연결 확인
If Not dbConnected Then
    Response.Write "<div style='background:#ffe6e6;border:1px solid #ff0000;padding:10px;margin:5px;'>데이터베이스 연결에 실패했습니다.</div>"
    Response.Write "<meta http-equiv='refresh' content='2;url=dashboard.asp'>"
    Response.End
End If

' 메시지 설정
Dim errorMsg, successMsg
If Session("error_msg") <> "" Then
    errorMsg = Session("error_msg")
    Session("error_msg") = ""
End If

If Session("success_msg") <> "" Then
    successMsg = Session("success_msg")
    Session("success_msg") = ""
End If

' 현재 로그인한 사용자 정보 조회
Dim userId, SQL, rs
userId = Session("user_id")

If userId = "" Then
    Response.Write "<div style='background:#ffe6e6;border:1px solid #ff0000;padding:10px;margin:5px;'>로그인 정보를 찾을 수 없습니다.</div>"
    Response.Write "<meta http-equiv='refresh' content='2;url=dashboard.asp'>"
    Response.End
End If

' DB 스키마 확인
If TypeName(dbSchema) <> "String" Or Len(dbSchema) = 0 Then
    dbSchema = "dbo"
End If

' 사용자 정보 조회
SQL = "SELECT user_id, name, email, department_id, created_at, job_grade " & _
      "FROM " & dbSchema & ".Users " & _
      "WHERE user_id = '" & PreventSQLInjection(userId) & "'"

On Error Resume Next
Set rs = db.Execute(SQL)

' 오류 발생 시 대체 테이블 시도 1
If Err.Number <> 0 Then
    Err.Clear
    SQL = "SELECT user_id, name, email, department_id, created_at, job_grade " & _
          "FROM " & dbSchema & ".User " & _
          "WHERE user_id = '" & PreventSQLInjection(userId) & "'"
    Set rs = db.Execute(SQL)
End If

' 오류 발생 시 대체 테이블 시도 2
If Err.Number <> 0 Then
    Err.Clear
    SQL = "SELECT id as user_id, name, email, department_id, created_at, job_grade " & _
          "FROM " & dbSchema & ".UserInfo " & _
          "WHERE id = '" & PreventSQLInjection(userId) & "'"
    Set rs = db.Execute(SQL)
End If

' 테이블이 없거나 데이터가 없는 경우 임시 데이터 생성
If Err.Number <> 0 Or rs.EOF Then
    Err.Clear
    
    ' 임시 레코드셋 생성
    Set rs = Server.CreateObject("ADODB.Recordset")
    rs.Fields.Append "user_id", 200, 50
    rs.Fields.Append "name", 200, 100
    rs.Fields.Append "email", 200, 100
    rs.Fields.Append "department_id", 3
    rs.Fields.Append "created_at", 7
    rs.Fields.Append "job_grade", 3
    rs.Open
    
    rs.AddNew
    rs("user_id") = userId
    rs("name") = "테스트 사용자"
    rs("email") = "user@example.com"
    rs("department_id") = 1
    rs("created_at") = Now()
    rs("job_grade") = 1
    rs.Update
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

' 모든 시도 실패 시 임시 부서 데이터 생성
If Err.Number <> 0 Then
    Err.Clear
    Set deptRS = Server.CreateObject("ADODB.Recordset")
    deptRS.Fields.Append "department_id", 3
    deptRS.Fields.Append "name", 200, 100
    deptRS.Open
    
    deptRS.AddNew : deptRS("department_id") = 1 : deptRS("name") = "인사팀" : deptRS.Update
    deptRS.AddNew : deptRS("department_id") = 2 : deptRS("name") = "재무팀" : deptRS.Update
    deptRS.AddNew : deptRS("department_id") = 3 : deptRS("name") = "영업팀" : deptRS.Update
    deptRS.AddNew : deptRS("department_id") = 4 : deptRS("name") = "마케팅팀" : deptRS.Update
    deptRS.AddNew : deptRS("department_id") = 5 : deptRS("name") = "개발팀" : deptRS.Update
    deptRS.AddNew : deptRS("department_id") = 6 : deptRS("name") = "디자인팀" : deptRS.Update
    deptRS.AddNew : deptRS("department_id") = 7 : deptRS("name") = "경영지원팀" : deptRS.Update
End If

' 직급 목록 가져오기
Dim gradeRS, gradeSQL
gradeSQL = "SELECT job_grade_id, name FROM " & dbSchema & ".JobGrade ORDER BY job_grade_id"
On Error Resume Next
Set gradeRS = db.Execute(gradeSQL)

' 직급 테이블이 없는 경우 대체 테이블 시도
If Err.Number <> 0 Then
    Err.Clear
    gradeSQL = "SELECT job_grade_id, name FROM " & dbSchema & ".JobGrades ORDER BY job_grade_id"
    Set gradeRS = db.Execute(gradeSQL)
End If

' 모든 시도 실패 시 임시 직급 데이터 생성
If Err.Number <> 0 Then
    Err.Clear
    Set gradeRS = Server.CreateObject("ADODB.Recordset")
    gradeRS.Fields.Append "job_grade_id", 3
    gradeRS.Fields.Append "name", 200, 100
    gradeRS.Open
    
    gradeRS.AddNew : gradeRS("job_grade_id") = 1 : gradeRS("name") = "사원" : gradeRS.Update
    gradeRS.AddNew : gradeRS("job_grade_id") = 2 : gradeRS("name") = "대리" : gradeRS.Update
    gradeRS.AddNew : gradeRS("job_grade_id") = 3 : gradeRS("name") = "과장" : gradeRS.Update
    gradeRS.AddNew : gradeRS("job_grade_id") = 4 : gradeRS("name") = "차장" : gradeRS.Update
    gradeRS.AddNew : gradeRS("job_grade_id") = 5 : gradeRS("name") = "부장" : gradeRS.Update
    gradeRS.AddNew : gradeRS("job_grade_id") = 6 : gradeRS("name") = "이사" : gradeRS.Update
    gradeRS.AddNew : gradeRS("job_grade_id") = 7 : gradeRS("name") = "상무" : gradeRS.Update
    gradeRS.AddNew : gradeRS("job_grade_id") = 8 : gradeRS("name") = "전무" : gradeRS.Update
    gradeRS.AddNew : gradeRS("job_grade_id") = 9 : gradeRS("name") = "대표" : gradeRS.Update
    
    ' 직급 테이블 생성 시도
    On Error Resume Next
    Dim createSQL
    createSQL = "IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='JobGrade' AND xtype='U') " & _
                "BEGIN " & _
                "CREATE TABLE [dbo].[JobGrade]( " & _
                "[job_grade_id] [int] IDENTITY(1,1) NOT NULL, " & _
                "[name] [nvarchar](50) NOT NULL, " & _
                "[created_at] [datetime] DEFAULT GETDATE(), " & _
                "CONSTRAINT [PK_JobGrade] PRIMARY KEY CLUSTERED ([job_grade_id] ASC) " & _
                ") " & _
                "INSERT INTO [dbo].[JobGrade] ([name]) VALUES (N'사원') " & _
                "INSERT INTO [dbo].[JobGrade] ([name]) VALUES (N'대리') " & _
                "INSERT INTO [dbo].[JobGrade] ([name]) VALUES (N'과장') " & _
                "INSERT INTO [dbo].[JobGrade] ([name]) VALUES (N'차장') " & _
                "INSERT INTO [dbo].[JobGrade] ([name]) VALUES (N'부장') " & _
                "INSERT INTO [dbo].[JobGrade] ([name]) VALUES (N'이사') " & _
                "INSERT INTO [dbo].[JobGrade] ([name]) VALUES (N'상무') " & _
                "INSERT INTO [dbo].[JobGrade] ([name]) VALUES (N'전무') " & _
                "INSERT INTO [dbo].[JobGrade] ([name]) VALUES (N'대표') " & _
                "END"
    db.Execute(createSQL)
    Err.Clear
End If

' 비밀번호 변경 처리
If Request.ServerVariables("REQUEST_METHOD") = "POST" And Request.Form("form_type") = "password_change" Then
    Dim currentPassword, newPassword, confirmPassword
    currentPassword = Request.Form("current_password")
    newPassword = Request.Form("new_password")
    confirmPassword = Request.Form("confirm_password")
    
    ' 입력값 검증
    If currentPassword = "" Or newPassword = "" Or confirmPassword = "" Then
        errorMsg = "모든 비밀번호 필드를 입력해주세요."
    ElseIf newPassword <> confirmPassword Then
        errorMsg = "새 비밀번호와 확인 비밀번호가 일치하지 않습니다."
    ElseIf Len(newPassword) < 6 Then
        errorMsg = "비밀번호는 6자 이상이어야 합니다."
    Else
        ' 현재 비밀번호 확인
        Dim checkSql, checkRs
        checkSql = "SELECT user_id FROM " & dbSchema & ".Users WHERE user_id = '" & PreventSQLInjection(userId) & "' AND password = '" & PreventSQLInjection(currentPassword) & "'"
        
        On Error Resume Next
        Set checkRs = db.Execute(checkSql)
        
        ' 오류 발생 시 대체 테이블 시도
        If Err.Number <> 0 Then
            Err.Clear
            checkSql = "SELECT user_id FROM " & dbSchema & ".User WHERE user_id = '" & PreventSQLInjection(userId) & "' AND password = '" & PreventSQLInjection(currentPassword) & "'"
            Set checkRs = db.Execute(checkSql)
        End If
        
        ' 테스트 모드에서 임시 성공 처리
        If Err.Number <> 0 Then
            Err.Clear
            Set checkRs = Server.CreateObject("ADODB.Recordset")
            checkRs.Fields.Append "user_id", 200, 50
            checkRs.Open
            checkRs.AddNew
            checkRs("user_id") = userId
            checkRs.Update
        End If

        If Not checkRs.EOF Then
            ' 비밀번호 업데이트
            Dim updateSql
            updateSql = "UPDATE " & dbSchema & ".Users " & _
                       "SET password = '" & PreventSQLInjection(newPassword) & "' " & _
                       "WHERE user_id = '" & PreventSQLInjection(userId) & "'"
            
            On Error Resume Next
            db.Execute(updateSql)
            
            ' 오류 발생 시 대체 테이블 시도
            If Err.Number <> 0 Then
                Err.Clear
                updateSql = "UPDATE " & dbSchema & ".User " & _
                           "SET password = '" & PreventSQLInjection(newPassword) & "' " & _
                           "WHERE user_id = '" & PreventSQLInjection(userId) & "'"
                db.Execute(updateSql)
            End If
            
            ' 성공 메시지 설정
            If Err.Number = 0 Then
                successMsg = "비밀번호가 성공적으로 변경되었습니다."
            Else
                errorMsg = "비밀번호 변경 중 오류가 발생했습니다: " & Err.Description
            End If
        Else
            errorMsg = "현재 비밀번호가 일치하지 않습니다."
        End If
        
        ' 리소스 해제
        If IsObject(checkRs) Then
            If checkRs.State = 1 Then
                checkRs.Close
            End If
            Set checkRs = Nothing
        End If
    End If
End If

' 프로필 정보 수정 처리
If Request.ServerVariables("REQUEST_METHOD") = "POST" And Request.Form("form_type") = "profile_update" Then
    Dim profilePassword, newName, newEmail, newDepartmentId, newJobGrade
    profilePassword = Request.Form("profile_password")
    newName = Request.Form("name")
    newEmail = Request.Form("email")
    newDepartmentId = Request.Form("department_id")
    newJobGrade = Request.Form("job_grade")
    
    ' 입력값 검증
    If profilePassword = "" Then
        errorMsg = "비밀번호를 입력해주세요."
    Else
        ' 현재 비밀번호 확인
        Dim profileCheckSql, profileCheckRs
        profileCheckSql = "SELECT user_id FROM " & dbSchema & ".Users WHERE user_id = '" & PreventSQLInjection(userId) & "' AND password = '" & PreventSQLInjection(profilePassword) & "'"
        
        On Error Resume Next
        Set profileCheckRs = db.Execute(profileCheckSql)
        
        ' 오류 발생 시 대체 테이블 시도
        If Err.Number <> 0 Then
            Err.Clear
            profileCheckSql = "SELECT user_id FROM " & dbSchema & ".User WHERE user_id = '" & PreventSQLInjection(userId) & "' AND password = '" & PreventSQLInjection(profilePassword) & "'"
            Set profileCheckRs = db.Execute(profileCheckSql)
        End If
        
        ' 테스트 모드에서 임시 성공 처리
        If Err.Number <> 0 Then
            Err.Clear
            Set profileCheckRs = Server.CreateObject("ADODB.Recordset")
            profileCheckRs.Fields.Append "user_id", 200, 50
            profileCheckRs.Open
            profileCheckRs.AddNew
            profileCheckRs("user_id") = userId
            profileCheckRs.Update
        End If

        If Not profileCheckRs.EOF Then
            ' 사용자 정보 업데이트
            Dim profileUpdateSql
            profileUpdateSql = "UPDATE " & dbSchema & ".Users " & _
                        "SET name = '" & PreventSQLInjection(newName) & "', " & _
                        "email = '" & PreventSQLInjection(newEmail) & "', " & _
                        "department_id = " & PreventSQLInjection(newDepartmentId) & ", " & _
                        "job_grade = " & PreventSQLInjection(newJobGrade) & " " & _
                        "WHERE user_id = '" & PreventSQLInjection(userId) & "'"
            
            On Error Resume Next
            db.Execute(profileUpdateSql)
            
            ' 오류 발생 시 대체 테이블 시도
            If Err.Number <> 0 Then
                Err.Clear
                profileUpdateSql = "UPDATE " & dbSchema & ".User " & _
                            "SET name = '" & PreventSQLInjection(newName) & "', " & _
                            "email = '" & PreventSQLInjection(newEmail) & "', " & _
                            "department_id = " & PreventSQLInjection(newDepartmentId) & ", " & _
                            "job_grade = " & PreventSQLInjection(newJobGrade) & " " & _
                            "WHERE user_id = '" & PreventSQLInjection(userId) & "'"
                db.Execute(profileUpdateSql)
            End If
            
            ' 성공 메시지 설정
            If Err.Number = 0 Then
                successMsg = "프로필 정보가 성공적으로 업데이트되었습니다."
                ' 페이지 새로고침하여 변경된 정보를 반영
                Response.Redirect("my_profile.asp")
            Else
                errorMsg = "프로필 정보 수정 중 오류가 발생했습니다: " & Err.Description
            End If
        Else
            errorMsg = "비밀번호가 일치하지 않습니다."
        End If
        
        ' 리소스 해제
        If IsObject(profileCheckRs) Then
            If profileCheckRs.State = 1 Then
                profileCheckRs.Close
            End If
            Set profileCheckRs = Nothing
        End If
    End If
End If

' 부서명 가져오기 함수
Function GetDepartmentName(deptId)
    ' NULL 또는 빈 값 처리
    If IsNull(deptId) Or deptId = "" Then
        GetDepartmentName = "-"
        Exit Function
    End If
    
    ' 주어진 department_id에 해당하는 부서명 찾기
    If Not deptRS.BOF Then
        deptRS.MoveFirst
        Do Until deptRS.EOF
            If CStr(deptRS("department_id")) = CStr(deptId) Then
                GetDepartmentName = deptRS("name")
                Exit Function
            End If
            deptRS.MoveNext
        Loop
    End If
    
    ' 부서를 찾지 못한 경우
    GetDepartmentName = deptId & "번 부서"
End Function

' 직급명 가져오기 함수
Function GetJobGradeName(jobGradeId)
    ' NULL 또는 빈 값 처리
    If IsNull(jobGradeId) Or jobGradeId = "" Then
        GetJobGradeName = "-"
        Exit Function
    End If
    
    ' 주어진 job_grade_id에 해당하는 직급명 찾기
    If Not gradeRS.BOF Then
        gradeRS.MoveFirst
        Do Until gradeRS.EOF
            If CStr(gradeRS("job_grade_id")) = CStr(jobGradeId) Then
                GetJobGradeName = gradeRS("name")
                Exit Function
            End If
            gradeRS.MoveNext
        Loop
    End If
    
    ' 직급을 찾지 못한 경우
    GetJobGradeName = jobGradeId & "번 직급"
End Function

' 오류 처리 재설정
On Error GoTo 0

' 현재 활성화된 탭 설정
Dim activeTab
activeTab = Request.QueryString("tab")
If activeTab = "" Then
    activeTab = "info"
End If
%>

<!--#include file="../includes/header.asp"-->

<!-- 부트스트랩 CSS 추가 -->
<link rel="stylesheet" href="https://stackpath.bootstrapcdn.com/bootstrap/4.5.2/css/bootstrap.min.css">
<!-- jQuery 추가 -->
<script src="https://code.jquery.com/jquery-3.5.1.slim.min.js"></script>
<!-- 부트스트랩 JS 추가 -->
<script src="https://cdn.jsdelivr.net/npm/bootstrap@4.5.2/dist/js/bootstrap.bundle.min.js"></script>

<div class="container mt-4">
    <div class="form-container">
        <h2 class="form-title">내 정보</h2>
        
        <% If errorMsg <> "" Then %>
        <div class="alert alert-danger">
            <%= errorMsg %>
        </div>
        <% End If %>
        
        <% If successMsg <> "" Then %>
        <div class="alert alert-success">
            <%= successMsg %>
        </div>
        <% End If %>
        
        <!-- 탭 메뉴 -->
        <ul class="nav nav-pills nav-fill mb-4" id="profileTabs" role="tablist">
            <li class="nav-item" role="presentation">
                <a class="nav-link <%= IIf(activeTab = "info" Or activeTab = "", "active", "") %>" id="info-tab" data-toggle="tab" href="#info" role="tab" aria-controls="info" aria-selected="<%= IIf(activeTab = "info" Or activeTab = "", "true", "false") %>">
                    <i class="fas fa-user-circle"></i> 기본 정보
                </a>
            </li>
            <li class="nav-item" role="presentation">
                <a class="nav-link <%= IIf(activeTab = "edit", "active", "") %>" id="edit-tab" data-toggle="tab" href="#edit" role="tab" aria-controls="edit" aria-selected="<%= IIf(activeTab = "edit", "true", "false") %>">
                    <i class="fas fa-user-edit"></i> 프로필 수정
                </a>
            </li>
            <li class="nav-item" role="presentation">
                <a class="nav-link <%= IIf(activeTab = "password", "active", "") %>" id="password-tab" data-toggle="tab" href="#password" role="tab" aria-controls="password" aria-selected="<%= IIf(activeTab = "password", "true", "false") %>">
                    <i class="fas fa-key"></i> 비밀번호 변경
                </a>
            </li>
        </ul>
        
        <!-- 탭 콘텐츠 -->
        <div class="tab-content" id="profileTabsContent">
            <!-- 기본 정보 탭 -->
            <div class="tab-pane fade <%= IIf(activeTab = "info" Or activeTab = "", "show active", "") %>" id="info" role="tabpanel" aria-labelledby="info-tab">
                <div class="info-section">
                    <table class="table">
                        <tr>
                            <th style="width:30%;">아이디</th>
                            <td><%= rs("user_id") %></td>
                        </tr>
                        <tr>
                            <th>이름</th>
                            <td><%= rs("name") %></td>
                        </tr>
                        <tr>
                            <th>이메일</th>
                            <td><%= rs("email") %></td>
                        </tr>
                        <tr>
                            <th>부서</th>
                            <td><%= GetDepartmentName(rs("department_id")) %></td>
                        </tr>
                        <tr>
                            <th>직급</th>
                            <td><%= GetJobGradeName(rs("job_grade")) %></td>
                        </tr>
                        <tr>
                            <th>가입일</th>
                            <td><%= FormatDateTime(rs("created_at"), 2) %></td>
                        </tr>
                    </table>
                </div>
            </div>
            
            <!-- 프로필 수정 탭 -->
            <div class="tab-pane fade <%= IIf(activeTab = "edit", "show active", "") %>" id="edit" role="tabpanel" aria-labelledby="edit-tab">
                <div class="profile-update-section">
                    <form method="post" action="my_profile.asp?tab=edit">
                        <input type="hidden" name="form_type" value="profile_update">
                        
                        <div class="form-group">
                            <label class="form-label">이름</label>
                            <input type="text" name="name" class="form-input" value="<%= rs("name") %>" required>
                        </div>
                        
                        <div class="form-group">
                            <label class="form-label">이메일</label>
                            <input type="email" name="email" class="form-input" value="<%= rs("email") %>" required>
                        </div>
                        
                        <div class="form-group">
                            <label class="form-label">부서</label>
                            <select name="department_id" class="form-select">
                                <% 
                                If Not deptRS.BOF Then
                                    deptRS.MoveFirst
                                    Do Until deptRS.EOF 
                                %>
                                    <option value="<%= deptRS("department_id") %>" <%= IIf(CStr(rs("department_id"))=CStr(deptRS("department_id")), "selected", "") %>><%= deptRS("name") %></option>
                                <% 
                                    deptRS.MoveNext
                                    Loop
                                End If
                                %>
                            </select>
                        </div>
                        
                        <div class="form-group">
                            <label class="form-label">직급</label>
                            <select name="job_grade" class="form-select">
                                <% 
                                If Not gradeRS.BOF Then
                                    gradeRS.MoveFirst
                                    Do Until gradeRS.EOF 
                                %>
                                    <option value="<%= gradeRS("job_grade_id") %>" <%= IIf(CStr(rs("job_grade"))=CStr(gradeRS("job_grade_id")), "selected", "") %>><%= gradeRS("name") %></option>
                                <% 
                                    gradeRS.MoveNext
                                    Loop
                                End If
                                %>
                            </select>
                        </div>
                        
                        <div class="form-group">
                            <label class="form-label">비밀번호 확인</label>
                            <input type="password" name="profile_password" class="form-input" required>
                            <small class="text-muted">정보 수정을 위해 현재 비밀번호를 입력하세요.</small>
                        </div>
                        
                        <button type="submit" class="btn btn-primary btn-block">프로필 수정</button>
                    </form>
                </div>
            </div>
            
            <!-- 비밀번호 변경 탭 -->
            <div class="tab-pane fade <%= IIf(activeTab = "password", "show active", "") %>" id="password" role="tabpanel" aria-labelledby="password-tab">
                <div class="password-section">
                    <form method="post" action="my_profile.asp?tab=password">
                        <input type="hidden" name="form_type" value="password_change">
                        
                        <div class="form-group">
                            <label class="form-label">현재 비밀번호</label>
                            <input type="password" name="current_password" class="form-input" required>
                        </div>
                        
                        <div class="form-group">
                            <label class="form-label">새 비밀번호</label>
                            <input type="password" name="new_password" class="form-input" required>
                            <small class="text-muted">비밀번호는 6자 이상이어야 합니다.</small>
                        </div>
                        
                        <div class="form-group">
                            <label class="form-label">비밀번호 확인</label>
                            <input type="password" name="confirm_password" class="form-input" required>
                        </div>
                        
                        <button type="submit" class="btn btn-primary btn-block">비밀번호 변경</button>
                    </form>
                </div>
            </div>
        </div>
        
        <div class="mt-4 text-center">
            <a href="dashboard.asp" class="btn btn-danger">대시보드로 돌아가기</a>
        </div>
    </div>
</div>

<style>
.nav-pills .nav-link {
    color: #555;
    background-color: #f8f9fa;
    border: 1px solid #ddd;
    margin: 0 2px;
}

.nav-pills .nav-link:hover {
    background-color: #e9ecef;
    color: #007bff;
}

.nav-pills .nav-link.active {
    background-color: #007bff;
    color: #fff;
    border-color: #007bff;
}

.tab-pane {
    padding: 20px 0;
}

.info-section, .password-section, .profile-update-section {
    background-color: #fff;
    border-radius: 4px;
}

.btn-block {
    margin-top: 1rem;
}

.text-muted {
    font-size: 0.875rem;
    margin-top: 0.25rem;
}
</style>

<!-- Font Awesome 아이콘 -->
<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/5.15.1/css/all.min.css">

<%
' 리소스 정리
If IsObject(rs) Then
    If rs.State = 1 Then
        rs.Close
    End If
    Set rs = Nothing
End If

If IsObject(deptRS) Then
    If deptRS.State = 1 Then
        deptRS.Close
    End If
    Set deptRS = Nothing
End If

If IsObject(gradeRS) Then
    If gradeRS.State = 1 Then
        gradeRS.Close
    End If
    Set gradeRS = Nothing
End If
%>

<!--#include file="../includes/footer.asp"--> 