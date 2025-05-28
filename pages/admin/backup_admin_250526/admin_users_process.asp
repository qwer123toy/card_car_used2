<%@ Language="VBScript" CodePage="65001" %>
<% 
Response.CodePage = 65001
Response.CharSet = "utf-8"
%>

<!--#include file="../../db.asp"-->
<!--#include file="../../includes/functions.asp"-->
<%
' 로그인 체크
If Not IsAuthenticated() Then
    RedirectTo("../../index.asp")
End If

' 관리자 권한 체크
If Not IsAdmin() Then
    Response.Write("<script>alert('관리자 권한이 필요합니다.'); window.location.href='../dashboard.asp';</script>")
    Response.End
End If

' POST 요청인지 확인
If Request.ServerVariables("REQUEST_METHOD") <> "POST" Then
    Response.Write("<script>alert('잘못된 접근입니다.'); window.location.href='admin_users.asp';</script>")
    Response.End
End If

' 폼 데이터 가져오기
Dim action, userId, userName, userEmail, userPhone, userDeptId, userGradeId, isAdmin, isActive
Dim userPassword, userConfirmPassword

action = Request.Form("action")
userName = PreventSQLInjection(Request.Form("name"))
userEmail = PreventSQLInjection(Request.Form("email"))
userPhone = PreventSQLInjection(Request.Form("phone"))

' 부서, 직급 ID 처리
If Request.Form("department_id") <> "" Then
    userDeptId = CInt(Request.Form("department_id"))
Else
    userDeptId = Null
End If

If Request.Form("job_grade") <> "" Then
    userGradeId = CInt(Request.Form("job_grade"))
Else
    userGradeId = Null
End If

' 관리자 및 활성화 상태 확인
If Request.Form("is_admin") = "1" Then
    isAdmin = True
Else
    isAdmin = False
End If

If Request.Form("is_active") = "1" Then
    isActive = True
Else
    isActive = False
End If

' 유효성 검사
If userName = "" Then
    Response.Write("<script>alert('이름을 입력해주세요.'); history.back();</script>")
    Response.End
End If

On Error Resume Next

If action = "add" Then
    ' 사용자 추가
    userId = PreventSQLInjection(Request.Form("user_id"))
    userPassword = Request.Form("password")
    userConfirmPassword = Request.Form("confirm_password")
    
    ' 추가 유효성 검사
    If userId = "" Then
        Response.Write("<script>alert('사용자 ID를 입력해주세요.'); history.back();</script>")
        Response.End
    End If
    
    If userPassword = "" Then
        Response.Write("<script>alert('비밀번호를 입력해주세요.'); history.back();</script>")
        Response.End
    End If
    
    If userPassword <> userConfirmPassword Then
        Response.Write("<script>alert('비밀번호와 비밀번호 확인이 일치하지 않습니다.'); history.back();</script>")
        Response.End
    End If
    
    ' 중복 ID 확인
    Dim checkSQL, checkRS
    checkSQL = "SELECT COUNT(*) AS cnt FROM " & dbSchema & ".Users WHERE user_id = '" & userId & "'"
    Set checkRS = db.Execute(checkSQL)
    
    If Not checkRS.EOF And checkRS("cnt") > 0 Then
        Response.Write("<script>alert('이미 사용 중인 사용자 ID입니다. 다른 ID를 입력해주세요.'); history.back();</script>")
        Response.End
    End If
    
    ' 비밀번호 해시
    Dim hashedPassword
    hashedPassword = HashPassword(userPassword)
    
    ' 사용자 추가
    Dim addSQL
    addSQL = "INSERT INTO " & dbSchema & ".Users " & _
             "(user_id, password, name, email, phone, department_id, job_grade, is_admin, is_active, created_at) " & _
             "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, GETDATE())"
    
    Dim cmdAdd
    Set cmdAdd = Server.CreateObject("ADODB.Command")
    cmdAdd.ActiveConnection = db
    cmdAdd.CommandText = addSQL
    cmdAdd.Parameters.Append cmdAdd.CreateParameter("@user_id", 200, 1, 50, userId)
    cmdAdd.Parameters.Append cmdAdd.CreateParameter("@password", 200, 1, 100, hashedPassword)
    cmdAdd.Parameters.Append cmdAdd.CreateParameter("@name", 200, 1, 50, userName)
    cmdAdd.Parameters.Append cmdAdd.CreateParameter("@email", 200, 1, 100, IIf(userEmail = "", Null, userEmail))
    cmdAdd.Parameters.Append cmdAdd.CreateParameter("@phone", 200, 1, 20, IIf(userPhone = "", Null, userPhone))
    cmdAdd.Parameters.Append cmdAdd.CreateParameter("@department_id", 3, 1, , IIf(IsNull(userDeptId), Null, userDeptId))
    cmdAdd.Parameters.Append cmdAdd.CreateParameter("@job_grade", 3, 1, , IIf(IsNull(userGradeId), Null, userGradeId))
    cmdAdd.Parameters.Append cmdAdd.CreateParameter("@is_admin", 11, 1, , isAdmin)
    cmdAdd.Parameters.Append cmdAdd.CreateParameter("@is_active", 11, 1, , isActive)
    
    cmdAdd.Execute
    
    If Err.Number <> 0 Then
        Response.Write("<script>alert('사용자 추가 중 오류가 발생했습니다: " & Server.HTMLEncode(Err.Description) & "'); history.back();</script>")
        Response.End
    Else
        ' 활동 로그 기록
        LogActivity Session("user_id"), "사용자추가", "사용자 추가 (ID: " & userId & ", 이름: " & userName & ")"
        Response.Write("<script>alert('사용자가 추가되었습니다.'); window.location.href='admin_users.asp';</script>")
        Response.End
    End If
    
ElseIf action = "edit" Then
    ' 사용자 수정
    userId = PreventSQLInjection(Request.Form("user_id"))
    userPassword = Request.Form("password")
    userConfirmPassword = Request.Form("confirm_password")
    
    If userId = "" Then
        Response.Write("<script>alert('사용자 ID가 필요합니다.'); window.location.href='admin_users.asp';</script>")
        Response.End
    End If
    
    ' 비밀번호 변경 여부 확인
    Dim passwordSQL, hashedPassword
    If userPassword <> "" Then
        If userPassword <> userConfirmPassword Then
            Response.Write("<script>alert('비밀번호와 비밀번호 확인이 일치하지 않습니다.'); history.back();</script>")
            Response.End
        End If
        
        hashedPassword = HashPassword(userPassword)
        passwordSQL = "password = '" & hashedPassword & "', "
    Else
        passwordSQL = ""
    End If
    
    ' 현재 로그인한 사용자가 자신의 관리자 권한을 제거하는 경우 방지
    If userId = Session("user_id") And Not isAdmin Then
        Response.Write("<script>alert('자신의 관리자 권한을 제거할 수 없습니다.'); history.back();</script>")
        Response.End
    End If
    
    ' 사용자 수정
    Dim editSQL
    editSQL = "UPDATE " & dbSchema & ".Users SET " & _
              passwordSQL & _
              "name = ?, email = ?, phone = ?, " & _
              "department_id = ?, job_grade = ?, is_admin = ?, is_active = ? " & _
              "WHERE user_id = ?"
    
    Dim cmdEdit
    Set cmdEdit = Server.CreateObject("ADODB.Command")
    cmdEdit.ActiveConnection = db
    cmdEdit.CommandText = editSQL
    cmdEdit.Parameters.Append cmdEdit.CreateParameter("@name", 200, 1, 50, userName)
    cmdEdit.Parameters.Append cmdEdit.CreateParameter("@email", 200, 1, 100, IIf(userEmail = "", Null, userEmail))
    cmdEdit.Parameters.Append cmdEdit.CreateParameter("@phone", 200, 1, 20, IIf(userPhone = "", Null, userPhone))
    cmdEdit.Parameters.Append cmdEdit.CreateParameter("@department_id", 3, 1, , IIf(IsNull(userDeptId), Null, userDeptId))
    cmdEdit.Parameters.Append cmdEdit.CreateParameter("@job_grade", 3, 1, , IIf(IsNull(userGradeId), Null, userGradeId))
    cmdEdit.Parameters.Append cmdEdit.CreateParameter("@is_admin", 11, 1, , isAdmin)
    cmdEdit.Parameters.Append cmdEdit.CreateParameter("@is_active", 11, 1, , isActive)
    cmdEdit.Parameters.Append cmdEdit.CreateParameter("@user_id", 200, 1, 50, userId)
    
    cmdEdit.Execute
    
    If Err.Number <> 0 Then
        Response.Write("<script>alert('사용자 수정 중 오류가 발생했습니다: " & Server.HTMLEncode(Err.Description) & "'); history.back();</script>")
        Response.End
    Else
        ' 활동 로그 기록
        LogActivity Session("user_id"), "사용자수정", "사용자 수정 (ID: " & userId & ", 이름: " & userName & ")"
        Response.Write("<script>alert('사용자 정보가 수정되었습니다.'); window.location.href='admin_users.asp';</script>")
        Response.End
    End If
    
Else
    Response.Write("<script>alert('잘못된 요청입니다.'); window.location.href='admin_users.asp';</script>")
End If

On Error GoTo 0
%> 