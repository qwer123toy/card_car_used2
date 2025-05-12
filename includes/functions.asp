<%
' 입력값 검증 함수
Function SanitizeInput(inputValue)
    If IsNull(inputValue) Or inputValue = "" Then
        SanitizeInput = ""
        Exit Function
    End If
    
    ' XSS 방지
    Dim sanitized
    sanitized = Replace(inputValue, "<", "&lt;")
    sanitized = Replace(sanitized, ">", "&gt;")
    sanitized = Replace(sanitized, "'", "''")
    
    SanitizeInput = sanitized
End Function

' SQL 인젝션 방지 함수
Function PreventSQLInjection(str)
    If IsNull(str) Or str = "" Then
        PreventSQLInjection = ""
        Exit Function
    End If
    
    Dim result
    result = Replace(str, "'", "''")
    result = Replace(result, ";", "")
    result = Replace(result, "--", "")
    result = Replace(result, "/*", "")
    result = Replace(result, "*/", "")
    result = Replace(result, "xp_", "")
    
    PreventSQLInjection = result
End Function

' 날짜 형식 변환 (YYYY-MM-DD)
Function FormatDate(dateValue)
    If IsDate(dateValue) Then
        FormatDate = Year(dateValue) & "-" & Right("0" & Month(dateValue), 2) & "-" & Right("0" & Day(dateValue), 2)
    Else
        FormatDate = ""
    End If
End Function

' 숫자 형식 변환 (천 단위 콤마)
Function FormatNumber(numValue)
    If IsNull(numValue) Or numValue = "" Then
        FormatNumber = "0"
    Else
        FormatNumber = FormatCurrency(numValue, 0)
    End If
End Function

' 사용자 인증 확인
Function IsAuthenticated()
    If Session("user_id") = "" Then
        IsAuthenticated = False
    Else
        IsAuthenticated = True
    End If
End Function

' 관리자 권한 확인
Function IsAdmin()
    If Session("is_admin") = "Y" Then
        IsAdmin = True
    Else
        IsAdmin = False
    End If
End Function

' 페이지 리디렉션
Sub RedirectTo(url)
    Response.Clear
    Response.Redirect(url)
    Response.End
End Sub

' 로그 기록
Sub LogActivity(user_id, action, description)
    Dim SQL
    SQL = "INSERT INTO ActivityLogs (user_id, action, description, created_at) VALUES ('" & PreventSQLInjection(user_id) & "', '" & PreventSQLInjection(action) & "', '" & PreventSQLInjection(description) & "', GETDATE())"
    
    On Error Resume Next
    dbConn.Execute SQL
    On Error GoTo 0
End Sub
%> 