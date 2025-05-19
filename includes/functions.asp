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
    ' SQL 인젝션 방지 처리
    Dim safeUserId, safeAction, safeDesc
    safeUserId = PreventSQLInjection(user_id)
    safeAction = PreventSQLInjection(action)
    safeDesc = PreventSQLInjection(description)
    
    ' 파라미터화된 쿼리 사용
    Dim cmd
    Set cmd = Server.CreateObject("ADODB.Command")
    
    On Error Resume Next
    cmd.ActiveConnection = db
    cmd.CommandText = "INSERT INTO ActivityLogs (user_id, action, description, created_at) VALUES (?, ?, ?, GETDATE())"
    
    ' 파라미터 추가
    cmd.Parameters.Append cmd.CreateParameter("@user_id", 200, 1, 30, safeUserId)
    cmd.Parameters.Append cmd.CreateParameter("@action", 200, 1, 50, safeAction)
    cmd.Parameters.Append cmd.CreateParameter("@description", 200, 1, 200, safeDesc)
    
    ' 명령 실행
    cmd.Execute
    On Error GoTo 0
End Sub

' 조건부 표현식 함수 (IIf 대체)
Function If_Cond(condition, trueValue, falseValue)
    If condition Then
        If_Cond = trueValue
    Else
        If_Cond = falseValue
    End If
End Function

' IIf 함수 - 조건부 표현식을 간결하게 처리하는 함수
Function IIf(condition, trueValue, falseValue)
    If condition Then
        IIf = trueValue
    Else
        IIf = falseValue
    End If
End Function

' 부서명 가져오기 함수
Function GetDepartmentName(departmentId)
    ' Null 검사 및 빈 문자열 검사
    If IsNull(departmentId) Or departmentId = "" Then
        GetDepartmentName = "-"
        Exit Function
    End If
    
    ' 부서 코드표 - Department 테이블이 없는 경우를 대비하여 하드코딩합니다
    Dim deptName
    
    Select Case CStr(departmentId)
        Case "1": deptName = "인사팀"
        Case "2": deptName = "재무팀"
        Case "3": deptName = "영업팀"
        Case "4": deptName = "마케팅팀"
        Case "5": deptName = "개발팀"
        Case "6": deptName = "디자인팀"
        Case "7": deptName = "경영지원팀"
        Case Else: deptName = departmentId & " 부서"
    End Select
    
    GetDepartmentName = deptName
End Function

' 페이징 처리 함수 (OFFSET 사용 불가 시 대체 방법)
Function GetPagedRecords(tableName, orderByColumn, pageNumber, pageSize, whereCondition)
    Dim topCount, SQL, totalSQL
    
    If whereCondition = "" Then
        whereCondition = "1=1"
    End If
    
    ' 전체 레코드 수 조회
    totalSQL = "SELECT COUNT(*) AS total FROM " & tableName & " WHERE " & whereCondition
    
    ' 페이징 처리 (OFFSET 사용하지 않는 대안)
    topCount = pageSize
    If pageNumber > 1 Then
        ' 이전 페이지들의 레코드를 건너뛰기 위한 SQL
        SQL = "SELECT TOP " & topCount & " * FROM " & tableName & " " & _
              "WHERE " & whereCondition & " AND " & orderByColumn & " NOT IN " & _
              "(SELECT TOP " & ((pageNumber - 1) * pageSize) & " " & orderByColumn & " " & _
              "FROM " & tableName & " WHERE " & whereCondition & " " & _
              "ORDER BY " & orderByColumn & " DESC) " & _
              "ORDER BY " & orderByColumn & " DESC"
    Else
        ' 첫 페이지 조회
        SQL = "SELECT TOP " & topCount & " * FROM " & tableName & " " & _
              "WHERE " & whereCondition & " " & _
              "ORDER BY " & orderByColumn & " DESC"
    End If
    
    GetPagedRecords = SQL
End Function
%> 