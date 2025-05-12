<%
' 실제 프로젝트에서 사용하는 DB 연결 파일은 .gitignore에 포함되어 있으며,
' 이 파일은 GitHub에 올라가는 버전으로 비밀번호 등 민감 정보는 생략되어 있습니다.

Dim strConnection, dbConn

' 실제 프로젝트에서는 이 부분에 적절한 연결 문자열을 입력하세요
strConnection = "Provider=SQLOLEDB;Data Source=YOUR_SERVER;Initial Catalog=YOUR_DATABASE;User ID=YOUR_USERNAME;Password=YOUR_PASSWORD;"

Set dbConn = Server.CreateObject("ADODB.Connection")
dbConn.ConnectionTimeout = 20000
dbConn.Open strConnection
dbConn.cursorlocation = 3

' 오류 처리
If Err.Number <> 0 Then
    Response.Write "데이터베이스 연결 오류: " & Err.Description
    Response.End
End If
%> 