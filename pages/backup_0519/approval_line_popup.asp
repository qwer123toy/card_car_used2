<%@ Language="VBScript" CodePage="65001" %>
<% 
Response.CodePage = 65001
Response.Charset = "utf-8"
%>

<!--#include file="../db.asp"-->
<!--#include file="../includes/functions.asp"-->

<%
Dim sql, rs
sql = "SELECT U.user_id, U.name, U.job_grade, D.name AS dept_name " & _
      "FROM Users U LEFT JOIN Department D ON U.department_id = D.department_id " & _
      "ORDER BY D.name, U.job_grade"
Set rs = db.Execute(sql)
%>

<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8">
    <title>결재라인 지정</title>
    <style>
        table {
            width: 100%;
            border-collapse: collapse;
        }
        th, td {
            border: 1px solid #ccc;
            padding: 8px;
            text-align: center;
        }
        th {
            background-color: #f0f0f0;
        }
    </style>
</head>
<body>
<h2>결재라인 지정</h2>

<table>
    <tr>
        <th>부서</th>
        <th>직급</th>
        <th>이름</th>
        <th>결재라인 추가</th>
    </tr>
<% Do Until rs.EOF %>
    <tr>
        <td><%= rs("dept_name") %></td>
        <td><%= rs("job_grade") %></td>
        <td><%= rs("name") %></td>
        <td>
            <button onclick="addApprover('<%= rs("user_id") %>', '<%= rs("name") %>')">추가</button>
        </td>
    </tr>
<% 
    rs.MoveNext
Loop
rs.Close
%>
</table>

<script>
function addApprover(userId, userName) {
    const maxSteps = 3;
    for (let i = 1; i <= maxSteps; i++) {
        const nameCell = window.opener.document.getElementById("step" + i + "_name");
        const hiddenInput = window.opener.document.getElementById("step" + i + "_id");

        if (nameCell && nameCell.innerText === "-") {
            nameCell.innerText = userName;
            hiddenInput.value = userId;
            window.close();
            return;
        }
    }
    alert("결재자는 최대 3명까지 지정할 수 있습니다.");
}
</script>
</body>
</html>