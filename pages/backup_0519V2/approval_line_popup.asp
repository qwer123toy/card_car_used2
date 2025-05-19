<%@ Language="VBScript" CodePage="65001" %>
<% 
Response.CodePage = 65001
Response.CharSet = "utf-8"
%>

<!--#include file="../db.asp"-->
<!--#include file="../includes/functions.asp"-->
<%
' 로그인 체크
If Not IsAuthenticated() Then
    Response.Write "로그인이 필요합니다."
    Response.End
End If

' 사용자 목록 조회 (직급 정보 포함)
Dim userSQL, userRS
userSQL = "SELECT u.user_id, u.name, j.name AS job_grade_name " & _
          "FROM " & dbSchema & ".Users u " & _
          "LEFT JOIN " & dbSchema & ".job_grade j ON u.job_grade = j.job_grade_id " & _
          "ORDER BY j.job_grade_id DESC, u.name ASC"
Set userRS = db.Execute(userSQL)
%>

<!DOCTYPE html>
<html>
<head>
    <title>결재자 선택</title>
    <meta charset="utf-8">
    <style>
        body { font-family: 'Malgun Gothic', sans-serif; margin: 20px; }
        .user-list {
            border: 1px solid #ddd;
            padding: 10px;
            height: 400px;
            overflow-y: auto;
        }
        .user-item {
            padding: 8px;
            cursor: pointer;
            border-bottom: 1px solid #eee;
            display: flex;
            justify-content: space-between;
        }
        .user-item:hover {
            background-color: #f5f5f5;
        }
        .job-grade {
            color: #666;
            font-size: 0.9em;
        }
        .selected {
            background-color: #e3f2fd;
        }
        .buttons {
            margin-top: 20px;
            text-align: center;
        }
        .btn {
            padding: 8px 20px;
            margin: 0 5px;
            border: none;
            border-radius: 4px;
            cursor: pointer;
        }
        .btn-primary {
            background-color: #007bff;
            color: white;
        }
        .btn-secondary {
            background-color: #6c757d;
            color: white;
        }
    </style>
</head>
<body>
    <h2>결재자 선택</h2>
    <div class="user-list">
        <% 
        If Not userRS.EOF Then
            Do While Not userRS.EOF 
                Dim jobGrade
                If IsNull(userRS("job_grade_name")) Then
                    jobGrade = "직급없음"
                Else
                    jobGrade = userRS("job_grade_name")
                End If
        %>
            <div class="user-item" onclick="selectUser('<%= userRS("user_id") %>', '<%= userRS("name") %>', '<%= jobGrade %>')">
                <span class="name"><%= userRS("name") %></span>
                <span class="job-grade"><%= jobGrade %></span>
            </div>
        <% 
                userRS.MoveNext
            Loop
        End If
        %>
    </div>
    
    <div class="buttons">
        <button class="btn btn-primary" onclick="applySelection()">선택</button>
        <button class="btn btn-secondary" onclick="window.close()">취소</button>
    </div>

    <script>
        var selectedUserId = '';
        var selectedUserName = '';
        var selectedJobGrade = '';
        
        function selectUser(userId, userName, jobGrade) {
            // 이전 선택 해제
            var items = document.getElementsByClassName('user-item');
            for (var i = 0; i < items.length; i++) {
                items[i].classList.remove('selected');
            }
            
            // 현재 선택한 항목 표시
            event.currentTarget.classList.add('selected');
            
            selectedUserId = userId;
            selectedUserName = userName;
            selectedJobGrade = jobGrade;
        }
        
        function applySelection() {
            if (!selectedUserId) {
                alert('결재자를 선택해주세요.');
                return;
            }
            
            // 부모 창의 빈 결재자 칸을 찾아 설정
            var parentWindow = window.opener;
            var step = 1;
            
            // 이미 설정된 결재자와 중복 체크
            for (var i = 1; i <= 3; i++) {
                var existingId = parentWindow.document.getElementById('approver_step' + i).value;
                if (existingId === selectedUserId) {
                    alert('이미 선택된 결재자입니다.');
                    return;
                }
                if (!existingId) {
                    step = i;
                    break;
                }
            }
            
            // 결재자 정보 설정
            parentWindow.setApprover(step, selectedUserId, selectedUserName, selectedJobGrade);
            window.close();
        }
    </script>
</body>
</html>