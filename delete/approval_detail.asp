<%@ Language="VBScript" CodePage="65001" %>
<%
Response.CodePage = 65001
Response.CharSet = "utf-8"

' 디버그 모드 설정
Dim isDebugMode
isDebugMode = False

' 오류 처리 초기화
On Error Resume Next
%>

<!--#include file="../db.asp"-->
<!--#include file="../includes/functions.asp"-->

<%
' 로그인 상태 확인
If Not IsAuthenticated() Then
    Response.Write "<script>alert('로그인이 필요합니다.'); location.href='../index.asp';</script>"
    Response.End
End If

' 파라미터 검증
Dim usageId
usageId = Request.QueryString("id")
If usageId = "" Then
    Response.Write "<script>alert('잘못된 접근입니다.'); history.back();</script>"
    Response.End
End If

' 현재 로그인한 사용자 정보
Dim currentUserId
currentUserId = Session("user_id")

' CardUsage 정보 조회
Dim usageSQL, usageRS
usageSQL = "SELECT cu.*, u.name as requester_name, u.department_id as requester_dept_id " & _
           "FROM " & dbSchema & ".CardUsage cu " & _
           "LEFT JOIN " & dbSchema & ".Users u ON cu.user_id = u.user_id " & _
           "WHERE cu.usage_id = " & usageId

Set usageRS = db.Execute(usageSQL)

If usageRS.EOF Then
    Response.Write "<script>alert('존재하지 않는 결재 문서입니다.'); history.back();</script>"
    Response.End
End If

' 결재 로그 조회
Dim logsSQL, logsRS
logsSQL = "SELECT al.*, u.name as approver_name, u.department_id, u.job_grade " & _
          "FROM " & dbSchema & ".ApprovalLogs al " & _
          "LEFT JOIN " & dbSchema & ".Users u ON al.approver_id = u.user_id " & _
          "WHERE al.target_table_name = 'CardUsage' " & _
          "AND al.target_id = " & usageId & " " & _
          "ORDER BY al.approval_step"

Set logsRS = db.Execute(logsSQL)

' 현재 사용자의 결재 단계 확인
Dim currentUserStep, canApprove
currentUserStep = 0
canApprove = False

If Not logsRS.EOF Then
    logsRS.MoveFirst
    Do Until logsRS.EOF
        If logsRS("approver_id") = currentUserId Then
            currentUserStep = logsRS("approval_step")
            ' 이전 단계가 모두 승인되었고, 현재 단계가 대기 상태인 경우에만 결재 가능
            If currentUserStep = 1 Or _
               (currentUserStep > 1 And IsPreviousStepApproved(usageId, currentUserStep - 1)) Then
                If logsRS("status") = "대기" Then
                    canApprove = True
                End If
            End If
            Exit Do
        End If
        logsRS.MoveNext
    Loop
    logsRS.MoveFirst
End If

' 이전 단계 승인 여부 확인 함수
Function IsPreviousStepApproved(targetId, step)
    Dim prevSQL, prevRS
    prevSQL = "SELECT status FROM " & dbSchema & ".ApprovalLogs " & _
              "WHERE target_table_name = 'CardUsage' " & _
              "AND target_id = " & targetId & " " & _
              "AND approval_step = " & step
    
    Set prevRS = db.Execute(prevSQL)
    
    If Not prevRS.EOF Then
        IsPreviousStepApproved = (prevRS("status") = "승인")
    Else
        IsPreviousStepApproved = False
    End If
    
    Set prevRS = Nothing
End Function

' 부서명 조회 함수
Function GetDepartmentName(deptId)
    Dim deptSQL, deptRS
    deptSQL = "SELECT name FROM " & dbSchema & ".Department WHERE department_id = " & deptId
    Set deptRS = db.Execute(deptSQL)
    
    If Not deptRS.EOF Then
        GetDepartmentName = deptRS("name")
    Else
        GetDepartmentName = "알 수 없음"
    End If
    
    Set deptRS = Nothing
End Function

' 직급명 조회 함수
Function GetJobGradeName(gradeId)
    Dim gradeSQL, gradeRS
    gradeSQL = "SELECT name FROM " & dbSchema & ".JobGrade WHERE job_grade_id = " & gradeId
    Set gradeRS = db.Execute(gradeSQL)
    
    If Not gradeRS.EOF Then
        GetJobGradeName = gradeRS("name")
    Else
        GetJobGradeName = "알 수 없음"
    End If
    
    Set gradeRS = Nothing
End Function
%>

<!--#include file="../includes/header.asp"-->

<div class="container mt-4">
    <h2 class="mb-4">결재 상세 정보</h2>
    
    <!-- 카드 사용 정보 -->
    <div class="card mb-4">
        <div class="card-header">
            <h5 class="card-title mb-0">카드 사용 내역</h5>
        </div>
        <div class="card-body">
            <table class="table table-bordered">
                <tr>
                    <th style="width: 150px;">신청자</th>
                    <td>
                        <%= usageRS("requester_name") %>
                        (<%= GetDepartmentName(usageRS("requester_dept_id")) %>)
                    </td>
                    <th style="width: 150px;">신청일</th>
                    <td><%= FormatDateTime(usageRS("created_at"), 2) %></td>
                </tr>
                <tr>
                    <th>사용 카드</th>
                    <td><%= usageRS("card_name") %></td>
                    <th>사용일자</th>
                    <td><%= FormatDateTime(usageRS("usage_date"), 2) %></td>
                </tr>
                <tr>
                    <th>사용처</th>
                    <td><%= usageRS("store_name") %></td>
                    <th>금액</th>
                    <td><%= FormatNumber(usageRS("amount"), 0) %>원</td>
                </tr>
                <tr>
                    <th>사용목적</th>
                    <td colspan="3"><%= usageRS("purpose") %></td>
                </tr>
                <tr>
                    <th>현재 상태</th>
                    <td colspan="3">
                        <span class="badge badge-<%= GetStatusClass(usageRS("approval_status")) %>">
                            <%= usageRS("approval_status") %>
                        </span>
                    </td>
                </tr>
            </table>
        </div>
    </div>
    
    <!-- 결재선 정보 -->
    <div class="card mb-4">
        <div class="card-header">
            <h5 class="card-title mb-0">결재선</h5>
        </div>
        <div class="card-body">
            <table class="table table-bordered">
                <thead>
                    <tr>
                        <th style="width: 100px;">단계</th>
                        <th>결재자</th>
                        <th style="width: 150px;">상태</th>
                        <th style="width: 200px;">처리일시</th>
                    </tr>
                </thead>
                <tbody>
                    <% 
                    If Not logsRS.EOF Then
                        Do Until logsRS.EOF 
                    %>
                    <tr>
                        <td><%= logsRS("approval_step") %>차 결재</td>
                        <td>
                            <%= logsRS("approver_name") %>
                            (<%= GetDepartmentName(logsRS("department_id")) %> / 
                             <%= GetJobGradeName(logsRS("job_grade")) %>)
                        </td>
                        <td>
                            <span class="badge badge-<%= GetStatusClass(logsRS("status")) %>">
                                <%= logsRS("status") %>
                            </span>
                        </td>
                        <td>
                            <% If logsRS("processed_at") <> "" Then %>
                                <%= FormatDateTime(logsRS("processed_at"), 2) %>
                            <% End If %>
                        </td>
                    </tr>
                    <%
                        logsRS.MoveNext
                        Loop
                    End If
                    %>
                </tbody>
            </table>
        </div>
    </div>
    
    <!-- 결재 처리 버튼 -->
    <% If canApprove Then %>
    <div class="card mb-4">
        <div class="card-header">
            <h5 class="card-title mb-0">결재 처리</h5>
        </div>
        <div class="card-body">
            <form method="post" action="approval_update.asp" onsubmit="return validateForm();">
                <input type="hidden" name="usage_id" value="<%= usageId %>">
                <input type="hidden" name="approval_step" value="<%= currentUserStep %>">
                
                <div class="form-group">
                    <label for="comment">결재 의견</label>
                    <textarea class="form-control" id="comment" name="comment" rows="3"></textarea>
                </div>
                
                <div class="btn-group">
                    <button type="submit" name="action" value="approve" class="btn btn-success mr-2">
                        <i class="fas fa-check"></i> 승인
                    </button>
                    <button type="submit" name="action" value="reject" class="btn btn-danger">
                        <i class="fas fa-times"></i> 반려
                    </button>
                </div>
            </form>
        </div>
    </div>
    <% End If %>
    
    <div class="text-center mt-4">
        <a href="dashboard.asp" class="btn btn-secondary">
            <i class="fas fa-arrow-left"></i> 목록으로 돌아가기
        </a>
    </div>
</div>

<%
' 상태에 따른 배지 클래스 반환
Function GetStatusClass(status)
    Select Case status
        Case "대기"
            GetStatusClass = "warning"
        Case "승인"
            GetStatusClass = "success"
        Case "반려"
            GetStatusClass = "danger"
        Case Else
            GetStatusClass = "secondary"
    End Select
End Function
%>

<style>
.badge {
    font-size: 0.9rem;
    padding: 0.5em 1em;
}
.badge-warning {
    background-color: #ffc107;
    color: #000;
}
.badge-success {
    background-color: #28a745;
    color: #fff;
}
.badge-danger {
    background-color: #dc3545;
    color: #fff;
}
.badge-secondary {
    background-color: #6c757d;
    color: #fff;
}
</style>

<script>
function validateForm() {
    var comment = document.getElementById('comment').value.trim();
    var action = event.submitter.value;
    
    if (action === 'reject' && comment === '') {
        alert('반려 시에는 반드시 의견을 입력해주세요.');
        return false;
    }
    
    return true;
}
</script>

<!--#include file="../includes/footer.asp"--> 