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
    RedirectTo("../index.asp")
End If

On Error Resume Next

' 현재 사용자의 직급 정보 가져오기
Dim userJobGradeName
userJobGradeName = ""
If Session("user_id") <> "" Then
    Dim userSQL, userRS
    userSQL = "SELECT j.name as job_grade_name FROM " & dbSchema & ".Users u " & _
              "LEFT JOIN " & dbSchema & ".Job_Grade j ON u.job_grade = j.job_grade_id " & _
              "WHERE u.user_id = '" & Session("user_id") & "'"
    Set userRS = db99.Execute(userSQL)
    If Not userRS.EOF And Not IsNull(userRS("job_grade_name")) Then
        userJobGradeName = userRS("job_grade_name")
    End If
    If Not userRS Is Nothing Then
        If userRS.State = 1 Then userRS.Close
        Set userRS = Nothing
    End If
End If

' 카드 계정 목록 조회
Dim cardSQL, cardRS
cardSQL = "SELECT card_id, account_name, issuer FROM " & dbSchema & ".CardAccount ORDER BY account_name"
Set cardRS = db99.Execute(cardSQL)

' 계정 과목 목록 조회
Dim accountTypeSQL, accountTypeRS
accountTypeSQL = "SELECT account_type_id, type_name FROM " & dbSchema & ".CardAccountTypes ORDER BY type_name"
Set accountTypeRS = db99.Execute(accountTypeSQL)


' 에러 메시지 및 성공 메시지 초기화
Dim errorMsg, successMsg

' 카드 사용 등록 처리
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    Dim cardId, usageDate, amount, purpose, accountTypeId, storeName, title
    Dim approver1, approver2, approver3
    
    cardId = PreventSQLInjection(Request.Form("card_id"))
    usageDate = PreventSQLInjection(Request.Form("usage_date"))
    amount = PreventSQLInjection(Request.Form("amount"))
    purpose = PreventSQLInjection(Request.Form("purpose"))
    accountTypeId = PreventSQLInjection(Request.Form("account_type_id"))
    storeName = PreventSQLInjection(Request.Form("store_name"))
    title = PreventSQLInjection(Request.Form("title"))
    approver1 = PreventSQLInjection(Request.Form("approver_step1"))
    approver2 = PreventSQLInjection(Request.Form("approver_step2"))
    approver3 = PreventSQLInjection(Request.Form("approver_step3"))
    Dim approver4, approver5
    approver4 = PreventSQLInjection(Request.Form("approver_step4"))
    approver5 = PreventSQLInjection(Request.Form("approver_step5"))
    
    ' 입력값 검증
    If cardId = "" Or usageDate = "" Or amount = "" Or accountTypeId = "" Or storeName = "" Or title = "" Then
        errorMsg = "필수 항목을 모두 입력해주세요."
    ElseIf Not IsNumeric(amount) Then
        errorMsg = "금액은 숫자만 입력 가능합니다."
    Else
        ' SQL 쿼리 수정 - department_id 필드 추가
        Dim cmd, departmentId
        Set cmd = Server.CreateObject("ADODB.Command")
        cmd.ActiveConnection = db
        
        ' 사용자의 부서 ID 가져오기 (세션에서)
        If Session("department_id") <> "" Then
            departmentId = Session("department_id")
        Else
            departmentId = 1 ' 기본값 설정 (관리부)
        End If
        
        ' 트랜잭션 시작
        db.BeginTrans
        
        On Error Resume Next
        
        ' SQL 명령문 - 실제 DB 구조에 맞게 수정
        cmd.CommandText = "INSERT INTO " & dbSchema & ".CardUsage (user_id, card_id, usage_date, amount, expense_category_id, department_id, store_name, purpose, title, approval_status) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)"
        
        ' 파라미터 추가
        cmd.Parameters.Append cmd.CreateParameter("@user_id", 200, 1, 30, Session("user_id"))
        cmd.Parameters.Append cmd.CreateParameter("@card_id", 3, 1, , cardId)
        cmd.Parameters.Append cmd.CreateParameter("@usage_date", 7, 1, , usageDate)
        cmd.Parameters.Append cmd.CreateParameter("@amount", 6, 1, , CDbl(amount))
        cmd.Parameters.Append cmd.CreateParameter("@expense_category_id", 3, 1, , accountTypeId)
        cmd.Parameters.Append cmd.CreateParameter("@department_id", 3, 1, , departmentId)
        cmd.Parameters.Append cmd.CreateParameter("@store_name", 200, 1, 100, storeName)
        cmd.Parameters.Append cmd.CreateParameter("@purpose", 200, 1, 200, purpose)
        cmd.Parameters.Append cmd.CreateParameter("@title", 200, 1, 200, title)
        cmd.Parameters.Append cmd.CreateParameter("@approval_status", 200, 1, 20, "대기")
        
        cmd.Execute
        
        If Err.Number = 0 Then
            ' 방금 삽입된 CardUsage의 ID 가져오기
            Dim usageId
            Set cmd = Server.CreateObject("ADODB.Command")
            cmd.ActiveConnection = db
            cmd.CommandText = "SELECT IDENT_CURRENT('CardUsage') AS usage_id"
            Dim rs
            Set rs = cmd.Execute()
            usageId = rs("usage_id")
            
            ' ApprovalLogs 테이블에 결재선 데이터 삽입
            Dim approverIds(4)
            approverIds(0) = approver1
            approverIds(1) = approver2
            approverIds(2) = approver3
            approverIds(3) = approver4
            approverIds(4) = approver5
            
            Dim i
            For i = 0 To 4
                If approverIds(i) <> "" Then
                    Dim approvalSQL
                    approvalSQL = "INSERT INTO " & dbSchema & ".ApprovalLogs " & _
                                "(approver_id, target_table_name, target_id, approval_step, status, created_at) " & _
                                "VALUES (?, 'CardUsage', ?, ?, '대기', GETDATE())"
                    
                    Set cmd = Server.CreateObject("ADODB.Command")
                    cmd.ActiveConnection = db
                    cmd.CommandText = approvalSQL
                    cmd.Parameters.Append cmd.CreateParameter("@approver_id", 200, 1, 30, approverIds(i))
                    cmd.Parameters.Append cmd.CreateParameter("@target_id", 3, 1, , usageId)
                    cmd.Parameters.Append cmd.CreateParameter("@approval_step", 3, 1, , i + 1)
                    
                    cmd.Execute
                    
                    If Err.Number <> 0 Then
                        Exit For
                    End If
                End If
            Next
            
            If Err.Number = 0 Then
                db.CommitTrans
                successMsg = "카드 사용 내역이 등록되었습니다."
                Response.Redirect "card_usage.asp"
            Else
                db.RollbackTrans
                errorMsg = "결재선 등록 중 오류가 발생했습니다: " & Err.Description
            End If
        Else
            db.RollbackTrans
            errorMsg = "카드 사용 내역 등록 중 오류가 발생했습니다: " & Err.Description
        End If
        
        On Error GoTo 0
    End If
End If

On Error GoTo 0
%>

<!--#include file="../includes/header.asp"-->

<style>
.container { 
    max-width: 900px; 
    padding: 2rem 1rem;
}

.card {
    border: none;
    box-shadow: 0 10px 20px rgba(0,0,0,0.05);
    border-radius: 16px;
    margin-bottom: 2rem;
    background: #fff;
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
    margin-bottom: 1.75rem;
}

.form-label {
    font-weight: 600;
    color: #2C3E50;
    margin-bottom: 0.75rem;
    font-size: 0.95rem;
}

.form-control {
    border-radius: 8px;
    border: 2px solid #E9ECEF;
    padding: 0.875rem 1rem;
    font-size: 1rem;
    transition: all 0.2s ease;
}

.form-control:focus {
    border-color: #4A90E2;
    box-shadow: 0 0 0 4px rgba(74,144,226,0.1);
}

.form-select {
    border-radius: 8px;
    border: 2px solid #E9ECEF;
    padding: 0.875rem 1rem;
    font-size: 1rem;
    background-position: right 1rem center;
}

.required-mark {
    color: #E74C3C;
    margin-left: 4px;
}

/* 결재선 표 스타일 */
.approval-line-table-container {
    border: 2px solid #E9ECEF;
    border-radius: 12px;
    padding: 1.5rem;
    margin-bottom: 1.75rem;
    background-color: #fff;
    transition: all 0.2s ease;
}

.approval-line-table-container:hover {
    border-color: #4A90E2;
    box-shadow: 0 4px 12px rgba(74,144,226,0.1);
    transform: translateY(-2px);
}

.approval-line-table-container::after {
    
    position: absolute;
    bottom: 10px;
    right: 15px;
    font-size: 0.85rem;
    color: #94A3B8;
    font-style: italic;
}

.approval-line-table-container {
    position: relative;
}

.approval-line-table {
    width: 100%;
    border-collapse: collapse;
    margin-bottom: 0;
}

.approval-cell {
    border: 2px solid #2C3E50;
    padding: 1rem;
    text-align: center;
    vertical-align: middle;
    background: #fff;
    position: relative;
    min-height: 80px;
    width: 20%;
}

/* 첫 번째 행 (직급) 스타일 */
.position-row .approval-cell {
    height: 50px;
    font-weight: 600;
    color: #2C3E50;
    font-size: 1rem;
    background: #F8FAFC;
}

/* 두 번째 행 (이름과 순서) 스타일 */
.name-row .approval-cell {
    height: 80px;
    position: relative;
    padding: 1.5rem 1rem;
}

.step-number {
    position: absolute;
    top: 8px;
    left: 8px;
    background: #4A90E2;
    color: white;
    width: 20px;
    height: 20px;
    border-radius: 50%;
    display: flex;
    align-items: center;
    justify-content: center;
    font-size: 0.8rem;
    font-weight: 600;
}

.name-cell .approver-name {
    font-weight: 600;
    color: #2C3E50;
    font-size: 1rem;
    margin-top: 10px;
    line-height: 1.2;
}

/* 미지정 상태 스타일 */
.name-cell .approver-name:contains("미지정") {
    color: #94A3B8;
    font-style: italic;
}

#approverName2:empty::after,
#approverName3:empty::after,
#approverName4:empty::after,
#approverName5:empty::after {
    content: "미지정";
    color: #94A3B8;
    font-style: italic;
}

.btn {
    padding: 0.875rem 1.5rem;
    font-weight: 600;
    border-radius: 8px;
    transition: all 0.2s ease;
    letter-spacing: 0.3px;
}

.btn-outline-primary {
    border-width: 2px;
    border-color: #4A90E2;
    color: #4A90E2;
}

.btn-outline-primary:hover {
    background: #4A90E2;
    color: #fff;
    transform: translateY(-2px);
    box-shadow: 0 4px 12px rgba(74,144,226,0.2);
}

.btn-primary {
    background: linear-gradient(to right, #4A90E2, #5A9EEA);
    border: none;
    box-shadow: 0 4px 12px rgba(74,144,226,0.2);
}

.btn-primary:hover {
    transform: translateY(-2px);
    box-shadow: 0 6px 16px rgba(74,144,226,0.3);
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

.d-grid .btn {
    padding: 1rem 2rem;
    font-size: 1.05rem;
}

.d-grid {
    gap: 1rem;
}

.alert {
    border: none;
    border-radius: 12px;
    padding: 1.25rem 1.5rem;
    margin-bottom: 2rem;
    font-weight: 500;
    box-shadow: 0 4px 12px rgba(0,0,0,0.05);
}

.alert-danger {
    background: #FDF1F1;
    color: #E74C3C;
}

.btn-sm {
    padding: 0.625rem 1.25rem;
    font-size: 0.9rem;
}
</style>

<div class="container mt-4">
    <div class="card">
        <div class="card-header bg-white py-3">
            <h5 class="card-title mb-0">카드 사용 내역 등록</h5>
        </div>
        <div class="card-body">
            <% If errorMsg <> "" Then %>
                <div class="alert alert-danger" role="alert">
                    <%= errorMsg %>
                </div>
            <% End If %>
            
            <form method="post" action="card_usage_add.asp" id="cardUsageForm">
                <!-- 결재선 지정 영역 -->
                <div class="mb-4">
                    <label class="form-label">결재선 지정</label>
                    <div class="approval-line-table-container" onclick="openApprovalLinePopup()" style="cursor: pointer;">
                        <table class="approval-line-table">
                            <tbody>
                                <!-- 첫 번째 행: 직급 -->
                                <tr class="position-row">
                                    <td class="approval-cell" id="position1">
                                        <%= userJobGradeName %>
                                    </td>
                                    <td class="approval-cell" id="position2">
                                        <!-- 2차 결재자 직급 -->
                                    </td>
                                    <td class="approval-cell" id="position3">
                                        <!-- 3차 결재자 직급 -->
                                    </td>
                                    <td class="approval-cell" id="position4">
                                        <!-- 4차 결재자 직급 -->
                                    </td>
                                    <td class="approval-cell" id="position5">
                                        <!-- 5차 결재자 직급 -->
                                    </td>
                                </tr>
                                <!-- 두 번째 행: 이름과 순서 -->
                                <tr class="name-row">
                                    <td class="approval-cell name-cell">
                                        <span class="step-number">1</span>
                                        <div class="approver-name"><%= Session("name") %></div>
                                    </td>
                                    <td class="approval-cell name-cell" id="nameCell2">
                                        <span class="step-number">2</span>
                                        <div class="approver-name" id="approverName2">미지정</div>
                                    </td>
                                    <td class="approval-cell name-cell" id="nameCell3">
                                        <span class="step-number">3</span>
                                        <div class="approver-name" id="approverName3">미지정</div>
                                    </td>
                                    <td class="approval-cell name-cell" id="nameCell4">
                                        <span class="step-number">4</span>
                                        <div class="approver-name" id="approverName4">미지정</div>
                                    </td>
                                    <td class="approval-cell name-cell" id="nameCell5">
                                        <span class="step-number">5</span>
                                        <div class="approver-name" id="approverName5">미지정</div>
                                    </td>
                                </tr>
                            </tbody>
                        </table>

                    </div>
                </div>

                <!-- 숨겨진 결재자 입력 필드들 -->
                <input type="hidden" name="approver_step1" id="approver_step1" value="<%= Session("user_id") %>">
                <input type="hidden" name="approver_step2" id="approver_step2" value="">
                <input type="hidden" name="approver_step3" id="approver_step3" value="">
                <input type="hidden" name="approver_step4" id="approver_step4" value="">
                <input type="hidden" name="approver_step5" id="approver_step5" value="">
                <!-- 카드 사용 내역 정보 테이블 -->
                <div class="approval-line-table-container" style="margin-top: 1rem;">
                    <table class="approval-line-table">
                        <tbody>
                            <tr>
                                <td class="approval-cell" style="background: #F8FAFC; font-weight: 600; width: 20%;">제목 <span style="color: #E74C3C;">*</span></td>
                                <td class="approval-cell" colspan="4" style="text-align: left; padding: 1rem;">
                                    <input type="text" class="form-control" id="title" name="title" placeholder="결재 제목을 입력하세요" required style="border: 1px solid #E9ECEF; width: 100%;">
                                </td>
                            </tr>
                            <tr>
                                <td class="approval-cell" style="background: #F8FAFC; font-weight: 600;">카드 <span style="color: #E74C3C;">*</span></td>
                                <td class="approval-cell" style="text-align: left; padding: 1rem;">
                                    <select class="form-control" id="card_id" name="card_id" required style="border: 1px solid #E9ECEF; width: 100%;">
                                        <option value="">카드 선택</option>
                                        <% 
                                        If Not cardRS.EOF Then
                                            Do While Not cardRS.EOF 
                                        %>
                                            <option value="<%= cardRS("card_id") %>"><%= cardRS("account_name") %> (<%= cardRS("issuer") %>)</option>
                                        <% 
                                            cardRS.MoveNext
                                            Loop
                                        End If
                                        %>
                                    </select>
                                </td>
                                <td class="approval-cell" style="background: #F8FAFC; font-weight: 600;">사용일자 <span style="color: #E74C3C;">*</span></td>
                                <td class="approval-cell" colspan="2" style="text-align: left; padding: 1rem;">
                                    <input type="date" class="form-control" id="usage_date" name="usage_date" required style="border: 1px solid #E9ECEF; width: 100%;">
                                </td>
                            </tr>
                            <tr>
                                <td class="approval-cell" style="background: #F8FAFC; font-weight: 600;">사용처 <span style="color: #E74C3C;">*</span></td>
                                <td class="approval-cell" style="text-align: left; padding: 1rem;">
                                    <input type="text" class="form-control" id="store_name" name="store_name" required style="border: 1px solid #E9ECEF; width: 100%;">
                                </td>
                                <td class="approval-cell" style="background: #F8FAFC; font-weight: 600;">금액 <span style="color: #E74C3C;">*</span></td>
                                <td class="approval-cell" colspan="2" style="text-align: left; padding: 1rem;">
                                    <input type="text" class="form-control" id="amount" name="amount" required pattern="^\d{1,3}(,\d{3})*$" inputmode="numeric" onkeyup="cleanNumberInput(this);" style="border: 1px solid #E9ECEF; width: 100%;">
                                </td>
                            </tr>
                            <tr>
                                <td class="approval-cell" style="background: #F8FAFC; font-weight: 600;">계정과목 <span style="color: #E74C3C;">*</span></td>
                                <td class="approval-cell" colspan="4" style="text-align: left; padding: 1rem;">
                                    <select class="form-control" id="account_type_id" name="account_type_id" required style="border: 1px solid #E9ECEF; width: 100%;">
                                        <option value="">계정 선택</option>
                                        <% 
                                        If Not accountTypeRS.EOF Then
                                            Do While Not accountTypeRS.EOF 
                                        %>
                                            <option value="<%= accountTypeRS("account_type_id") %>"><%= accountTypeRS("type_name") %></option>
                                        <% 
                                            accountTypeRS.MoveNext
                                            Loop
                                        End If
                                        %>
                                    </select>
                                </td>
                            </tr>
                            <tr>
                                <td class="approval-cell" style="background: #F8FAFC; font-weight: 600;">사용 목적 <span style="color: #E74C3C;">*</span></td>
                                <td class="approval-cell" colspan="4" style="text-align: left; padding: 1rem;">
                                    <textarea class="form-control" id="purpose" name="purpose" rows="3" required style="border: 1px solid #E9ECEF; width: 100%; resize: vertical;"></textarea>
                                </td>
                            </tr>
                        </tbody>
                    </table>
                </div>
                
                <div class="d-grid gap-2">
                    <button type="submit" class="btn btn-primary">등록</button>
                    <a href="card_usage.asp" class="btn btn-secondary">취소</a>
                </div>
            </form>
        </div>
    </div>
</div>

<script>
// 결재선 데이터 저장 변수
let approvalLineData = null;

// 결재선 지정 팝업 열기
function openApprovalLinePopup() {
    const width = 1200;
    const height = 800;
    const left = (screen.width - width) / 2;
    const top = (screen.height - height) / 2;
    
    // 현재 사용자의 결재선 데이터만 팝업으로 전달
    // 세션 스토리지 초기화 후 현재 데이터만 저장
    sessionStorage.removeItem('currentApprovalLine');
    
    if (approvalLineData && approvalLineData.length > 0) {
        // 현재 사용자 ID와 일치하는 결재선만 전달
        const currentUserId = '<%= Session("user_id") %>';
        if (approvalLineData[0] && approvalLineData[0].userId === currentUserId) {
            sessionStorage.setItem('currentApprovalLine', JSON.stringify(approvalLineData));
        }
    }
    
    window.open('approval_line_popup.asp', 'approvalLinePopup',
        `width=${width},height=${height},left=${left},top=${top},scrollbars=yes`);
}

// 결재선 데이터 설정 (팝업에서 호출)
function setApprovalLine(data) {
    if (!data || !Array.isArray(data) || data.length < 1) {
        alert('올바른 결재선을 지정해주세요.');
        return;
    }

    approvalLineData = data;
    updateApprovalLineDisplay();
    updateHiddenFields();
}

// 결재선 표시 업데이트
function updateApprovalLineDisplay() {
    console.log('결재선 데이터:', approvalLineData); // 디버깅용
    
    // 1차 결재자(본인) 직급 정보 업데이트
    if (approvalLineData && approvalLineData.length > 0) {
        const firstApprover = approvalLineData[0];
        if (firstApprover.jobGradeName) {
            document.getElementById('position1').textContent = firstApprover.jobGradeName;
        }
    }
    
    // 실제 결재자 수만큼만 표시
    const approverCount = approvalLineData ? approvalLineData.length : 1;
    
    // 모든 열을 먼저 숨기기
    for (let i = 1; i <= 5; i++) {
        const positionCell = document.getElementById('position' + i);
        const nameCell = document.getElementById('nameCell' + i);
        
        if (i <= approverCount) {
            // 보여줄 열
            if (positionCell) positionCell.style.display = '';
            if (nameCell) nameCell.style.display = '';
            
            if (i === 1) {
                // 1차 결재자는 항상 표시
                continue;
            }
            
            if (approvalLineData && approvalLineData.length > i - 1) {
                const approver = approvalLineData[i - 1];
                console.log(`${i}차 결재자:`, approver); // 디버깅용
                
                document.getElementById('position' + i).textContent = approver.jobGradeName || '';
                document.getElementById('approverName' + i).textContent = approver.userName || '';
                document.getElementById('nameCell' + i).style.color = '#2C3E50';
            }
        } else {
            // 숨길 열
            if (positionCell) positionCell.style.display = 'none';
            if (nameCell) nameCell.style.display = 'none';
        }
    }
    
    // 테이블 셀 너비 조정
    const cellWidth = 100 / approverCount;
    for (let i = 1; i <= approverCount; i++) {
        const positionCell = document.getElementById('position' + i);
        const nameCell = document.getElementById('nameCell' + i);
        if (positionCell) positionCell.style.width = cellWidth + '%';
        if (nameCell) nameCell.style.width = cellWidth + '%';
    }
}

// 숨겨진 입력 필드 업데이트
function updateHiddenFields() {
    // 1차 결재자는 이미 설정되어 있음
    for (let i = 2; i <= 5; i++) {
        const approver = approvalLineData[i - 1];
        document.getElementById('approver_step' + i).value = approver?.userId || '';
    }
}

// 폼 제출 전 유효성 검사
document.getElementById('cardUsageForm').addEventListener('submit', function(e) {
    // 1차 결재자(본인)만으로도 등록 가능하므로 추가 결재자 검증 제거
    // 필요시 다른 유효성 검사 추가 가능
});

// 숫자 입력 필드에서 쉼표 제거하는 함수
function cleanNumberInput(input) {
    // 현재 커서 위치 저장
    const start = input.selectionStart;
    const end = input.selectionEnd;
    
    // 입력된 값에서 쉼표 제거
    let value = input.value.replace(/,/g, '');
    
    // 숫자만 허용
    value = value.replace(/[^\d]/g, '');
    
    // 빈 값이 아닌 경우에만 포맷팅
    if (value) {
        // 포맷팅 전의 길이와 커서 위치 저장
        const beforeLen = value.length;
        const cursorPos = start;
        
        // 천단위 콤마 추가
        value = Number(value).toLocaleString('ko-KR');
        
        // 값 갱신
        input.value = value;
        
        // 커서 위치 복원
        const newCursorPos = cursorPos + Math.floor((cursorPos - 1) / 3);
        setTimeout(() => {
            input.setSelectionRange(newCursorPos, newCursorPos);
        }, 0);
    } else {
        input.value = '';
    }
}


// 페이지 로드 시 이벤트 리스너 등록
document.addEventListener('DOMContentLoaded', function() {
    // 결재선 데이터 초기화 (다른 사용자 데이터 방지)
    approvalLineData = null;
    sessionStorage.removeItem('currentApprovalLine');
    
    const amountField = document.getElementById('amount');
    if (amountField) {
        // 입력 시 실시간으로 포맷팅
        amountField.addEventListener('input', function(e) {
            const cursorPos = this.selectionStart;
            let value = this.value.replace(/[^\d,]/g, '');
            
            // 콤마 제거 후 숫자만 남김
            value = value.replace(/,/g, '');
            
            if (value) {
                // 천단위 콤마 추가
                this.value = Number(value).toLocaleString('ko-KR');
                
                // 커서 위치 조정
                const newCursorPos = cursorPos + Math.floor((cursorPos - 1) / 3);
                setTimeout(() => {
                    this.setSelectionRange(newCursorPos, newCursorPos);
                }, 0);
            }
        });
        
        // 포커스 아웃 시 최종 포맷팅
        amountField.addEventListener('blur', function() {
            if (this.value) {
                const value = this.value.replace(/,/g, '');
                this.value = Number(value).toLocaleString('ko-KR');
            }
        });
    }
    
    prepareFormSubmission();
});

// 사용일자 기본값을 오늘로 설정
document.getElementById('usage_date').valueAsDate = new Date();
</script>

<!--#include file="../includes/footer.asp"--> 