<%@ Language="VBScript" CodePage="65001" %>
<% 
Response.CodePage = 65001
Response.CharSet = "utf-8"
%>

<!--#include virtual="/card_car_used/db.asp"-->
<!--#include virtual="/card_car_used/includes/functions.asp"-->
<%
' 로그인 체크
If Not IsAuthenticated() Then
    Response.Write "로그인이 필요합니다."
    Response.End
End If

' 현재 사용자의 직급 정보 가져오기
Dim currentUserJobGrade
currentUserJobGrade = ""
If Session("user_id") <> "" Then
    Dim currentUserSQL, currentUserRS
    currentUserSQL = "SELECT j.name as job_grade_name FROM dbo.Users u " & _
                     "LEFT JOIN dbo.job_grade j ON u.job_grade = j.job_grade_id " & _
                     "WHERE u.user_id = '" & Session("user_id") & "'"
    Set currentUserRS = db99.Execute(currentUserSQL)
    If Not currentUserRS.EOF And Not IsNull(currentUserRS("job_grade_name")) Then
        currentUserJobGrade = currentUserRS("job_grade_name")
    End If
    If Not currentUserRS Is Nothing Then
        If currentUserRS.State = 1 Then currentUserRS.Close
        Set currentUserRS = Nothing
    End If
End If

' 팝업 크기 조정을 위한 스크립트
Response.Write "<script>window.resizeTo(1200, 800);</script>"

' 부서 목록 조회
Dim deptSQL, deptRS
deptSQL = "SELECT department_id, name FROM dbo.Department ORDER BY name"
Set deptRS = db99.Execute(deptSQL)

' 사용자 목록 조회 (부서별, 관리자 제외)
Dim userSQL, userRS
userSQL = "SELECT u.user_id, u.name, u.department_id, d.name as dept_name, j.name as job_grade_name " & _
          "FROM dbo.Users u " & _
          "LEFT JOIN dbo.Department d ON u.department_id = d.department_id " & _
          "LEFT JOIN dbo.job_grade j ON u.job_grade = j.job_grade_id " & _
          "WHERE u.user_id != 'admin'" & _
          "ORDER BY d.name, j.job_grade_id DESC, u.name"
Set userRS = db99.Execute(userSQL)
%>

<!DOCTYPE html>
<html>
<head>
    <title>결재선 지정</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.2.3/dist/css/bootstrap.min.css" rel="stylesheet">
    <link href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css" rel="stylesheet">
</head>
<body class="bg-light">
    <div class="container-fluid p-4">
        <div class="row">
            <!-- 부서별 사용자 목록 -->
            <div class="col-md-8">
                <div class="card">
                    <div class="card-header py-3">
                        <h5 class="card-title mb-0">부서별 사용자 목록</h5>
                    </div>
                    <div class="search-box">
                        <input type="text" class="form-control search-input" id="userSearch" placeholder="사용자 이름 또는 직급으로 검색...">
                    </div>
                    <div class="card-body p-0">
                        <div class="accordion" id="departmentAccordion">
                            <% Do While Not deptRS.EOF %>
                                <div class="accordion-item">
                                    <h2 class="accordion-header">
                                        <button class="accordion-button" type="button" data-bs-toggle="collapse" 
                                                data-bs-target="#dept_<%= deptRS("department_id") %>">
                                            <%= deptRS("name") %>
                                        </button>
                                    </h2>
                                    <div id="dept_<%= deptRS("department_id") %>" class="accordion-collapse collapse show">
                                        <div class="accordion-body p-0">
                                            <div class="list-group list-group-flush user-list">
                                                <% 
                                                userRS.MoveFirst
                                                Do While Not userRS.EOF 
                                                    If CStr(userRS("department_id")) = CStr(deptRS("department_id")) Then
                                                %>
                                                    <div class="list-group-item user-item" draggable="true" 
                                                         data-user-id="<%= userRS("user_id") %>"
                                                         data-user-name="<%= userRS("name") %>"
                                                         data-dept-name="<%= userRS("dept_name") %>"
                                                         data-job-grade="<%= userRS("job_grade_name") %>">
                                                        <div class="d-flex align-items-center">
                                                            <div>
                                                                <h6 class="mb-0"><%= userRS("name") %></h6>
                                                                <small class="text-muted">
                                                                    <%= userRS("job_grade_name") %>
                                                                </small>
                                                            </div>
                                                        </div>
                                                    </div>
                                                <%
                                                    End If
                                                    userRS.MoveNext
                                                Loop
                                                %>
                                            </div>
                                        </div>
                                    </div>
                                </div>
                            <%
                                deptRS.MoveNext
                            Loop
                            %>
                        </div>
                    </div>
                </div>
            </div>

            <!-- 결재선 지정 영역 -->
            <div class="col-md-4">
                <div class="card">
                    <div class="card-header py-3">
                        <h5 class="card-title mb-0">결재선 지정</h5>
                        <small class="text-white-50">드래그하거나 더블클릭으로 결재자를 추가하세요</small>
                    </div>
                    <div class="card-body">
                        <div class="approval-line-container">
                            <!-- 1차 결재자 (본인) -->
                            <div class="approval-step mb-3">
                                <div class="d-flex align-items-center mb-2">
                                    <span class="badge bg-primary me-2">1차 결재</span>
                                    <small class="text-muted">(본인)</small>
                                </div>
                                <div class="approval-box current-user" id="step1">
                                    <div class="d-flex align-items-center">
                                        <div>
                                            <h6 class="mb-0"><%= Session("name") %></h6>
                                            <small class="text-muted"><%= Session("department_name") %><% If currentUserJobGrade <> "" Then %> / <%= currentUserJobGrade %><% End If %></small>
                                        </div>
                                    </div>
                                </div>
                            </div>

                            <!-- 추가 결재자 영역 -->
                            <div id="approvalSteps">
                                <div class="drag-hint" id="dragHint">
                                    <i class="fas fa-hand-pointer mb-2" style="font-size: 2rem; color: #4A90E2;"></i>
                                    <p class="mb-0">왼쪽에서 결재자를 드래그하거나<br>더블클릭하여 추가하세요</p>
                                </div>
                            </div>

                            <!-- 확인 버튼 -->
                            <div class="d-grid gap-2 mt-4">
                                <button type="button" class="btn btn-primary" onclick="saveApprovalLine()">
                                    결재선 저장
                                </button>
                                <button type="button" class="btn btn-secondary" onclick="window.close()">
                                    취소
                                </button>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
        </div>
    </div>

<style>
body {
    font-family: 'Pretendard', 'Noto Sans KR', sans-serif;
    background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
    min-height: 100vh;
}

.container-fluid {
    background: rgba(255, 255, 255, 0.95);
    border-radius: 20px;
    margin: 20px;
    box-shadow: 0 20px 40px rgba(0, 0, 0, 0.1);
    backdrop-filter: blur(10px);
}

.user-item {
    cursor: pointer;
    transition: all 0.3s ease;
    user-select: none;
    border-radius: 12px;
    margin: 4px 0;
    border: 2px solid transparent;
    background: linear-gradient(145deg, #ffffff, #f8f9fa);
    box-shadow: 0 4px 8px rgba(0, 0, 0, 0.05);
}

.user-item:hover {
    background: linear-gradient(145deg, #4A90E2, #5A9EEA);
    color: white;
    transform: translateY(-2px);
    box-shadow: 0 8px 16px rgba(74, 144, 226, 0.3);
    border-color: #4A90E2;
}

.user-item:hover .text-muted {
    color: rgba(255, 255, 255, 0.8) !important;
}

.user-item.selected {
    background: linear-gradient(145deg, #28a745, #34ce57);
    color: white;
    border-color: #28a745;
    transform: translateY(-2px);
    box-shadow: 0 8px 16px rgba(40, 167, 69, 0.3);
}

.user-item.selected .text-muted {
    color: rgba(255, 255, 255, 0.8) !important;
}

.approval-box {
    border: 2px solid #e9ecef;
    border-radius: 12px;
    padding: 1.25rem;
    background: linear-gradient(145deg, #ffffff, #f8f9fa);
    margin-bottom: 0.75rem;
    position: relative;
    transition: all 0.3s ease;
    box-shadow: 0 4px 12px rgba(0, 0, 0, 0.05);
}

.approval-box:hover {
    transform: translateY(-2px);
    box-shadow: 0 8px 20px rgba(0, 0, 0, 0.1);
    border-color: #4A90E2;
}

.approval-box.dragging {
    opacity: 0.6;
    transform: rotate(5deg);
}

.approval-box.dragover {
    border: 2px dashed #4A90E2;
    background: linear-gradient(145deg, #E3F2FD, #BBDEFB);
    transform: scale(1.02);
}

.approval-box .remove-btn {
    position: absolute;
    right: 0.75rem;
    top: 50%;
    transform: translateY(-50%);
    cursor: pointer;
    color: #dc3545;
    opacity: 0;
    transition: all 0.3s ease;
    background: rgba(220, 53, 69, 0.1);
    border-radius: 50%;
    width: 32px;
    height: 32px;
    display: flex;
    align-items: center;
    justify-content: center;
}

.approval-box:hover .remove-btn {
    opacity: 1;
}

.remove-btn:hover {
    background: rgba(220, 53, 69, 0.2);
    transform: translateY(-50%) scale(1.1);
}

.current-user {
    background: linear-gradient(145deg, #E3F2FD, #BBDEFB);
    border-color: #4A90E2;
    border-width: 2px;
}

.badge {
    font-weight: 600;
    padding: 0.5rem 1rem;
    border-radius: 20px;
    background: linear-gradient(145deg, #4A90E2, #5A9EEA);
    border: none;
    box-shadow: 0 2px 8px rgba(74, 144, 226, 0.3);
}

.accordion-button {
    background: linear-gradient(145deg, #ffffff, #f8f9fa);
    border: none;
    border-radius: 12px !important;
    margin: 4px 0;
    font-weight: 600;
    color: #2C3E50;
    box-shadow: 0 2px 8px rgba(0, 0, 0, 0.05);
    transition: all 0.3s ease;
}

.accordion-button:not(.collapsed) {
    background: linear-gradient(145deg, #4A90E2, #5A9EEA);
    color: white;
    box-shadow: 0 4px 12px rgba(74, 144, 226, 0.3);
}

.accordion-button:hover {
    transform: translateY(-1px);
    box-shadow: 0 4px 12px rgba(0, 0, 0, 0.1);
}

.accordion-item {
    border: none;
    margin-bottom: 8px;
}

.accordion-body {
    border-radius: 0 0 12px 12px;
    background: rgba(255, 255, 255, 0.8);
}

.list-group-item {
    border: none;
    background: transparent;
}

.card {
    border: none;
    box-shadow: 0 10px 30px rgba(0, 0, 0, 0.1);
    border-radius: 20px;
    background: rgba(255, 255, 255, 0.95);
    backdrop-filter: blur(10px);
}

.card-header {
    background: linear-gradient(145deg, #4A90E2, #5A9EEA);
    color: white;
    border: none;
    border-radius: 20px 20px 0 0 !important;
    padding: 1.5rem;
}

.card-title {
    font-weight: 700;
    font-size: 1.25rem;
}

.btn-primary {
    background: linear-gradient(145deg, #4A90E2, #5A9EEA);
    border: none;
    border-radius: 12px;
    padding: 0.75rem 1.5rem;
    font-weight: 600;
    transition: all 0.3s ease;
    box-shadow: 0 4px 12px rgba(74, 144, 226, 0.3);
}

.btn-primary:hover {
    transform: translateY(-2px);
    box-shadow: 0 6px 16px rgba(74, 144, 226, 0.4);
}

.btn-secondary {
    background: linear-gradient(145deg, #6c757d, #7c848c);
    border: none;
    border-radius: 12px;
    padding: 0.75rem 1.5rem;
    font-weight: 600;
    transition: all 0.3s ease;
    box-shadow: 0 4px 12px rgba(108, 117, 125, 0.3);
}

.btn-secondary:hover {
    transform: translateY(-2px);
    box-shadow: 0 6px 16px rgba(108, 117, 125, 0.4);
}

.approval-step {
    position: relative;
}

.approval-step::before {
    content: '';
    position: absolute;
    left: -10px;
    top: 50%;
    width: 4px;
    height: 100%;
    background: linear-gradient(to bottom, #4A90E2, #5A9EEA);
    border-radius: 2px;
    transform: translateY(-50%);
}

.search-box {
    position: sticky;
    top: 0;
    z-index: 10;
    background: rgba(255, 255, 255, 0.95);
    backdrop-filter: blur(10px);
    padding: 1rem;
    border-radius: 12px;
    margin-bottom: 1rem;
    box-shadow: 0 2px 8px rgba(0, 0, 0, 0.1);
}

.search-input {
    border: 2px solid #e9ecef;
    border-radius: 12px;
    padding: 0.75rem 1rem;
    transition: all 0.3s ease;
}

.search-input:focus {
    border-color: #4A90E2;
    box-shadow: 0 0 0 4px rgba(74, 144, 226, 0.1);
}
</style>

<script src="https://cdn.jsdelivr.net/npm/bootstrap@5.2.3/dist/js/bootstrap.bundle.min.js"></script>
<script>
    // 드래그 앤 드롭 관련 변수
    let draggedItem = null;
    let approvalStepCount = 1; // 1차 결재자(본인)가 이미 있으므로 1부터 시작

    // 페이지 로드 시 기존 결재선 데이터 로드
    document.addEventListener('DOMContentLoaded', function() {
        loadExistingApprovalLine();
        setupSearchFunction();
    });

    // 기존 결재선 데이터 로드
    function loadExistingApprovalLine() {
        try {
            const existingData = sessionStorage.getItem('currentApprovalLine');
            if (existingData) {
                const approvalLine = JSON.parse(existingData);
                console.log('기존 결재선 데이터 로드:', approvalLine);
                
                // 현재 사용자 ID 확인
                const currentUserId = '<%= Session("user_id") %>';
                
                // 1차 결재자가 현재 사용자인지 확인
                if (approvalLine.length > 0 && approvalLine[0].userId === currentUserId) {
                    // 1차 결재자(본인) 제외하고 2차부터 추가
                    for (let i = 1; i < approvalLine.length; i++) {
                        const approver = approvalLine[i];
                        approvalStepCount++;
                        
                        const newApprovalBox = createApprovalBox({
                            userId: approver.userId,
                            userName: approver.userName,
                            deptName: approver.deptName,
                            jobGrade: approver.jobGradeName,
                            step: approvalStepCount
                        });
                        
                        document.getElementById('approvalSteps').appendChild(newApprovalBox);
                    }
                    
                    // 결재자가 있으면 힌트 메시지 숨기기
                    if (approvalLine.length > 1) {
                        const dragHint = document.getElementById('dragHint');
                        if (dragHint) {
                            dragHint.style.display = 'none';
                        }
                    }
                    
                    updateApprovalSteps();
                } else {
                    console.log('다른 사용자의 결재선 데이터이므로 로드하지 않습니다.');
                }
                
                // 세션 스토리지에서 데이터 제거
                sessionStorage.removeItem('currentApprovalLine');
            }
        } catch (error) {
            console.error('기존 결재선 데이터 로드 중 오류:', error);
            // 오류 발생 시 세션 스토리지 정리
            sessionStorage.removeItem('currentApprovalLine');
        }
    }

    // 검색 기능 설정
    function setupSearchFunction() {
        const searchInput = document.getElementById('userSearch');
        searchInput.addEventListener('input', function() {
            const searchTerm = this.value.toLowerCase();
            const userItems = document.querySelectorAll('.user-item');
            
            userItems.forEach(item => {
                const userName = item.querySelector('h6').textContent.toLowerCase();
                const jobGrade = item.querySelector('small').textContent.toLowerCase();
                
                if (userName.includes(searchTerm) || jobGrade.includes(searchTerm)) {
                    item.style.display = '';
                    item.parentElement.style.display = '';
                } else {
                    item.style.display = 'none';
                }
            });
            
            // 부서별로 숨김/표시 처리
            document.querySelectorAll('.accordion-item').forEach(deptItem => {
                const visibleUsers = deptItem.querySelectorAll('.user-item[style=""], .user-item:not([style])');
                if (visibleUsers.length === 0) {
                    deptItem.style.display = 'none';
                } else {
                    deptItem.style.display = '';
                }
            });
        });
    }

    // 사용자 목록 이벤트 처리
    document.querySelectorAll('.user-item').forEach(item => {
        // 드래그 이벤트
        item.addEventListener('dragstart', handleDragStart);
        item.addEventListener('dragend', handleDragEnd);
        
        // 더블클릭 이벤트
        item.addEventListener('dblclick', function() {
            addApprover(this.dataset);
        });
    });

    // 결재선 영역 드래그 이벤트 처리
    const approvalSteps = document.getElementById('approvalSteps');
    approvalSteps.addEventListener('dragover', handleDragOver);
    approvalSteps.addEventListener('drop', handleDrop);

    function handleDragStart(e) {
        draggedItem = this;
        this.classList.add('dragging');
        e.dataTransfer.effectAllowed = 'move';
    }

    function handleDragEnd(e) {
        draggedItem = null;
        this.classList.remove('dragging');
    }

    function handleDragOver(e) {
        e.preventDefault();
        e.dataTransfer.dropEffect = 'move';
    }

    function handleDrop(e) {
        e.preventDefault();
        if (!draggedItem) return;
        
        addApprover(draggedItem.dataset);
    }

    function addApprover(data) {
        // 이미 추가된 결재자인지 확인
        const existingApprovers = document.querySelectorAll('#approvalSteps .approval-box');
        for (let approver of existingApprovers) {
            if (approver.dataset.userId === data.userId) {
                alert('이미 추가된 결재자입니다.');
                return;
            }
        }

        // 본인을 결재선에 추가하려는 경우 방지
        if (data.userId === '<%= Session("user_id") %>') {
            alert('본인은 이미 1차 결재자로 지정되어 있습니다.');
            return;
        }

        // 힌트 메시지 숨기기
        const dragHint = document.getElementById('dragHint');
        if (dragHint) {
            dragHint.style.display = 'none';
        }

        approvalStepCount++;
        
        // 새로운 결재자 박스 생성
        const newApprovalBox = createApprovalBox({
            userId: data.userId,
            userName: data.userName,
            deptName: data.deptName,
            jobGrade: data.jobGrade,
            step: approvalStepCount
        });

        approvalSteps.appendChild(newApprovalBox);
        updateApprovalSteps();
    }

    function createApprovalBox(data) {
        const div = document.createElement('div');
        div.className = 'approval-step mb-3';
        div.innerHTML = `
            <div class="d-flex align-items-center mb-2">
                <span class="badge bg-primary me-2">${data.step}차 결재</span>
            </div>
            <div class="approval-box" data-user-id="${data.userId}">
                <div class="d-flex align-items-center justify-content-between">
                    <div>
                        <h6 class="mb-0">${data.userName}</h6>
                        <small class="text-muted">${data.deptName} / ${data.jobGrade}</small>
                    </div>
                    <button type="button" class="btn btn-link text-danger remove-btn" 
                            onclick="removeApprovalStep(this)">
                        <i class="fas fa-times"></i>
                    </button>
                </div>
            </div>
        `;
        return div;
    }

    function removeApprovalStep(button) {
        const approvalStep = button.closest('.approval-step');
        approvalStep.remove();
        approvalStepCount--;
        updateApprovalSteps();
        
        // 결재자가 모두 제거되면 힌트 메시지 다시 표시
        const remainingApprovers = document.querySelectorAll('#approvalSteps .approval-box');
        const dragHint = document.getElementById('dragHint');
        if (remainingApprovers.length === 0 && dragHint) {
            dragHint.style.display = 'block';
        }
    }

    function updateApprovalSteps() {
        // 결재 단계 순서 업데이트
        const steps = document.querySelectorAll('#approvalSteps .approval-step');
        steps.forEach((step, index) => {
            const badge = step.querySelector('.badge');
            badge.textContent = `${index + 2}차 결재`; // 1차는 본인이므로 2차부터 시작
        });
    }

    function saveApprovalLine() {
        const approvalLine = [];
        
        // 1차 결재자(본인) 추가 - 직급 정보도 포함
        approvalLine.push({
            step: 1,
            userId: '<%= Session("user_id") %>',
            userName: '<%= Session("name") %>',
            deptName: '<%= Session("department_name") %>',
            jobGradeName: '<%= currentUserJobGrade %>'
        });

        // 추가된 결재자들 수집
        document.querySelectorAll('#approvalSteps .approval-box').forEach((box, index) => {
            const smallText = box.querySelector('small').textContent;
            const parts = smallText.split(' / ');
            const deptName = parts[0] || '';
            const jobGradeName = parts[1] || '';
            
            approvalLine.push({
                step: index + 2, // 1차 다음부터 시작
                userId: box.dataset.userId,
                userName: box.querySelector('h6').textContent,
                deptName: deptName,
                jobGradeName: jobGradeName
            });
        });

        // 1차 결재자(본인)만 있어도 저장 가능
        if (approvalLine.length >= 1) {
            try {
                // 부모 창이 존재하는지 확인
                if (!window.opener) {
                    alert('부모 창을 찾을 수 없습니다.');
                    return;
                }

                // 부모 창의 함수 존재 여부 확인
                if (typeof window.opener.setApprovalLine === 'function') {
                    window.opener.setApprovalLine(approvalLine);
                    window.close();
                } else {
                    // 부모 창에 결재선 데이터를 직접 설정
                    window.opener.approvalLineData = approvalLine;
                    if (typeof window.opener.updateApprovalLineDisplay === 'function') {
                        window.opener.updateApprovalLineDisplay();
                    }
                    if (typeof window.opener.updateHiddenFields === 'function') {
                        window.opener.updateHiddenFields();
                    }
                    alert('결재선이 저장되었습니다.');
                    window.close();
                }
            } catch (error) {
                console.error('결재선 저장 중 오류 발생:', error);
                alert('결재선 저장 중 오류가 발생했습니다.\n' + error.message);
            }
        } else {
            alert('최소 1명의 결재자가 필요합니다.');
        }
    }
</script>

</body>
</html>