<%@ Language="VBScript" CodePage="65001" %>
<% 
Response.CodePage = 65001
Response.CharSet = "utf-8"
%>

<!--#include virtual="/contents/card_car_used/db.asp"-->
<!--#include virtual="/contents/card_car_used/includes/functions.asp"-->
<%
' 로그인 체크
If Not IsAuthenticated() Then
    Response.Write "로그인이 필요합니다."
    Response.End
End If

' 팝업 크기 조정을 위한 스크립트
Response.Write "<script>window.resizeTo(1200, 800);</script>"

' 부서 목록 조회
Dim deptSQL, deptRS
deptSQL = "SELECT department_id, name FROM dbo.Department ORDER BY name"
Set deptRS = db99.Execute(deptSQL)

' 사용자 목록 조회 (부서별)
Dim userSQL, userRS
userSQL = "SELECT u.user_id, u.name, u.department_id, d.name as dept_name, j.name as job_grade_name " & _
          "FROM dbo.Users u " & _
          "LEFT JOIN dbo.Department d ON u.department_id = d.department_id " & _
          "LEFT JOIN dbo.job_grade j ON u.job_grade = j.job_grade_id " & _
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
                    <div class="card-header bg-white py-3">
                        <h5 class="card-title mb-0">부서별 사용자 목록</h5>
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
                    <div class="card-header bg-white py-3">
                        <h5 class="card-title mb-0">결재선 지정</h5>
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
                                            <small class="text-muted"><%= Session("department_name") %></small>
                                        </div>
                                    </div>
                                </div>
                            </div>

                            <!-- 추가 결재자 영역 -->
                            <div id="approvalSteps">
                                <!-- 여기에 드래그된 결재자들이 추가됩니다 -->
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
.user-item {
    cursor: pointer;
    transition: all 0.2s;
    user-select: none;
}

.user-item:hover {
    background-color: #f8f9fa;
}

.user-item.selected {
    background-color: #e9ecef;
}

.approval-box {
    border: 1px solid #dee2e6;
    border-radius: 0.5rem;
    padding: 1rem;
    background-color: white;
    margin-bottom: 0.5rem;
    position: relative;
}

.approval-box.dragging {
    opacity: 0.5;
}

.approval-box.dragover {
    border: 2px dashed #4A90E2;
    background-color: #F8F9FA;
}

.approval-box .remove-btn {
    position: absolute;
    right: 0.5rem;
    top: 50%;
    transform: translateY(-50%);
    cursor: pointer;
    color: #dc3545;
    display: none;
}

.approval-box:hover .remove-btn {
    display: block;
}

.current-user {
    background-color: #E3F2FD;
    border-color: #90CAF9;
}

.badge {
    font-weight: 500;
}

.accordion-button:not(.collapsed) {
    background-color: #F8F9FA;
    color: #2C3E50;
}

.list-group-item {
    border-left: none;
    border-right: none;
}

.card {
    border: none;
    box-shadow: 0 0.125rem 0.25rem rgba(0, 0, 0, 0.075);
    border-radius: 0.75rem;
}

.card-header {
    border-bottom: 1px solid #E9ECEF;
    border-radius: 0.75rem 0.75rem 0 0 !important;
}
</style>

<script src="https://cdn.jsdelivr.net/npm/bootstrap@5.2.3/dist/js/bootstrap.bundle.min.js"></script>
<script>
    // 드래그 앤 드롭 관련 변수
    let draggedItem = null;
    let approvalStepCount = 1; // 1차 결재자(본인)가 이미 있으므로 1부터 시작

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
        
        // 1차 결재자(본인) 추가
        approvalLine.push({
            step: 1,
            userId: '<%= Session("user_id") %>',
            userName: '<%= Session("name") %>',
            deptName: '<%= Session("department_name") %>'
        });

        // 추가된 결재자들 수집
        document.querySelectorAll('#approvalSteps .approval-box').forEach((box, index) => {
            approvalLine.push({
                step: index + 2, // 1차 다음부터 시작
                userId: box.dataset.userId,
                userName: box.querySelector('h6').textContent,
                deptName: box.querySelector('small').textContent.split(' / ')[0]
            });
        });

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
                alert('결재선이 저장되었습니다.');
                window.close();
            }
        } catch (error) {
            console.error('결재선 저장 중 오류 발생:', error);
            alert('결재선 저장 중 오류가 발생했습니다.\n' + error.message);
        }
    }
</script>

</body>
</html>