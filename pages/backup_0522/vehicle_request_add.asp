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

' 최신 유류비 단가 조회
Dim fuelRateSQL, fuelRateRS, fuelRate
fuelRateSQL = "SELECT TOP 1 rate FROM FuelRate ORDER BY date DESC"
Set fuelRateRS = db.Execute(fuelRateSQL)

If Err.Number <> 0 Or fuelRateRS.EOF Then
    fuelRate = 2000 ' 기본값 설정
Else
    fuelRate = fuelRateRS("rate")
End If

If Not fuelRateRS Is Nothing Then
    If fuelRateRS.State = 1 Then ' adStateOpen
fuelRateRS.Close
    End If
    Set fuelRateRS = Nothing
End If

' 신청서 등록 처리
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    Dim startDate, endDate, purpose, startLocation, destination, distance, tollFee, parkingFee, totalAmount, errorMsg, successMsg
    Dim approver1, approver2, approver3
    
    startDate = PreventSQLInjection(Request.Form("start_date"))
    endDate = PreventSQLInjection(Request.Form("end_date"))
    purpose = PreventSQLInjection(Request.Form("purpose"))
    startLocation = PreventSQLInjection(Request.Form("start_location"))
    destination = PreventSQLInjection(Request.Form("destination"))
    approver1 = PreventSQLInjection(Request.Form("approver_step1"))
    approver2 = PreventSQLInjection(Request.Form("approver_step2"))
    approver3 = PreventSQLInjection(Request.Form("approver_step3"))
    
    ' 종료일자가 비어있으면 시작일자와 동일하게 설정
    If endDate = "" Then
        endDate = startDate
    End If
    
    ' 숫자 값 안전하게 변환
    distance = 0
    tollFee = 0
    parkingFee = 0
    
    If IsNumeric(Replace(Request.Form("distance"), ",", "")) Then
        distance = CDbl(Replace(Request.Form("distance"), ",", ""))
    End If
    
    If IsNumeric(Replace(Request.Form("toll_fee"), ",", "")) Then
        tollFee = CDbl(Replace(Request.Form("toll_fee"), ",", ""))
    End If
    
    If IsNumeric(Replace(Request.Form("parking_fee"), ",", "")) Then
        parkingFee = CDbl(Replace(Request.Form("parking_fee"), ",", ""))
    End If
    
    ' 총 금액 계산
    totalAmount = (distance * fuelRate) + tollFee + parkingFee
    
    ' 입력값 검증
    If startDate = "" Or endDate = "" Or purpose = "" Or startLocation = "" Or destination = "" Then
        errorMsg = "필수 항목을 모두 입력해주세요."
    ElseIf distance <= 0 Then
        errorMsg = "운행거리는 양수여야 합니다."
    ElseIf approver1 = "" Or approver2 = "" Then
        errorMsg = "최소 1명 이상의 결재자를 지정해주세요."
    Else
        ' 트랜잭션 시작
        db.BeginTrans

        ' 파라미터화된 쿼리 사용 - start_date와 end_date 필드 사용
        Dim cmd
        Set cmd = Server.CreateObject("ADODB.Command")
        cmd.ActiveConnection = db
        
        ' 유류비 계산
        Dim fuelCost
        fuelCost = distance * fuelRate
        
        cmd.CommandText = "INSERT INTO VehicleRequests (user_id, start_date, end_date, purpose, start_location, destination, " & _
                         "distance, toll_fee, parking_fee, approval_status, is_deleted) " & _
                         "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)"
        
        ' 파라미터 추가
        cmd.Parameters.Append cmd.CreateParameter("@user_id", 200, 1, 30, Session("user_id"))
        cmd.Parameters.Append cmd.CreateParameter("@start_date", 7, 1, , startDate)
        cmd.Parameters.Append cmd.CreateParameter("@end_date", 7, 1, , endDate)
        cmd.Parameters.Append cmd.CreateParameter("@purpose", 200, 1, 100, purpose)
        cmd.Parameters.Append cmd.CreateParameter("@start_location", 200, 1, 100, startLocation)
        cmd.Parameters.Append cmd.CreateParameter("@destination", 200, 1, 100, destination)
        cmd.Parameters.Append cmd.CreateParameter("@distance", 6, 1, , distance)
        cmd.Parameters.Append cmd.CreateParameter("@toll_fee", 6, 1, , tollFee)
        cmd.Parameters.Append cmd.CreateParameter("@parking_fee", 6, 1, , parkingFee)
        cmd.Parameters.Append cmd.CreateParameter("@approval_status", 200, 1, 20, "대기")
        cmd.Parameters.Append cmd.CreateParameter("@is_deleted", 11, 1, , 0)
        
        ' 명령 실행
        On Error Resume Next
        cmd.Execute
        
        If Err.Number <> 0 Then
            errorMsg = "차량 이용 신청서 등록 중 오류가 발생했습니다: " & Err.Description
            db.RollbackTrans
        Else
            ' 방금 등록한 신청서 ID 조회
            Dim newRequestIdSQL, newRequestIdRS, newRequestId
            newRequestIdSQL = "SELECT TOP 1 request_id FROM VehicleRequests WHERE user_id = ? ORDER BY request_id DESC"
            
            Dim cmdId
            Set cmdId = Server.CreateObject("ADODB.Command")
            cmdId.ActiveConnection = db
            cmdId.CommandText = newRequestIdSQL
            cmdId.Parameters.Append cmdId.CreateParameter("@user_id", 200, 1, 30, Session("user_id"))
            
            Set newRequestIdRS = cmdId.Execute()
            
            If Not newRequestIdRS.EOF Then
                newRequestId = newRequestIdRS("request_id")
                newRequestIdRS.Close
                
                ' ApprovalLogs 테이블에 결재선 데이터 삽입
                Dim approverIds(2)
                approverIds(0) = approver1
                approverIds(1) = approver2
                approverIds(2) = approver3
                
                Dim i
                For i = 0 To 2
                    If approverIds(i) <> "" Then
                        Dim approvalSQL
                        approvalSQL = "INSERT INTO " & dbSchema & ".ApprovalLogs " & _
                                    "(approver_id, target_table_name, target_id, approval_step, status, created_at) " & _
                                    "VALUES (?, 'VehicleRequests', ?, ?, '대기', GETDATE())"
                        
                        Set cmd = Server.CreateObject("ADODB.Command")
                        cmd.ActiveConnection = db
                        cmd.CommandText = approvalSQL
                        cmd.Parameters.Append cmd.CreateParameter("@approver_id", 200, 1, 30, approverIds(i))
                        cmd.Parameters.Append cmd.CreateParameter("@target_id", 3, 1, , newRequestId)
                        cmd.Parameters.Append cmd.CreateParameter("@approval_step", 3, 1, , i + 1)
                        
                        cmd.Execute
                        
                        If Err.Number <> 0 Then
                            Exit For
                        End If
                    End If
                Next
                
                If Err.Number = 0 Then
                    db.CommitTrans
                    successMsg = "차량 이용 신청서가 성공적으로 등록되었습니다."
                    
                    ' 활동 로그 기록
                    LogActivity Session("user_id"), "차량이용신청", "개인차량 이용 신청서 등록 (ID: " & newRequestId & ", 거리: " & distance & "km, 총액: " & FormatNumber(totalAmount) & "원)"
                    
                    ' 신청서 상세 페이지로 리디렉션
                    RedirectTo("vehicle_request_view.asp?id=" & newRequestId)
                Else
                    db.RollbackTrans
                    errorMsg = "결재선 등록 중 오류가 발생했습니다: " & Err.Description
                End If
            Else
                db.RollbackTrans
                errorMsg = "신청서 등록 후 ID를 찾는 데 실패했습니다."
            End If
        End If
        On Error GoTo 0
    End If
End If

On Error GoTo 0
%>
<!--#include file="../includes/header.asp"-->

<div class="vehicle-request-add-container">
    <div class="shadcn-card" style="max-width: 700px; margin: 30px auto;">
        <div class="shadcn-card-header">
            <h2 class="shadcn-card-title">개인차량 이용 신청서 작성</h2>
            <p class="shadcn-card-description">개인차량 이용에 대한 신청서를 작성합니다.</p>
        </div>
        
        <% If errorMsg <> "" Then %>
        <div class="shadcn-alert shadcn-alert-error">
            <div>
                <span class="shadcn-alert-title">오류</span>
                <span class="shadcn-alert-description"><%= errorMsg %></span>
            </div>
        </div>
        <% End If %>
        
        <% If successMsg <> "" Then %>
        <div class="shadcn-alert shadcn-alert-success">
            <div>
                <span class="shadcn-alert-title">성공</span>
                <span class="shadcn-alert-description"><%= successMsg %></span>
            </div>
        </div>
        <% End If %>
        
        <div class="shadcn-card-content">
            <form id="vehicleRequestForm" method="post" action="vehicle_request_add.asp" onsubmit="prepareFormSubmission(); return validateForm('vehicleRequestForm', vehicleRequestRules)">
                <!-- 결재선 지정 영역 -->
                <div class="mb-4">
                    <label class="shadcn-input-label">결재선 지정</label>
                    <div class="approval-line-box">
                        <!-- 1차 결재자 (본인) -->
                        <div class="approval-step">
                            <span class="step-label">1차 결재</span>
                            <div class="approver-info">
                                <div class="approver-name"><%= Session("name") %></div>
                                <div class="approver-dept"><%= Session("department_name") %></div>
                            </div>
                        </div>
                        
                        <!-- 추가 결재자들 -->
                        <div id="approvalLineDisplay"></div>
                        
                        <!-- 결재선 지정 버튼 -->
                        <div class="text-end mt-3">
                            <button type="button" class="shadcn-btn shadcn-btn-outline" onclick="openApprovalLinePopup()">
                                <i class="fas fa-users me-1"></i> 결재선 지정
                            </button>
                        </div>
                    </div>
                </div>
                
                <!-- 숨겨진 결재자 입력 필드들 -->
                <input type="hidden" name="approver_step1" id="approver_step1" value="<%= Session("user_id") %>">
                <input type="hidden" name="approver_step2" id="approver_step2" value="">
                <input type="hidden" name="approver_step3" id="approver_step3" value="">

                <div class="form-row" style="display: flex; gap: 10px;">
                    <div class="form-group" style="flex: 1;">
                        <label class="shadcn-input-label" for="start_date">시작일자</label>
                        <input class="shadcn-input" type="date" id="start_date" name="start_date" value="<%= FormatDate(Date()) %>">
                    </div>
                    
                    <div class="form-group" style="flex: 1;">
                        <label class="shadcn-input-label" for="end_date">종료일자</label>
                        <input class="shadcn-input" type="date" id="end_date" name="end_date" value="<%= FormatDate(Date()) %>">
                    </div>
                </div>
                
                <div class="form-group">
                    <label class="shadcn-input-label" for="purpose">업무 목적</label>
                    <input class="shadcn-input" type="text" id="purpose" name="purpose" placeholder="업무 목적을 입력하세요">
                </div>
                
                <div class="form-group">
                    <label class="shadcn-input-label" for="start_location">출발지</label>
                    <input class="shadcn-input" type="text" id="start_location" name="start_location" placeholder="출발지를 입력하세요">
                </div>
                
                <div class="form-group">
                    <label class="shadcn-input-label" for="destination">목적지</label>
                    <input class="shadcn-input" type="text" id="destination" name="destination" placeholder="목적지를 입력하세요">
                </div>
                
                <div class="form-group">
                    <label class="shadcn-input-label" for="distance">운행거리 (km)</label>
                    <input class="shadcn-input" type="text" id="distance" name="distance" placeholder="운행거리를 입력하세요" onkeyup="cleanNumberInput(this); calculateAmount()">
                </div>
                
                <div class="form-group">
                    <label class="shadcn-input-label" for="fuel_rate">현재 유류비 단가</label>
                    <input class="shadcn-input" type="text" id="fuel_rate" name="fuel_rate" value="<%= fuelRate %>" readonly>
                </div>
                
                <div class="form-group">
                    <label class="shadcn-input-label" for="toll_fee">통행료</label>
                    <input class="shadcn-input" type="text" id="toll_fee" name="toll_fee" placeholder="통행료를 입력하세요" value="0" onkeyup="cleanNumberInput(this); calculateAmount()">
                </div>
                
                <div class="form-group">
                    <label class="shadcn-input-label" for="parking_fee">주차비</label>
                    <input class="shadcn-input" type="text" id="parking_fee" name="parking_fee" placeholder="주차비를 입력하세요" value="0" onkeyup="cleanNumberInput(this); calculateAmount()">
                </div>
                
                <div class="form-group">
                    <label class="shadcn-input-label" for="total_amount_display">총 예상 금액</label>
                    <input class="shadcn-input" type="text" id="total_amount_display" readonly>
                    <input type="hidden" id="total_amount" name="total_amount">
                </div>
                
                <div class="shadcn-card-footer" style="margin-top: 1.5rem;">
                    <button type="submit" class="shadcn-btn shadcn-btn-primary">등록하기</button>
                    <a href="vehicle_request.asp" class="shadcn-btn shadcn-btn-outline">취소</a>
                </div>
            </form>
        </div>
    </div>
</div>

<script>
    const vehicleRequestRules = {
        start_date: {
            required: true,
            message: '시작일자를 입력해주세요.'
        },
        end_date: {
            required: true,
            message: '종료일자를 입력해주세요.'
        },
        purpose: {
            required: true,
            message: '업무 목적을 입력해주세요.'
        },
        start_location: {
            required: true,
            message: '출발지를 입력해주세요.'
        },
        destination: {
            required: true,
            message: '목적지를 입력해주세요.'
        },
        distance: {
            required: true,
            numeric: true,
            message: '운행거리를 숫자로 입력해주세요.'
        },
        approver_step2: {
            required: true,
            message: '최소 1명 이상의 결재자를 지정해주세요.'
        }
    };
    
    // 폼 제출 전 숫자 필드의 쉼표 제거
    function prepareFormSubmission() {
        // 숫자 입력 필드의 쉼표 제거
        const numericFields = ['distance', 'toll_fee', 'parking_fee', 'total_amount'];
        numericFields.forEach(fieldId => {
            const field = document.getElementById(fieldId);
            if (field) {
                field.value = field.value.replace(/,/g, '');
            }
        });
    }
    
    // 총 금액 계산
    function calculateAmount() {
        const distanceInput = document.getElementById('distance').value || '0';
        const fuelRateInput = document.getElementById('fuel_rate').value || '0';
        const tollFeeInput = document.getElementById('toll_fee').value || '0';
        const parkingFeeInput = document.getElementById('parking_fee').value || '0';
        
        // 쉼표 제거 후 숫자로 변환
        const distance = parseFloat(distanceInput.replace(/,/g, '')) || 0;
        const fuelRate = parseFloat(fuelRateInput.replace(/,/g, '')) || 0;
        const tollFee = parseFloat(tollFeeInput.replace(/,/g, '')) || 0;
        const parkingFee = parseFloat(parkingFeeInput.replace(/,/g, '')) || 0;
        
        const fuelAmount = distance * fuelRate;
        const totalAmount = fuelAmount + tollFee + parkingFee;
        
        // total_amount 필드에는 숫자만 저장
        document.getElementById('total_amount').value = totalAmount;
        // 화면에는 포맷된 금액 표시
        document.getElementById('total_amount_display').value = totalAmount.toLocaleString('ko-KR');
    }
    
    // 숫자 입력 필드에서 쉼표 제거하는 함수
    function cleanNumberInput(input) {
        // 현재 선택 위치 저장
        const start = input.selectionStart;
        const end = input.selectionEnd;
        
        // 쉼표(,) 제거 및 숫자만 남기기
        let value = input.value.replace(/,/g, '');
        value = value.replace(/[^\d.]/g, ''); // 숫자와 마침표만 허용
        
        // 천 단위 콤마 추가
        if (value) {
            value = parseFloat(value).toLocaleString('ko-KR', {maximumFractionDigits: 0});
        }
        
        // 입력 값이 바뀌었는지 확인
        const hasChanged = input.value !== value;
        
        // 값 갱신
        input.value = value;
        
        // 선택 위치 복원 (값이 바뀌었을 경우)
        if (hasChanged) {
            // 콤마가 추가된 경우 위치 조정 필요
            const newCursorPos = Math.max(
                0,
                value.length - (input.value.length - end)
            );
            input.setSelectionRange(newCursorPos, newCursorPos);
        }
    }
    
    // 결재선 지정 팝업 열기
    function openApprovalLinePopup() {
        const width = 1200;
        const height = 800;
        const left = (screen.width - width) / 2;
        const top = (screen.height - height) / 2;
        
        window.open('approval_line_popup.asp', 'approvalLinePopup',
            `width=${width},height=${height},left=${left},top=${top},scrollbars=yes`);
    }

    // 결재선 데이터 설정 (팝업에서 호출)
    function setApprovalLine(data) {
        if (!data || !Array.isArray(data) || data.length < 2) {
            alert('올바른 결재선을 지정해주세요.');
            return;
        }

        approvalLineData = data;
        updateApprovalLineDisplay();
        updateHiddenFields();
    }

    // 결재선 표시 업데이트
    function updateApprovalLineDisplay() {
        const container = document.getElementById('approvalLineDisplay');
        container.innerHTML = '';
        
        // 1차 결재자(본인)는 이미 고정으로 표시되어 있으므로 2차부터 표시
        for (let i = 1; i < approvalLineData.length; i++) {
            const approver = approvalLineData[i];
            const stepDiv = document.createElement('div');
            stepDiv.className = 'approval-step';
            stepDiv.innerHTML = `
                <span class="step-label">${i + 1}차 결재</span>
                <div class="approver-info">
                    <div class="approver-name">${approver.userName}</div>
                    <div class="approver-dept">${approver.deptName}</div>
                </div>
            `;
            container.appendChild(stepDiv);
        }
    }

    // 숨겨진 입력 필드 업데이트
    function updateHiddenFields() {
        // 1차 결재자는 이미 설정되어 있음
        document.getElementById('approver_step2').value = approvalLineData[1]?.userId || '';
        document.getElementById('approver_step3').value = approvalLineData[2]?.userId || '';
    }

    // 결재선 데이터 저장 변수
    let approvalLineData = null;
    
    // 페이지 로딩 시 초기 계산
    document.addEventListener('DOMContentLoaded', function() {
        calculateAmount();
    });

    // 폼 제출 전 유효성 검사
    document.getElementById('vehicleRequestForm').addEventListener('submit', function(e) {
        if (!document.getElementById('approver_step2').value) {
            e.preventDefault();
            alert('최소 1명 이상의 결재자를 지정해주세요.');
            return false;
        }
    });
</script>

<style>
/* 결재선 관련 스타일 */
.approval-line-box {
    border: 2px solid #E9ECEF;
    border-radius: 12px;
    padding: 1.5rem;
    margin-bottom: 1.75rem;
    background-color: #fff;
    transition: all 0.2s ease;
}

.approval-line-box:hover {
    border-color: #4A90E2;
    box-shadow: 0 4px 12px rgba(74,144,226,0.1);
}

.approval-step {
    display: flex;
    align-items: center;
    margin-bottom: 1rem;
    padding: 1.25rem;
    background: #F8FAFC;
    border: 2px solid #E9ECEF;
    border-radius: 10px;
    transition: all 0.2s ease;
}

.approval-step:hover {
    background: #fff;
    border-color: #4A90E2;
    box-shadow: 0 4px 8px rgba(0,0,0,0.05);
    transform: translateY(-1px);
}

.approval-step:last-child {
    margin-bottom: 0;
}

.step-label {
    font-weight: 600;
    color: #2C3E50;
    margin-right: 1.5rem;
    min-width: 90px;
    font-size: 0.95rem;
    padding: 0.5rem 1rem;
    background: #E9ECEF;
    border-radius: 6px;
}

.approver-info {
    flex-grow: 1;
}

.approver-name {
    font-weight: 600;
    margin-bottom: 0.375rem;
    color: #2C3E50;
    font-size: 1.05rem;
}

.approver-dept {
    font-size: 0.9rem;
    color: #5D6D7E;
}
</style>

<!--#include file="../includes/footer.asp"--> 