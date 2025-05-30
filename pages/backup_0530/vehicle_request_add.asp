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
    Dim startDate, endDate, purpose, startLocation, destination, distance, tollFee, parkingFee, totalAmount, errorMsg, successMsg, title, fuelCost, total_cost
    Dim approver1, approver2, approver3
    
    startDate = PreventSQLInjection(Request.Form("start_date"))
    endDate = PreventSQLInjection(Request.Form("end_date"))
    purpose = PreventSQLInjection(Request.Form("purpose"))
    startLocation = PreventSQLInjection(Request.Form("start_location"))
    destination = PreventSQLInjection(Request.Form("destination"))
    distance = PreventSQLInjection(Request.Form("distance"))
    fuelCost = PreventSQLInjection(Request.Form("fuel_cost"))
    tollFee = PreventSQLInjection(Request.Form("toll_fee"))
    parkingFee = PreventSQLInjection(Request.Form("parking_fee"))
    total_cost = PreventSQLInjection(Request.Form("total_cost"))
    title = PreventSQLInjection(Request.Form("title"))
    approver1 = PreventSQLInjection(Request.Form("approver_step1"))
    approver2 = PreventSQLInjection(Request.Form("approver_step2"))
    approver3 = PreventSQLInjection(Request.Form("approver_step3"))
    Dim approver4, approver5
    approver4 = PreventSQLInjection(Request.Form("approver_step4"))
    approver5 = PreventSQLInjection(Request.Form("approver_step5"))
    
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
    total_cost = (distance * fuelRate) + tollFee + parkingFee
    
    ' 입력값 검증
    If startDate = "" Or endDate = "" Or purpose = "" Or startLocation = "" Or destination = "" Or title = "" Then
        errorMsg = "필수 항목을 모두 입력해주세요."
    ElseIf distance <= 0 Then
        errorMsg = "운행거리는 양수여야 합니다."
    Else
        ' 트랜잭션 시작
        db.BeginTrans

        ' 파라미터화된 쿼리 사용 - start_date와 end_date 필드 사용
        Dim cmd
        Set cmd = Server.CreateObject("ADODB.Command")
        cmd.ActiveConnection = db
        
        ' 유류비 계산    
        fuelCost = distance * fuelRate
        
        ' SQL 파라미터 13개로 수정
cmd.CommandText = "INSERT INTO VehicleRequests (user_id, start_date, end_date, purpose, start_location, destination, " & _
                  "distance, fuel_cost, toll_fee, parking_fee, total_cost, title, approval_status, is_deleted) " & _
                  "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)"
        
        ' 파라미터 추가
        cmd.Parameters.Append cmd.CreateParameter("@user_id", 200, 1, 30, Session("user_id"))
        cmd.Parameters.Append cmd.CreateParameter("@start_date", 7, 1, , startDate)
        cmd.Parameters.Append cmd.CreateParameter("@end_date", 7, 1, , endDate)
        cmd.Parameters.Append cmd.CreateParameter("@purpose", 200, 1, 100, purpose)
        cmd.Parameters.Append cmd.CreateParameter("@start_location", 200, 1, 100, startLocation)
        cmd.Parameters.Append cmd.CreateParameter("@destination", 200, 1, 100, destination)
        cmd.Parameters.Append cmd.CreateParameter("@distance", 6, 1, , distance)
        cmd.Parameters.Append cmd.CreateParameter("@fuel_cost", 6, 1, , fuelCost)
        cmd.Parameters.Append cmd.CreateParameter("@toll_fee", 6, 1, , tollFee)
        cmd.Parameters.Append cmd.CreateParameter("@parking_fee", 6, 1, , parkingFee)
        cmd.Parameters.Append cmd.CreateParameter("@total_cost", 6, 1, , total_cost)
        cmd.Parameters.Append cmd.CreateParameter("@title", 200, 1, 200, title)
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
                    LogActivity Session("user_id"), "차량이용신청", "개인차량 이용 신청서 등록 (ID: " & newRequestId & ", 거리: " & distance & "km, 총액: " & FormatNumber(total_cost) & "원)"
                    
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

                <!-- 신청서 정보 테이블 -->
                <div class="approval-line-table-container" style="margin-top: 1rem;">
                    <table class="approval-line-table">
                        <tbody>
                            <tr>
                                <td class="approval-cell" style="background: #F8FAFC; font-weight: 600; width: 20%;">제목</td>
                                <td class="approval-cell" colspan="4" style="text-align: left; padding: 1rem;">
                                    <input class="shadcn-input" type="text" id="title" name="title" placeholder="결재 제목을 입력하세요" style="border: 1px solid #E9ECEF; width: 100%;">
                                </td>
                            </tr>
                            <tr>
                                <td class="approval-cell" style="background: #F8FAFC; font-weight: 600;">시작일자</td>
                                <td class="approval-cell" style="text-align: left; padding: 1rem;">
                                    <input class="shadcn-input" type="date" id="start_date" name="start_date" value="<%= FormatDate(Date()) %>" style="border: 1px solid #E9ECEF; width: 100%;">
                                </td>
                                <td class="approval-cell" style="background: #F8FAFC; font-weight: 600;">종료일자</td>
                                <td class="approval-cell" colspan="2" style="text-align: left; padding: 1rem;">
                                    <input class="shadcn-input" type="date" id="end_date" name="end_date" value="<%= FormatDate(Date()) %>" style="border: 1px solid #E9ECEF; width: 100%;">
                                </td>
                            </tr>
                            <tr>
                                <td class="approval-cell" style="background: #F8FAFC; font-weight: 600;">업무 목적</td>
                                <td class="approval-cell" colspan="4" style="text-align: left; padding: 1rem;">
                                    <input class="shadcn-input" type="text" id="purpose" name="purpose" placeholder="업무 목적을 입력하세요" style="border: 1px solid #E9ECEF; width: 100%;">
                                </td>
                            </tr>
                            <tr>
                                <td class="approval-cell" style="background: #F8FAFC; font-weight: 600;">출발지</td>
                                <td class="approval-cell" style="text-align: left; padding: 1rem;">
                                    <input class="shadcn-input" type="text" id="start_location" name="start_location" placeholder="출발지를 입력하세요" style="border: 1px solid #E9ECEF; width: 100%;">
                                </td>
                                <td class="approval-cell" style="background: #F8FAFC; font-weight: 600;">목적지</td>
                                <td class="approval-cell" colspan="2" style="text-align: left; padding: 1rem;">
                                    <input class="shadcn-input" type="text" id="destination" name="destination" placeholder="목적지를 입력하세요" style="border: 1px solid #E9ECEF; width: 100%;">
                                </td>
                            </tr>
                            <tr>
                                <td class="approval-cell" style="background: #F8FAFC; font-weight: 600;">운행거리 (km)</td>
                                <td class="approval-cell" style="text-align: left; padding: 1rem;">
                                    <input class="shadcn-input" type="text" id="distance" name="distance" placeholder="운행거리를 입력하세요" onkeyup="cleanNumberInput(this); calculateAmount()" style="border: 1px solid #E9ECEF; width: 100%;">
                                </td>
                                <td class="approval-cell" style="background: #F8FAFC; font-weight: 600;">유류비 단가</td>
                                <td class="approval-cell" colspan="2" style="text-align: left; padding: 1rem;">
                                    <input class="shadcn-input" type="text" id="fuel_cost" name="fuel_cost" value="<%= fuelRate %>" readonly style="border: 1px solid #E9ECEF; width: 100%; background: #F8FAFC;">
                                </td>
                            </tr>
                            <tr>
                                <td class="approval-cell" style="background: #F8FAFC; font-weight: 600;">통행료</td>
                                <td class="approval-cell" style="text-align: left; padding: 1rem;">
                                    <input class="shadcn-input" type="text" id="toll_fee" name="toll_fee" placeholder="통행료를 입력하세요" value="0" onkeyup="cleanNumberInput(this); calculateAmount()" style="border: 1px solid #E9ECEF; width: 100%;">
                                </td>
                                <td class="approval-cell" style="background: #F8FAFC; font-weight: 600;">주차비</td>
                                <td class="approval-cell" colspan="2" style="text-align: left; padding: 1rem;">
                                    <input class="shadcn-input" type="text" id="parking_fee" name="parking_fee" placeholder="주차비를 입력하세요" value="0" onkeyup="cleanNumberInput(this); calculateAmount()" style="border: 1px solid #E9ECEF; width: 100%;">
                                </td>
                            </tr>
                            <tr>
                                <td class="approval-cell" style="background: #F8FAFC; font-weight: 600;">총 예상 금액</td>
                                <td class="approval-cell" colspan="4" style="text-align: left; padding: 1rem;">
                                    <input class="shadcn-input" type="text" id="total_cost" readonly style="border: 1px solid #E9ECEF; width: 100%; background: #F8FAFC; font-weight: 600; color: #2563EB;">
                                </td>
                            </tr>
                        </tbody>
                    </table>
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
        title: {
            required: true,
            message: '제목을 입력해주세요.'
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

    };
    
    // 폼 제출 전 숫자 필드의 쉼표 제거
    function prepareFormSubmission() {
        // 숫자 입력 필드의 쉼표 제거
        const numericFields = ['distance', 'toll_fee', 'parking_fee', 'total_cost'];
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
        const tollFeeInput = document.getElementById('toll_fee').value || '0';
        const parkingFeeInput = document.getElementById('parking_fee').value || '0';
        const fuelRateInput = document.getElementById('fuel_cost').value || '0';
        
        // 쉼표 제거 후 숫자로 변환
        const distance = parseFloat(distanceInput.replace(/,/g, '')) || 0;
        const tollFee = parseFloat(tollFeeInput.replace(/,/g, '')) || 0;
        const parkingFee = parseFloat(parkingFeeInput.replace(/,/g, '')) || 0;
        const fuelRate = parseFloat(fuelRateInput.replace(/,/g, '')) || 0;
        
        const totalAmount = (distance * fuelRate) + tollFee + parkingFee;
        
        // 화면에는 포맷된 금액 표시
        document.getElementById('total_cost').value = totalAmount.toLocaleString('ko-KR');
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

    // 결재선 데이터 저장 변수
    let approvalLineData = null;
    
    // 페이지 로딩 시 초기 계산
    document.addEventListener('DOMContentLoaded', function() {
        // 결재선 데이터 초기화 (다른 사용자 데이터 방지)
        approvalLineData = null;
        sessionStorage.removeItem('currentApprovalLine');
        
        calculateAmount();
    });


</script>

<style>
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
</style>

<!--#include file="../includes/footer.asp"--> 