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

' URL 파라미터에서 신청서 ID 추출
Dim requestId, errorMsg, successMsg
requestId = PreventSQLInjection(Request.QueryString("id"))

If requestId = "" Then
    errorMsg = "잘못된 접근입니다. 신청서 ID가 필요합니다."
    Response.Redirect("vehicle_request.asp")
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

' 신청서 정보 조회
Dim cmd, rs
Set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = db
cmd.CommandText = "SELECT * FROM VehicleRequests WHERE request_id = ? AND is_deleted = 0"
cmd.Parameters.Append cmd.CreateParameter("@request_id", 3, 1, , CLng(requestId))

Set rs = cmd.Execute()

If Err.Number <> 0 Or rs.EOF Then
    errorMsg = "요청하신 신청서를 찾을 수 없습니다."
    RedirectTo("vehicle_request.asp")
ElseIf rs("user_id") <> Session("user_id") Then
    errorMsg = "본인이 작성한 신청서만 수정할 수 있습니다."
    RedirectTo("vehicle_request.asp")
ElseIf rs("approval_status") <> "작성중" Then
    errorMsg = "작성중 상태의 신청서만 수정할 수 있습니다."
    RedirectTo("vehicle_request.asp")
End If

' 신청서 정보 저장
Dim startDate, endDate, purpose, startLocation, destination, distance
startDate = rs("start_date")
endDate = rs("end_date")
purpose = rs("purpose")
startLocation = rs("start_location")
destination = rs("destination")
distance = rs("distance")

' 폼 제출 처리
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    startDate = PreventSQLInjection(Request.Form("start_date"))
    endDate = PreventSQLInjection(Request.Form("end_date"))
    purpose = PreventSQLInjection(Request.Form("purpose"))
    startLocation = PreventSQLInjection(Request.Form("start_location"))
    destination = PreventSQLInjection(Request.Form("destination"))
    
    ' 종료일자가 비어있으면 시작일자와 동일하게 설정
    If endDate = "" Then
        endDate = startDate
    End If
    
    ' 숫자 값 안전하게 변환
    distance = 0
    
    If IsNumeric(Replace(Request.Form("distance"), ",", "")) Then
        distance = CDbl(Replace(Request.Form("distance"), ",", ""))
    End If
    
    ' 입력값 검증
    If startDate = "" Or endDate = "" Or purpose = "" Or startLocation = "" Or destination = "" Then
        errorMsg = "필수 항목을 모두 입력해주세요."
    ElseIf distance <= 0 Then
        errorMsg = "운행거리는 양수여야 합니다."
    Else
        ' 파라미터화된 쿼리 사용하여 신청서 정보 업데이트
        Dim updateCmd
        Set updateCmd = Server.CreateObject("ADODB.Command")
        updateCmd.ActiveConnection = db
        updateCmd.CommandText = "UPDATE VehicleRequests SET start_date = ?, end_date = ?, purpose = ?, " & _
                              "start_location = ?, destination = ?, distance = ? " & _
                              "WHERE request_id = ?"
        
        ' 파라미터 추가
        updateCmd.Parameters.Append updateCmd.CreateParameter("@start_date", 7, 1, , startDate)
        updateCmd.Parameters.Append updateCmd.CreateParameter("@end_date", 7, 1, , endDate)
        updateCmd.Parameters.Append updateCmd.CreateParameter("@purpose", 200, 1, 100, purpose)
        updateCmd.Parameters.Append updateCmd.CreateParameter("@start_location", 200, 1, 100, startLocation)
        updateCmd.Parameters.Append updateCmd.CreateParameter("@destination", 200, 1, 100, destination)
        updateCmd.Parameters.Append updateCmd.CreateParameter("@distance", 6, 1, , distance)
        updateCmd.Parameters.Append updateCmd.CreateParameter("@request_id", 3, 1, , CLng(requestId))
        
        ' 명령 실행
        On Error Resume Next
        updateCmd.Execute
        
        If Err.Number <> 0 Then
            errorMsg = "차량 이용 신청서 수정 중 오류가 발생했습니다: " & Err.Description
        Else
            successMsg = "차량 이용 신청서가 성공적으로 수정되었습니다."
            
            ' 활동 로그 기록
            Dim totalAmount
            totalAmount = distance * fuelRate
            LogActivity Session("user_id"), "차량이용신청수정", "개인차량 이용 신청서 수정 (ID: " & requestId & ", 거리: " & distance & "km, 총액: " & FormatNumber(totalAmount) & "원)"
            
            ' 상세 페이지로 리디렉션
            RedirectTo("vehicle_request_view.asp?id=" & requestId)
        End If
        On Error GoTo 0
    End If
End If

On Error GoTo 0
%>
<!--#include file="../includes/header.asp"-->

<div class="vehicle-request-edit-container">
    <div class="shadcn-card" style="max-width: 700px; margin: 30px auto;">
        <div class="shadcn-card-header">
            <h2 class="shadcn-card-title">개인차량 이용 신청서 수정</h2>
            <p class="shadcn-card-description">개인차량 이용 신청서 내용을 수정합니다.</p>
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
            <form id="vehicleRequestForm" method="post" action="vehicle_request_edit.asp?id=<%= requestId %>" onsubmit="prepareFormSubmission(); return validateForm('vehicleRequestForm', vehicleRequestRules)">
                <div class="form-row" style="display: flex; gap: 10px;">
                    <div class="form-group" style="flex: 1;">
                        <label class="shadcn-input-label" for="start_date">시작일자</label>
                        <input class="shadcn-input" type="date" id="start_date" name="start_date" value="<%= FormatDate(startDate) %>">
                    </div>
                    
                    <div class="form-group" style="flex: 1;">
                        <label class="shadcn-input-label" for="end_date">종료일자</label>
                        <input class="shadcn-input" type="date" id="end_date" name="end_date" value="<%= FormatDate(endDate) %>">
                    </div>
                </div>
                
                <div class="form-group">
                    <label class="shadcn-input-label" for="purpose">업무 목적</label>
                    <input class="shadcn-input" type="text" id="purpose" name="purpose" placeholder="업무 목적을 입력하세요" value="<%= purpose %>">
                </div>
                
                <div class="form-group">
                    <label class="shadcn-input-label" for="start_location">출발지</label>
                    <input class="shadcn-input" type="text" id="start_location" name="start_location" placeholder="출발지를 입력하세요" value="<%= startLocation %>">
                </div>
                
                <div class="form-group">
                    <label class="shadcn-input-label" for="destination">목적지</label>
                    <input class="shadcn-input" type="text" id="destination" name="destination" placeholder="목적지를 입력하세요" value="<%= destination %>">
                </div>
                
                <div class="form-group">
                    <label class="shadcn-input-label" for="distance">운행거리 (km)</label>
                    <input class="shadcn-input" type="text" id="distance" name="distance" placeholder="운행거리를 입력하세요" value="<%= distance %>" onkeyup="cleanNumberInput(this); calculateAmount()">
                </div>
                
                <div class="form-group">
                    <label class="shadcn-input-label" for="fuel_rate">현재 유류비 단가</label>
                    <input class="shadcn-input" type="text" id="fuel_rate" name="fuel_rate" value="<%= fuelRate %>" readonly>
                </div>
                
                <div class="form-group">
                    <label class="shadcn-input-label" for="total_amount_display">총 예상 금액</label>
                    <input class="shadcn-input" type="text" id="total_amount_display" readonly>
                </div>
                
                <div class="shadcn-card-footer" style="margin-top: 1.5rem;">
                    <button type="submit" class="shadcn-btn shadcn-btn-primary">수정하기</button>
                    <a href="vehicle_request_view.asp?id=<%= requestId %>" class="shadcn-btn shadcn-btn-outline">취소</a>
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
        }
    };
    
    // 폼 제출 전 숫자 필드의 쉼표 제거
    function prepareFormSubmission() {
        // 숫자 입력 필드의 쉼표 제거
        const numericFields = ['distance'];
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
        
        // 쉼표 제거 후 숫자로 변환
        const distance = parseFloat(distanceInput.replace(/,/g, '')) || 0;
        const fuelRate = parseFloat(fuelRateInput.replace(/,/g, '')) || 0;
        
        const totalAmount = distance * fuelRate;
        
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
    
    // 페이지 로딩 시 초기 계산
    document.addEventListener('DOMContentLoaded', function() {
        calculateAmount();
    });
</script>

<!--#include file="../includes/footer.asp"--> 