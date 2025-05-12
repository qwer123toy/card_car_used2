<!--#include file="../includes/connection.asp"-->
<!--#include file="../includes/functions.asp"-->
<%
' 로그인 체크
If Not IsAuthenticated() Then
    RedirectTo("../index.asp")
End If

' 최신 유류비 단가 조회
Dim fuelRateSQL, fuelRateRS, fuelRate
fuelRateSQL = "SELECT TOP 1 rate FROM FuelRate ORDER BY date DESC"
Set fuelRateRS = dbConn.Execute(fuelRateSQL)

If Not fuelRateRS.EOF Then
    fuelRate = fuelRateRS("rate")
Else
    fuelRate = 0
End If
fuelRateRS.Close

' 신청서 등록 처리
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    Dim requestDate, purpose, startLocation, destination, distance, tollFee, parkingFee, totalAmount, errorMsg, successMsg, insertSQL
    
    requestDate = PreventSQLInjection(Request.Form("request_date"))
    purpose = PreventSQLInjection(Request.Form("purpose"))
    startLocation = PreventSQLInjection(Request.Form("start_location"))
    destination = PreventSQLInjection(Request.Form("destination"))
    distance = CDbl(Replace(PreventSQLInjection(Request.Form("distance")), ",", ""))
    tollFee = CDbl(Replace(PreventSQLInjection(Request.Form("toll_fee")), ",", ""))
    parkingFee = CDbl(Replace(PreventSQLInjection(Request.Form("parking_fee")), ",", ""))
    
    ' 총 금액 계산
    totalAmount = (distance * fuelRate) + tollFee + parkingFee
    
    ' 입력값 검증
    If requestDate = "" Or purpose = "" Or startLocation = "" Or destination = "" Then
        errorMsg = "필수 항목을 모두 입력해주세요."
    ElseIf Not IsNumeric(distance) Or Not IsNumeric(tollFee) Or Not IsNumeric(parkingFee) Then
        errorMsg = "거리와 비용은 숫자만 입력 가능합니다."
    Else
        ' 차량 이용 신청서 등록
        insertSQL = "INSERT INTO VehicleRequests (user_id, request_date, purpose, start_location, destination, " & _
                   "distance, toll_fee, parking_fee, total_amount, approval_status, is_deleted) VALUES ('" & _
                   Session("user_id") & "', '" & requestDate & "', '" & purpose & "', '" & startLocation & "', '" & _
                   destination & "', " & distance & ", " & tollFee & ", " & parkingFee & ", " & totalAmount & ", '작성중', 0)"
        
        On Error Resume Next
        dbConn.Execute insertSQL
        
        If Err.Number <> 0 Then
            errorMsg = "차량 이용 신청서 등록 중 오류가 발생했습니다: " & Err.Description
        Else
            ' 방금 등록한 신청서 ID 조회
            Dim newRequestIdSQL, newRequestIdRS, newRequestId
            newRequestIdSQL = "SELECT TOP 1 request_id FROM VehicleRequests WHERE user_id = '" & Session("user_id") & _
                             "' ORDER BY request_id DESC"
            Set newRequestIdRS = dbConn.Execute(newRequestIdSQL)
            
            If Not newRequestIdRS.EOF Then
                newRequestId = newRequestIdRS("request_id")
                newRequestIdRS.Close
                
                successMsg = "차량 이용 신청서가 성공적으로 등록되었습니다."
                
                ' 활동 로그 기록
                LogActivity Session("user_id"), "차량이용신청", "개인차량 이용 신청서 등록 (ID: " & newRequestId & ")"
                
                ' 신청서 상세 페이지로 리디렉션
                RedirectTo("vehicle_request_view.asp?id=" & newRequestId)
            Else
                errorMsg = "신청서 등록 후 ID를 찾는 데 실패했습니다."
            End If
        End If
        On Error GoTo 0
    End If
End If
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
            <form id="vehicleRequestForm" method="post" action="vehicle_request_add.asp" onsubmit="return validateForm('vehicleRequestForm', vehicleRequestRules)">
                <div class="form-group">
                    <label class="shadcn-input-label" for="request_date">이용일자</label>
                    <input class="shadcn-input" type="date" id="request_date" name="request_date" value="<%= FormatDate(Date()) %>">
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
                    <input class="shadcn-input" type="text" id="distance" name="distance" placeholder="운행거리를 입력하세요" onkeyup="calculateAmount()">
                </div>
                
                <div class="form-group">
                    <label class="shadcn-input-label" for="fuel_rate">현재 유류비 단가</label>
                    <input class="shadcn-input" type="text" id="fuel_rate" name="fuel_rate" value="<%= fuelRate %>" readonly>
                </div>
                
                <div class="form-group">
                    <label class="shadcn-input-label" for="toll_fee">통행료</label>
                    <input class="shadcn-input" type="text" id="toll_fee" name="toll_fee" placeholder="통행료를 입력하세요" value="0" onkeyup="formatCurrency(this); calculateAmount()">
                </div>
                
                <div class="form-group">
                    <label class="shadcn-input-label" for="parking_fee">주차비</label>
                    <input class="shadcn-input" type="text" id="parking_fee" name="parking_fee" placeholder="주차비를 입력하세요" value="0" onkeyup="formatCurrency(this); calculateAmount()">
                </div>
                
                <div class="form-group">
                    <label class="shadcn-input-label" for="total_amount">총 예상 금액</label>
                    <input class="shadcn-input" type="text" id="total_amount" name="total_amount" readonly>
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
        request_date: {
            required: true,
            message: '이용일자를 입력해주세요.'
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
    
    // 총 금액 계산
    function calculateAmount() {
        const distance = parseFloat(document.getElementById('distance').value.replace(/,/g, '')) || 0;
        const fuelRate = parseFloat(document.getElementById('fuel_rate').value) || 0;
        const tollFee = parseFloat(document.getElementById('toll_fee').value.replace(/,/g, '')) || 0;
        const parkingFee = parseFloat(document.getElementById('parking_fee').value.replace(/,/g, '')) || 0;
        
        const fuelAmount = distance * fuelRate;
        const totalAmount = fuelAmount + tollFee + parkingFee;
        
        document.getElementById('total_amount').value = totalAmount.toLocaleString('ko-KR');
    }
    
    // 페이지 로딩 시 초기 계산
    document.addEventListener('DOMContentLoaded', function() {
        calculateAmount();
    });
</script>

<!--#include file="../includes/footer.asp"--> 