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

' 데이터베이스 연결 확인
If Not dbConnected Then
    Response.Write("<script>alert('데이터베이스 연결에 실패했습니다.');window.location.href='card_usage.asp';</script>")
    Response.End
End If

' URL 파라미터에서 ID 가져오기
Dim usageId, errorMsg, usageIdOrig
usageIdOrig = Request.QueryString("id")
usageId = PreventSQLInjection(usageIdOrig)

' 디버깅용 정보 기록
Response.Write("<!-- 원본 id 파라미터: " & usageIdOrig & " -->")
Response.Write("<!-- 처리된 id 파라미터: " & usageId & " -->")

If usageId = "" Then
    errorMsg = "잘못된 접근입니다. 카드 사용 내역 ID가 필요합니다."
    Session("error_msg") = errorMsg
    Response.Write("<meta http-equiv='refresh' content='0;url=card_usage.asp'>")
    Response.End
End If

' dbSchema가 설정되지 않은 경우를 대비해 기본값 설정
If Not(IsObject(dbSchema)) And (TypeName(dbSchema) <> "String" Or Len(dbSchema) = 0) Then
    dbSchema = "dbo"
End If

' 카드 사용 내역 조회
Dim SQL, rs
SQL = "SELECT usage_id, user_id, card_id, department_id, expense_category_id, usage_date, " & _
      "store_name, amount, purpose, linked_table, linked_id, receipt_file, created_at, approval_status " & _
      "FROM " & dbSchema & ".CardUsage " & _
      "WHERE usage_id = " & usageId

' SQL 쿼리 실행 전에 확인
Response.Write("<!-- 디버그: SQL 쿼리: " & SQL & " -->")

' 테이블 이름이 다를 경우를 위한 대체 쿼리
On Error Resume Next
Set rs = db.Execute(SQL)

' 첫 번째 쿼리가 실패하면 다른 테이블 이름으로 시도합니다
If Err.Number <> 0 Then
    Err.Clear
    SQL = "SELECT * FROM " & dbSchema & ".CardHistory WHERE usage_id = " & usageId
    Response.Write("<!-- 대체 쿼리 시도: " & SQL & " -->")
    Set rs = db.Execute(SQL)
End If

' 그래도 실패하면 Card_Usage 테이블로 시도합니다
If Err.Number <> 0 Then
    Err.Clear
    SQL = "SELECT * FROM " & dbSchema & ".Card_Usage WHERE usage_id = " & usageId
    Response.Write("<!-- 두 번째 대체 쿼리 시도: " & SQL & " -->")
    Set rs = db.Execute(SQL)
End If

' 오류 디버깅을 위한 코드 추가
Response.Write("<!-- SQL 실행 후 Err.Number: " & Err.Number & ", Err.Description: " & Err.Description & " -->")

' 오류 상세 확인
If Err.Number <> 0 Then
    errorMsg = "카드 사용 내역을 조회하는 중 오류가 발생했습니다."
    Session("error_msg") = errorMsg
    Response.Write("<meta http-equiv='refresh' content='0;url=card_usage.asp'>")
    Response.End
End If

' 데이터 확인
If rs Is Nothing Then
    errorMsg = "데이터베이스 응답이 없습니다."
    Session("error_msg") = errorMsg
    Response.Write("<meta http-equiv='refresh' content='0;url=card_usage.asp'>")
    Response.End
End If

' 부서명 가져오기
Function GetDepartmentName(deptId)
    Dim deptName, deptSQL, deptRS
    
    If deptId <> "" And IsNumeric(deptId) Then
        deptSQL = "SELECT name FROM " & dbSchema & ".Department WHERE department_id = " & deptId
        
        On Error Resume Next
        Set deptRS = db.Execute(deptSQL)
        
        If Err.Number = 0 And Not deptRS.EOF Then
            deptName = deptRS("name")
        Else
            ' DB에서 정보를 찾지 못한 경우 ID 값 그대로 표시
            deptName = deptId
        End If
        
        If Not deptRS Is Nothing Then
            If deptRS.State = 1 Then
                deptRS.Close
            End If
            Set deptRS = Nothing
        End If
        On Error GoTo 0
    Else
        deptName = "-"
    End If
    
    GetDepartmentName = deptName
End Function

' 카드 이름 가져오기
Function GetCardName(cardId)
    Dim cardName, cardSQL, cardRS
    
    If cardId <> "" And IsNumeric(cardId) Then
        cardSQL = "SELECT account_name FROM " & dbSchema & ".CardAccount WHERE card_id = " & cardId
        
        On Error Resume Next
        Set cardRS = db.Execute(cardSQL)
        
        If Err.Number = 0 And Not cardRS.EOF Then
            cardName = cardRS("account_name")
        Else
            ' DB에서 정보를 찾지 못한 경우 ID 값 그대로 표시
            cardName = cardId
        End If
        
        If Not cardRS Is Nothing Then
            If cardRS.State = 1 Then
                cardRS.Close
            End If
            Set cardRS = Nothing
        End If
        On Error GoTo 0
    Else
        cardName = "-"
    End If
    
    GetCardName = cardName
End Function

' 계정과목명 가져오기
Function GetExpenseCategoryName(categoryId)
    Dim categoryName, catSQL, catRS
    
    If categoryId <> "" And IsNumeric(categoryId) Then
        catSQL = "SELECT type_name FROM " & dbSchema & ".CardAccountTypes WHERE account_type_id = " & categoryId
        
        On Error Resume Next
        Set catRS = db.Execute(catSQL)
        
        If Err.Number = 0 And Not catRS.EOF Then
            categoryName = catRS("type_name")
        Else
            ' DB에서 정보를 찾지 못한 경우 ID 값 그대로 표시
            categoryName = categoryId
        End If
        
        If Not catRS Is Nothing Then
            If catRS.State = 1 Then
                catRS.Close
            End If
            Set catRS = Nothing
        End If
        On Error GoTo 0
    Else
        categoryName = "-"
    End If
    
    GetExpenseCategoryName = categoryName
End Function
%>
<!--#include file="../includes/header.asp"-->

<div class="card-usage-view-container">
    <div class="shadcn-card" style="max-width: 800px; margin: 30px auto;">
        <div class="shadcn-card-header">
            <div style="display: flex; justify-content: space-between; align-items: center;">
                <h2 class="shadcn-card-title">카드 사용 내역 상세</h2>
                <div>
                    <div class="text-center mt-4">
                        <% If rs("approval_status") = "대기" Or rs("approval_status") = "반려" Then %>
                            <a href="card_usage_edit.asp?id=<%= rs("usage_id") %>" class="btn btn-primary me-2">
                                <i class="fas fa-edit me-1"></i> 수정
                            </a>
                        <% End If %>
                        <% If rs("approval_status") <> "완료" Then %>
                            <button onclick="confirmDelete('<%= rs("usage_id") %>')" class="btn btn-destructive me-2">삭제</button>
                        <% End If %>
                        <a href="card_usage.asp" class="btn btn-outline">목록으로</a>
                    </div>
                </div>
            </div>
            <p class="shadcn-card-description">카드 사용 내역의 상세 정보를 확인합니다.</p>
        </div>
        
        <div class="shadcn-card-content">
            <div class="shadcn-details">
                <div class="detail-section">
                    <h3 class="detail-heading">기본 정보</h3>
                    <div class="detail-grid">
                        <div class="detail-item">
                            <span class="detail-label">사용자 ID</span>
                            <span class="detail-value"><%= rs("user_id") %></span>
                        </div>
                        <div class="detail-item">
                            <span class="detail-label">부서</span>
                            <span class="detail-value"><%= GetDepartmentName(rs("department_id")) %></span>
                        </div>
                        <div class="detail-item">
                            <span class="detail-label">사용 일자</span>
                            <span class="detail-value"><%= FormatDate(rs("usage_date")) %></span>
                        </div>
                        <div class="detail-item">
                            <span class="detail-label">사용 금액</span>
                            <span class="detail-value"><%= FormatNumber(rs("amount")) %>원</span>
                        </div>
                    </div>
                </div>
                
                <div class="detail-section">
                    <h3 class="detail-heading">카드 정보</h3>
                    <div class="detail-grid">
                        <div class="detail-item">
                            <span class="detail-label">카드</span>
                            <span class="detail-value"><%= GetCardName(rs("card_id")) %></span>
                        </div>
                        <div class="detail-item">
                            <span class="detail-label">계정 과목</span>
                            <span class="detail-value"><%= GetExpenseCategoryName(rs("expense_category_id")) %></span>
                        </div>
                        <div class="detail-item">
                            <span class="detail-label">사용처</span>
                            <span class="detail-value"><%= rs("store_name") %></span>
                        </div>
                        <div class="detail-item">
                            <span class="detail-label">승인 상태</span>
                            <span class="detail-value">
                                <span class="status-badge status-<%= LCase(rs("approval_status")) %>">
                                    <%= rs("approval_status") %>
                                </span>
                            </span>
                        </div>
                    </div>
                </div>
                
                <div class="detail-section">
                    <h3 class="detail-heading">사용 목적</h3>
                    <div class="detail-item full-width">
                        <div class="detail-value">
                            <%= rs("purpose") %>
                        </div>
                    </div>
                </div>
                
                <div class="detail-section">
                    <h3 class="detail-heading">첨부 파일</h3>
                    <div class="detail-grid">
                        <div class="detail-item">
                            <span class="detail-label">영수증 파일</span>
                            <span class="detail-value">
                                <% If IsNull(rs("receipt_file")) Or rs("receipt_file") = "" Then %>
                                    -
                                <% Else %>
                                    <%= rs("receipt_file") %>
                                <% End If %>
                            </span>
                        </div>
                        <div class="detail-item">
                            <span class="detail-label">연결 테이블</span>
                            <span class="detail-value">
                                <% If IsNull(rs("linked_table")) Or rs("linked_table") = "" Then %>
                                    -
                                <% Else %>
                                    <%= rs("linked_table") %>
                                <% End If %>
                            </span>
                        </div>
                        <div class="detail-item">
                            <span class="detail-label">연결 ID</span>
                            <span class="detail-value">
                                <% If IsNull(rs("linked_id")) Or rs("linked_id") = 0 Then %>
                                    -
                                <% Else %>
                                    <%= rs("linked_id") %>
                                <% End If %>
                            </span>
                        </div>
                    </div>
                </div>
                
                <div class="detail-section">
                    <h3 class="detail-heading">등록 정보</h3>
                    <div class="detail-grid">
                        <div class="detail-item">
                            <span class="detail-label">등록일</span>
                            <span class="detail-value"><%= FormatDateTime(rs("created_at"), 2) %></span>
                        </div>
                        <div class="detail-item">
                            <span class="detail-label">수정일</span>
                            <span class="detail-value">-</span>
                        </div>
                    </div>
                </div>
            </div>
        </div>
        
        <div class="shadcn-card-footer">
            <div style="display: flex; justify-content: space-between; width: 100%;">
                <a href="card_usage.asp" class="shadcn-btn shadcn-btn-outline">목록으로 돌아가기</a>
            </div>
        </div>
    </div>
</div>

<style>
    .detail-section {
        margin-bottom: 24px;
        border-bottom: 1px solid #eaeaea;
        padding-bottom: 16px;
    }
    
    .detail-section:last-child {
        border-bottom: none;
    }
    
    .detail-heading {
        font-size: 16px;
        font-weight: 600;
        margin-bottom: 12px;
        color: #333;
    }
    
    .detail-grid {
        display: grid;
        grid-template-columns: repeat(auto-fill, minmax(250px, 1fr));
        gap: 16px;
    }
    
    .detail-item {
        display: flex;
        flex-direction: column;
    }
    
    .detail-item.full-width {
        grid-column: 1 / -1;
    }
    
    .detail-label {
        font-size: 14px;
        color: #666;
        margin-bottom: 4px;
    }
    
    .detail-value {
        font-size: 16px;
        color: #333;
    }
    
    .detail-textarea {
        padding: 12px;
        background-color: #f9f9f9;
        border-radius: 4px;
        min-height: 80px;
        white-space: pre-wrap;
        line-height: 1.5;
        font-size: 15px;
    }
    
    .purpose-box {
        background-color: #f0f7ff;
        border-left: 4px solid #3b82f6;
        padding: 15px;
        min-height: 100px;
        margin-top: 5px;
    }
    
    .status-badge {
        display: inline-block;
        padding: 4px 8px;
        border-radius: 4px;
        font-size: 12px;
        font-weight: 500;
    }
    
    .status-승인대기 {
        background-color: #fff8e6;
        color: #f59f00;
    }
    
    .status-승인완료 {
        background-color: #e6f7ee;
        color: #12b886;
    }
    
    .status-반려 {
        background-color: #ffe9e9;
        color: #fa5252;
    }
    
    .status-처리중 {
        background-color: #e7f5ff;
        color: #339af0;
    }
    
    .shadcn-btn-destructive {
        background-color: #ef4444;
        color: white;
    }
    
    .shadcn-btn-destructive:hover {
        background-color: #dc2626;
    }
</style>

<script>
    function confirmDelete(id) {
        if (confirm('정말로 이 카드 사용 내역을 삭제하시겠습니까? 이 작업은 되돌릴 수 없습니다.')) {
            window.location.href = 'card_usage_delete.asp?id=' + id;
        }
    }
</script>

<%
' 사용한 Recordset 닫기
If IsObject(rs) Then
    If rs.State = 1 Then ' adStateOpen
        rs.Close
    End If
    Set rs = Nothing
End If
%>

<!--#include file="../includes/footer.asp"--> 