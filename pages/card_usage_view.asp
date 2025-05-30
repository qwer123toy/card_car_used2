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

' 부서명 가져오기
Function GetDepartmentName(deptId)
    If IsNull(deptId) Or deptId = "" Then
        GetDepartmentName = "-"
        Exit Function
    End If
    
    Dim deptName, deptSQL, deptRS
    deptSQL = "SELECT name FROM " & dbSchema & ".Department WHERE department_id = " & deptId
    
    On Error Resume Next
    Set deptRS = db.Execute(deptSQL)
    
    If Err.Number = 0 And Not deptRS.EOF Then
        deptName = deptRS("name")
    Else
        deptName = deptId
    End If
    
    If Not deptRS Is Nothing Then
        If deptRS.State = 1 Then
            deptRS.Close
        End If
        Set deptRS = Nothing
    End If
    
    GetDepartmentName = deptName
End Function

' 카드 이름 가져오기
Function GetCardName(cardId)
    If IsNull(cardId) Or cardId = "" Then
        GetCardName = "-"
        Exit Function
    End If
    
    Dim cardName, cardSQL, cardRS
    cardSQL = "SELECT account_name FROM " & dbSchema & ".CardAccount WHERE card_id = " & cardId
    
    On Error Resume Next
    Set cardRS = db.Execute(cardSQL)
    
    If Err.Number = 0 And Not cardRS.EOF Then
        cardName = cardRS("account_name")
    Else
        cardName = cardId
    End If
    
    If Not cardRS Is Nothing Then
        If cardRS.State = 1 Then
            cardRS.Close
        End If
        Set cardRS = Nothing
    End If
    
    GetCardName = cardName
End Function

' 계정과목명 가져오기
Function GetExpenseCategoryName(categoryId)
    If IsNull(categoryId) Or categoryId = "" Then
        GetExpenseCategoryName = "-"
        Exit Function
    End If
    
    Dim categoryName, catSQL, catRS
    catSQL = "SELECT type_name FROM " & dbSchema & ".CardAccountTypes WHERE account_type_id = " & categoryId
    
    On Error Resume Next
    Set catRS = db.Execute(catSQL)
    
    If Err.Number = 0 And Not catRS.EOF Then
        categoryName = catRS("type_name")
    Else
        categoryName = categoryId
    End If
    
    If Not catRS Is Nothing Then
        If catRS.State = 1 Then
            catRS.Close
        End If
        Set catRS = Nothing
    End If
    
    GetExpenseCategoryName = categoryName
End Function

' 판관/제조명 가져오기
Function GetCostTypeName(costTypeId)
    If IsNull(costTypeId) Or costTypeId = "" Then
        GetCostTypeName = "-"
        Exit Function
    End If

    Dim costTypeName, costTypeSQL, costTypeRS
    costTypeSQL = "SELECT type_name FROM " & dbSchema & ".Cost_Type WHERE cost_type_id = " & costTypeId 
    
    On Error Resume Next
    Set costTypeRS = db.Execute(costTypeSQL)
    
    If Err.Number = 0 And Not costTypeRS.EOF Then
        costTypeName = costTypeRS("type_name")
    Else
        costTypeName = costTypeId
    End If

    If Not costTypeRS Is Nothing Then
        If costTypeRS.State = 1 Then
            costTypeRS.Close
        End If
        Set costTypeRS = Nothing
    End If

    GetCostTypeName = costTypeName
End Function


' 카드 사용 내역 조회
Dim SQL, rs
SQL = "SELECT usage_id, user_id, card_id, department_id, expense_category_id, usage_date, " & _
      "store_name, amount, purpose, linked_table, cost_type_id, created_at, approval_status, title " & _
      "FROM " & dbSchema & ".CardUsage " & _
      "WHERE usage_id = " & usageId

' SQL 쿼리 실행 전에 확인
Response.Write("<!-- 디버그: SQL 쿼리: " & SQL & " -->")

On Error Resume Next
Set rs = db.Execute(SQL)

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
If rs.EOF Then
    errorMsg = "요청하신 카드 사용 내역을 찾을 수 없습니다."
    Session("error_msg") = errorMsg
    Response.Write("<meta http-equiv='refresh' content='0;url=card_usage.asp'>")
    Response.End
End If

On Error GoTo 0

' 현재 레코드의 ID 저장
Dim currentUsageId
currentUsageId = rs("usage_id")
%>


<script>
    function confirmDelete(id) {
        if (confirm("정말로 이 카드 사용 내역을 삭제하시겠습니까? 이 작업은 되돌릴 수 없습니다.")) {
            window.location.href = "card_usage_delete.asp?id=" + id;
        }
    }
    </script>
<!--#include file="../includes/header.asp"-->

<% If errorMsg <> "" Then %>
<div class="alert alert-danger" role="alert">
    <%= errorMsg %>
</div>
<% End If %>

<% If Not rs.EOF Then %>
<div class="vehicle-request-edit-container">
    <div class="shadcn-card" style="max-width: 800px; margin: 30px auto;">
        <div class="shadcn-card-header">
            <div style="display: flex; justify-content: space-between; align-items: center;">
                <h2 class="shadcn-card-title">카드 사용 내역 상세</h2>
                <div>
                    <div class="text-center mt-4">
                        <% If rs("approval_status") = "대기" Or rs("approval_status") = "반려" Then %>
                            <a href="card_usage_edit.asp?id=<%= currentUsageId %>" class="btn btn-primary me-2">
                                <i class="fas fa-edit me-1"></i> 수정
                            </a>
                        <% End If %>
                        <% If rs("approval_status") <> "완료" Then %>
                            <button type="button" class="btn btn-destructive me-2" onclick="confirmDelete(<%= currentUsageId %>)">삭제</button>
                        <% End If %>
                        <% If rs("approval_status") = "완료" Then %>
                            <a href="approval_detail.asp?id=<%= currentUsageId %>&type=CardUsage" class="shadcn-btn shadcn-btn-secondary me-2" style="background-color: #2C3E50; color: white; border: none; padding: 0.75rem 1.5rem; border-radius: 8px; font-weight: 600; text-decoration: none; display: inline-flex; align-items: center; transition: all 0.2s ease;">
                                <i class="fas fa-file-alt" style="margin-right: 0.5rem;"></i>결재 정보 상세
                            </a>
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
                            <span class="detail-label">판관/제조</span> 
                            <span class="detail-value"><%= GetCostTypeName(rs("cost_type_id")) %></span>
                            
                        </div>
                        <div class="detail-item">
                            <span class="detail-label">승인 상태</span>
                            <span class="detail-value">
                                <span class="status-badge status-<%= LCase(rs("approval_status")) %>">
                                    <%= rs("approval_status") %>
                                </span>
                            </span>
                        </div>
                        <div class="detail-item">
                            <span class="detail-label">사용처</span>
                            <span class="detail-value"><%= rs("store_name") %></span>
                        </div>
                        <div class="detail-item">
                            <span class="detail-label">사용목적</span> 
                            <span class="detail-value"><%= rs("purpose") %></span>
                        </div>
                    </div>
                </div>
                
                <div class="detail-section">
                    <h3 class="detail-heading">제목</h3>
                    <div class="detail-item full-width">
                        <div class="detail-value">
                            <% If Not IsNull(rs("title")) And Trim(rs("title")) <> "" Then %>
                                <%= rs("title") %>
                            <% End If %>
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
                    <h3 class="detail-heading">등록 정보</h3>
                    <div class="detail-grid">
                        <div class="detail-item">
                            <span class="detail-label">등록일</span>
                            <span class="detail-value"><%= FormatDateTime(rs("created_at"), 2) %></span>
                        </div>
                        
                    </div>
                </div>
            </div>
        </div>
        
        <% If Session("user_id") = rs("user_id") And (rs("approval_status") = "대기" Or rs("approval_status") = "완료") Then %>
        <a href="approval_detail.asp?id=<%= rs("usage_id") %>&type=CardUsage" class="shadcn-btn" style="background-color: #2C3E50; color: white; border: none; padding: 0.75rem 1.5rem; border-radius: 8px; font-weight: 600; text-decoration: none; display: inline-flex; align-items: center; transition: all 0.2s ease;">
            <i class="fas fa-file-alt" style="margin-right: 0.5rem;"></i>결재 정보 상세
        </a>
        <% End If %>
        
    </div>
</div>
<% End If %>

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