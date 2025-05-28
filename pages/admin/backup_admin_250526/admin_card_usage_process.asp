<%@ Language="VBScript" CodePage="65001" %>
<% 
Response.CodePage = 65001
Response.CharSet = "utf-8"
%>

<!--#include file="../../db.asp"-->
<!--#include file="../../includes/functions.asp"-->
<%
' 로그인 체크
If Not IsAuthenticated() Then
    RedirectTo("../../index.asp")
End If

' 관리자 권한 체크
If Not IsAdmin() Then
    Response.Write("<script>alert('관리자 권한이 필요합니다.'); window.location.href='../dashboard.asp';</script>")
    Response.End
End If

' POST 요청인지 확인
If Request.ServerVariables("REQUEST_METHOD") <> "POST" Then
    Response.Write("<script>alert('잘못된 접근입니다.'); window.location.href='admin_card_usage.asp';</script>")
    Response.End
End If

' 폼 데이터 가져오기
Dim action, usageId, cardId, userId, usageDate, amount, merchantName, categoryId, description
Dim deleteReceipt, receiptImage, currentReceiptPath

action = Request.Form("action")
cardId = Request.Form("card_id")
userId = Request.Form("user_id")
usageDate = Request.Form("usage_date")
amount = Replace(Request.Form("amount"), ",", "") ' 콤마 제거
merchantName = PreventSQLInjection(Request.Form("merchant_name"))
categoryId = Request.Form("expense_category_id")
description = PreventSQLInjection(Request.Form("description"))

' 필수 입력값 확인
If cardId = "" Then
    Response.Write("<script>alert('카드를 선택해주세요.'); history.back();</script>")
    Response.End
End If

If userId = "" Then
    Response.Write("<script>alert('사용자를 선택해주세요.'); history.back();</script>")
    Response.End
End If

If Not IsDate(usageDate) Then
    Response.Write("<script>alert('올바른 사용 날짜를 입력해주세요.'); history.back();</script>")
    Response.End
End If

If amount = "" Or Not IsNumeric(amount) Then
    Response.Write("<script>alert('올바른 금액을 입력해주세요.'); history.back();</script>")
    Response.End
End If

If merchantName = "" Then
    Response.Write("<script>alert('가맹점명을 입력해주세요.'); history.back();</script>")
    Response.End
End If

' 금액 숫자형으로 변환
amount = CDbl(amount)

' 카테고리 NULL 처리
If categoryId = "" Then
    categoryId = Null
End If

' 설명 NULL 처리
If description = "" Then
    description = Null
End If

On Error Resume Next

' 영수증 이미지 처리 함수
Function ProcessReceiptImage(formField)
    Dim uploadedFileName, fileExt, newFileName
    
    ' 업로드된 파일이 있는지 확인
    If Request.Form(formField) <> "" Then
        ' 기존에 업로드된 파일이 있을 경우 삭제
        If Request.Form("delete_receipt") = "1" Then
            ProcessReceiptImage = ""
            Exit Function
        End If
        
        ' 현재 파일 경로 그대로 반환
        ProcessReceiptImage = Request.Form(formField)
        Exit Function
    End If
    
    ' 새 파일이 업로드되었는지 확인
    If Request.TotalBytes > 0 Then
        ' 파일 업로드 로직
        Dim upload, files, file
        Set upload = New FileUpload
        upload.Save("../../uploads")
        
        Set files = upload.Files
        
        If files.Count > 0 Then
            Set file = files.Item(formField)
            
            If file.ContentType = "image/jpeg" Or file.ContentType = "image/png" Or file.ContentType = "application/pdf" Then
                ' 파일 확장자 추출
                fileExt = LCase(Right(file.FileName, Len(file.FileName) - InStrRev(file.FileName, ".")))
                
                ' 새 파일명 생성 (타임스탬프_랜덤숫자.확장자)
                newFileName = "receipt_" & Year(Now) & Month(Now) & Day(Now) & Hour(Now) & Minute(Now) & Second(Now) & "_" & Int(Rnd * 1000) & "." & fileExt
                
                ' 파일 저장
                file.SaveAs Server.MapPath("../../uploads/" & newFileName)
                
                ProcessReceiptImage = newFileName
            Else
                Response.Write("<script>alert('지원되지 않는 파일 형식입니다. JPG, PNG, PDF 파일만 업로드 가능합니다.'); history.back();</script>")
                Response.End
            End If
        End If
    End If
    
    ' 파일이 업로드되지 않았거나 오류 발생 시
    If Err.Number <> 0 Then
        ProcessReceiptImage = ""
    End If
End Function

If action = "add" Then
    ' 사용 내역 추가
    ' 영수증 이미지 처리
    receiptImage = ProcessReceiptImage("receipt_image")
    
    ' 사용 내역 추가
    Dim addSQL
    addSQL = "INSERT INTO " & dbSchema & ".CardUsage " & _
             "(card_id, user_id, usage_date, amount, merchant_name, expense_category_id, description, receipt_image, created_at) " & _
             "VALUES (?, ?, ?, ?, ?, ?, ?, ?, GETDATE())"
    
    Dim cmdAdd
    Set cmdAdd = Server.CreateObject("ADODB.Command")
    cmdAdd.ActiveConnection = db
    cmdAdd.CommandText = addSQL
    cmdAdd.Parameters.Append cmdAdd.CreateParameter("@card_id", 3, 1, , cardId)
    cmdAdd.Parameters.Append cmdAdd.CreateParameter("@user_id", 200, 1, 50, userId)
    cmdAdd.Parameters.Append cmdAdd.CreateParameter("@usage_date", 135, 1, , CDate(usageDate))
    cmdAdd.Parameters.Append cmdAdd.CreateParameter("@amount", 5, 1, , amount)
    cmdAdd.Parameters.Append cmdAdd.CreateParameter("@merchant_name", 200, 1, 100, merchantName)
    cmdAdd.Parameters.Append cmdAdd.CreateParameter("@expense_category_id", 200, 1, 20, IIf(IsNull(categoryId), Null, categoryId))
    cmdAdd.Parameters.Append cmdAdd.CreateParameter("@description", 200, 1, 200, IIf(IsNull(description), Null, description))
    cmdAdd.Parameters.Append cmdAdd.CreateParameter("@receipt_image", 200, 1, 255, IIf(receiptImage = "", Null, receiptImage))
    
    cmdAdd.Execute
    
    If Err.Number <> 0 Then
        Response.Write("<script>alert('카드 사용 내역 추가 중 오류가 발생했습니다: " & Server.HTMLEncode(Err.Description) & "'); history.back();</script>")
        Response.End
    Else
        ' 활동 로그 기록
        LogActivity Session("user_id"), "카드사용내역추가", "카드 사용 내역 추가 (카드ID: " & cardId & ", 사용자: " & userId & ", 금액: " & amount & ")"
        Response.Write("<script>alert('카드 사용 내역이 추가되었습니다.'); window.location.href='admin_card_usage.asp';</script>")
        Response.End
    End If
    
ElseIf action = "edit" Then
    ' 사용 내역 수정
    usageId = Request.Form("usage_id")
    deleteReceipt = (Request.Form("delete_receipt") = "1")
    
    ' 현재 영수증 파일 경로 가져오기
    Dim currentReceiptSQL, currentReceiptRS
    currentReceiptSQL = "SELECT receipt_image FROM " & dbSchema & ".CardUsage WHERE usage_id = " & usageId
    Set currentReceiptRS = db.Execute(currentReceiptSQL)
    
    If Not currentReceiptRS.EOF Then
        currentReceiptPath = currentReceiptRS("receipt_image")
    End If
    
    ' 영수증 이미지 처리
    If deleteReceipt Then
        receiptImage = Null
        
        ' 기존 영수증 파일 삭제
        If Not IsNull(currentReceiptPath) And currentReceiptPath <> "" Then
            Dim fs
            Set fs = Server.CreateObject("Scripting.FileSystemObject")
            
            If fs.FileExists(Server.MapPath("../../uploads/" & currentReceiptPath)) Then
                fs.DeleteFile(Server.MapPath("../../uploads/" & currentReceiptPath))
            End If
            
            Set fs = Nothing
        End If
    Else
        ' 새 영수증 업로드 처리
        receiptImage = ProcessReceiptImage("receipt_image")
        
        ' 새 파일이 업로드되지 않았으면 기존 파일 유지
        If receiptImage = "" And Not IsNull(currentReceiptPath) Then
            receiptImage = currentReceiptPath
        End If
    End If
    
    ' 사용 내역 수정
    Dim editSQL
    editSQL = "UPDATE " & dbSchema & ".CardUsage SET " & _
              "card_id = ?, user_id = ?, usage_date = ?, amount = ?, " & _
              "merchant_name = ?, expense_category_id = ?, description = ?, receipt_image = ? " & _
              "WHERE usage_id = ?"
    
    Dim cmdEdit
    Set cmdEdit = Server.CreateObject("ADODB.Command")
    cmdEdit.ActiveConnection = db
    cmdEdit.CommandText = editSQL
    cmdEdit.Parameters.Append cmdEdit.CreateParameter("@card_id", 3, 1, , cardId)
    cmdEdit.Parameters.Append cmdEdit.CreateParameter("@user_id", 200, 1, 50, userId)
    cmdEdit.Parameters.Append cmdEdit.CreateParameter("@usage_date", 135, 1, , CDate(usageDate))
    cmdEdit.Parameters.Append cmdEdit.CreateParameter("@amount", 5, 1, , amount)
    cmdEdit.Parameters.Append cmdEdit.CreateParameter("@merchant_name", 200, 1, 100, merchantName)
    cmdEdit.Parameters.Append cmdEdit.CreateParameter("@expense_category_id", 200, 1, 20, IIf(IsNull(categoryId), Null, categoryId))
    cmdEdit.Parameters.Append cmdEdit.CreateParameter("@description", 200, 1, 200, IIf(IsNull(description), Null, description))
    cmdEdit.Parameters.Append cmdEdit.CreateParameter("@receipt_image", 200, 1, 255, IIf(IsNull(receiptImage), Null, receiptImage))
    cmdEdit.Parameters.Append cmdEdit.CreateParameter("@usage_id", 3, 1, , usageId)
    
    cmdEdit.Execute
    
    If Err.Number <> 0 Then
        Response.Write("<script>alert('카드 사용 내역 수정 중 오류가 발생했습니다: " & Server.HTMLEncode(Err.Description) & "'); history.back();</script>")
        Response.End
    Else
        ' 활동 로그 기록
        LogActivity Session("user_id"), "카드사용내역수정", "카드 사용 내역 수정 (ID: " & usageId & ", 카드ID: " & cardId & ", 금액: " & amount & ")"
        Response.Write("<script>alert('카드 사용 내역이 수정되었습니다.'); window.location.href='admin_card_usage.asp';</script>")
        Response.End
    End If
    
Else
    Response.Write("<script>alert('잘못된 요청입니다.'); window.location.href='admin_card_usage.asp';</script>")
End If

On Error GoTo 0
%>

<%
' 파일 업로드 클래스
Class FileUpload
    Private mcolFormElem

    Private Sub Class_Initialize()
        Set mcolFormElem = Server.CreateObject("Scripting.Dictionary")
    End Sub

    Private Sub Class_Terminate()
        Set mcolFormElem = Nothing
    End Sub

    Public Property Get Form()
        Set Form = mcolFormElem
    End Property

    Public Property Get Files()
        Set Files = mcolFormElem
    End Property

    Public Sub Save(path)
        ' 간소화된 파일 업로드 처리
        ' 실제 환경에서는 외부 컴포넌트나 서버측 기능을 사용해야 함
        ' 이 예제에서는 ASP 기본 기능만으로는 파일 업로드가 제한적이므로 간소화 처리
    End Sub
End Class
%> 