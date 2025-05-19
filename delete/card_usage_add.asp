' 카드 사용내역 저장 후 결재라인 저장
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    ' 트랜잭션 시작
    db.BeginTrans
    
    ' 카드 사용내역 저장
    Dim cardUsageSQL
    cardUsageSQL = "INSERT INTO " & dbSchema & ".CardUsage (" & _
                   "user_id, card_name, store_name, amount, usage_date, purpose, approval_status, created_at" & _
                   ") VALUES (" & _
                   "'" & PreventSQLInjection(Session("user_id")) & "', " & _
                   "'" & PreventSQLInjection(Request.Form("card_name")) & "', " & _
                   "'" & PreventSQLInjection(Request.Form("store_name")) & "', " & _
                   Request.Form("amount") & ", " & _
                   "'" & PreventSQLInjection(Request.Form("usage_date")) & "', " & _
                   "'" & PreventSQLInjection(Request.Form("purpose")) & "', " & _
                   "'대기', " & _
                   "GETDATE())"

    db.Execute(cardUsageSQL)
    
    If Err.Number <> 0 Then
        db.RollbackTrans
        Response.Write "<script>alert('카드 사용내역 저장 중 오류가 발생했습니다: " & Err.Description & "'); history.back();</script>"
        Response.End
    End If
    
    ' 새로 생성된 usage_id 조회
    Dim usageIdSQL, usageIdRS
    usageIdSQL = "SELECT IDENT_CURRENT('CardUsage') as new_id"
    Set usageIdRS = db.Execute(usageIdSQL)
    
    If Err.Number <> 0 Then
        db.RollbackTrans
        Response.Write "<script>alert('사용내역 ID 조회 중 오류가 발생했습니다: " & Err.Description & "'); history.back();</script>"
        Response.End
    End If
    
    Dim newUsageId
    newUsageId = usageIdRS("new_id")
    
    ' 결재라인 저장
    Dim approverIds(3)
    approverIds(1) = Request.Form("approver_step1")
    approverIds(2) = Request.Form("approver_step2")
    approverIds(3) = Request.Form("approver_step3")
    
    Dim i
    For i = 1 To 3
        If approverIds(i) <> "" Then
            Dim approvalLogSQL
            approvalLogSQL = "INSERT INTO " & dbSchema & ".ApprovalLogs (" & _
                            "target_table_name, target_id, approver_id, approval_step, status, created_at" & _
                            ") VALUES (" & _
                            "'CardUsage', " & _
                            newUsageId & ", " & _
                            "'" & PreventSQLInjection(approverIds(i)) & "', " & _
                            i & ", " & _
                            "'대기', " & _
                            "GETDATE())"
            
            db.Execute(approvalLogSQL)
            
            If Err.Number <> 0 Then
                db.RollbackTrans
                Response.Write "<script>alert('결재라인 저장 중 오류가 발생했습니다: " & Err.Description & "'); history.back();</script>"
                Response.End
            End If
        End If
    Next
    
    ' 트랜잭션 커밋
    db.CommitTrans
    
    ' 성공 메시지와 함께 목록 페이지로 리다이렉트
    Response.Write "<script>alert('카드 사용내역이 등록되었습니다.'); location.href='dashboard.asp';</script>"
    Response.End
End If 