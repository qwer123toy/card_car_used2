<%@ Language=VBScript %>
<!-- METADATA TYPE="typelib" NAME="ADODB Type Library"
File="C:\Program Files\Common Files\System\ado\msado15.dll" -->
<% Option Explicit %>
<% Response.Expires=-1 %>
<!--#include file="includes/connection.asp"-->
<!--#include file="includes/functions.asp"-->

<!--#include file="includes/header.asp"-->
<head>
    <title></title>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
    <link href="css/style.css" rel="stylesheet" type="text/css">

    
<%
' ?대? 濡쒓렇?명븳 寃쎌슦 硫붿씤 ?섏씠吏濡?由щ뵒?됱뀡
If IsAuthenticated() Then
    RedirectTo("/contents/card_car_used/pages/dashboard.asp")
End If

' 濡쒓렇??泥섎━
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    Dim userId, password, errorMsg, SQL, rs
    
    userId = PreventSQLInjection(Request.Form("user_id"))
    password = PreventSQLInjection(Request.Form("password"))
    
    If userId = "" Or password = "" Then
        errorMsg = "?꾩씠?붿? 鍮꾨?踰덊샇瑜?紐⑤몢 ?낅젰?댁＜?몄슂."
    Else
        ' ?ъ슜???뺤씤
        SQL = "SELECT * FROM Users WHERE user_id = '" & userId & "'"
        Set rs = dbConn.Execute(SQL)
        
        If rs.EOF Then
            errorMsg = "議댁옱?섏? ?딅뒗 ?꾩씠?붿엯?덈떎."
        Else
            ' 鍮꾨?踰덊샇 ?뺤씤 (?ㅼ젣 ?섍꼍?먯꽌???댁떆??鍮꾨?踰덊샇瑜?鍮꾧탳?댁빞 ??
            If rs("password") = password Then
                ' ?몄뀡 ?ㅼ젙
                Session("user_id") = rs("user_id")
                Session("name") = rs("name")
                Session("department_id") = rs("department_id")
                
                ' 愿由ъ옄 ?щ? ?뺤씤
                Dim sqlAdmin, rsAdmin
                sqlAdmin = "SELECT 1 FROM Administrators WHERE user_id = '" & userId & "'"
                Set rsAdmin = dbConn.Execute(sqlAdmin)
                
                If Not rsAdmin.EOF Then
                    Session("is_admin") = "Y"
                Else
                    Session("is_admin") = "N"
                End If
                
                ' 濡쒓렇??湲곕줉
                
                ' ?섏씠吏 ?대룞
                RedirectTo("/contents/card_car_used/pages/dashboard.asp")
            Else
                errorMsg = "鍮꾨?踰덊샇媛 ?쇱튂?섏? ?딆뒿?덈떎."
            End If
        End If
        
        rs.Close
        Set rs = Nothing
    End If
End If
%>
    <head>
        <meta charset="UTF-8">
        <title>?명듃 怨꾩궛 紐⑸줉</title>
    </head>
    <body>
<div class="login-container">
    <div class="shadcn-card" style="max-width: 450px; margin: 80px auto;">
        <div class="shadcn-card-header">
            <h2 class="shadcn-card-title">濡쒓렇??/h2>
            <p class="shadcn-card-description">移대뱶 吏異?寃곗쓽 諛?媛쒖씤李⑤웾 ?댁슜 ?대젰 愿由??쒖뒪?쒖뿉 ?ㅼ떊 寃껋쓣 ?섏쁺?⑸땲??</p>
        </div>
        
        <% If errorMsg <> "" Then %>
        <div class="shadcn-alert shadcn-alert-error">
            <div>
                <span class="shadcn-alert-title">?ㅻ쪟</span>
                <span class="shadcn-alert-description"><%= errorMsg %></span>
            </div>
        </div>
        <% End If %>
        
        <div class="shadcn-card-content">
            <form id="loginForm" method="post" action="/contents/card_car_used/index.asp">
                <div class="form-group">
                    <label class="shadcn-input-label" for="user_id">?꾩씠??/label>
                    <input class="shadcn-input" type="text" id="user_id" name="user_id" placeholder="?꾩씠?붾? ?낅젰?섏꽭??>
                </div>
                
                <div class="form-group">
                    <label class="shadcn-input-label" for="password">鍮꾨?踰덊샇</label>
                    <input class="shadcn-input" type="password" id="password" name="password" placeholder="鍮꾨?踰덊샇瑜??낅젰?섏꽭??>
                </div>
                
                <div class="shadcn-card-footer" style="margin-top: 1.5rem;">
                    <button type="submit" class="shadcn-btn shadcn-btn-primary">濡쒓렇??/button>
                    <a href="/contents/card_car_used/pages/register.asp" class="shadcn-btn shadcn-btn-outline">?뚯썝媛??/a>
                </div>
            </form>
        </div>
    </div>
</div>
</body>
</head>
</html>


<script>
    const loginRules = {
        user_id: {
            required: true,
            message: '?꾩씠?붾? ?낅젰?댁＜?몄슂.'
        },
        password: {
            required: true,
            message: '鍮꾨?踰덊샇瑜??낅젰?댁＜?몄슂.'
        }
    };
</script>

<!--#include file="includes/footer.asp"--> 
