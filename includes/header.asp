<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>카드 지출 결의/개인차량 이용 관리</title>
    <link rel="stylesheet" href="/contents/card_car_used/css/style.css">
    <!-- shadcn 스타일 파일 -->
    <link rel="stylesheet" href="/contents/card_car_used/css/shadcn.css">
    <!-- 공통 자바스크립트 -->
    <script src="/contents/card_car_used/js/common.js"></script>
</head>
<body>
    <header>
        <div class="container">
            <div class="logo">
                <a href="/contents/card_car_used/index.asp">카드 지출 결의/개인차량 이용 관리</a>
            </div>
            <nav>
                <ul>
                    <%
                    If Session("user_id") <> "" Then
                    %>
                        <li><a href="/contents/card_car_used/pages/card_usage.asp">카드사용 내역</a></li>
                        <li><a href="/contents/card_car_used/pages/vehicle_request.asp">개인차량이용 신청</a></li>
                        <%
                        If Session("is_admin") = "Y" Then
                        %>
                            <li><a href="/contents/card_car_used/pages/admin/admin_dashboard.asp">관리자</a></li>
                        <%
                        End If
                        %>
                        <li><a href="/contents/card_car_used/pages/logout.asp">로그아웃</a></li>
                    <%
                    Else
                    %>
                        <li><a href="/contents/card_car_used/index.asp">로그인</a></li>
                        <li><a href="/contents/card_car_used/pages/register.asp">회원가입</a></li>
                    <%
                    End If
                    %>
                </ul>
            </nav>
        </div>
    </header>
    <main class="container">