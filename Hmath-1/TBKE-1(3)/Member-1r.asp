
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<title>未命名 一般畫面</title>
</head>
<body background="../B01.jpg" bgcolor="#FFFFFF">
<h3>會員登錄</h3>
<hr>
<blockquote>
    <h5>新會員加入</h5>
    <form action="Join-1r.asp" method="POST">
        <blockquote>
            <table border="0">
                <tr>
                    <td align="right">帳號(身分証號)：</td>
                    <td><input type="text" size="20"
                    name="UserID"></td>
                </tr>
                <tr>
                    <td align="right">密碼：</td>
                    <td><input type="password" size="20"
                    name="Password"></td>
                </tr>
                <tr>
                    <td align="right">密碼再確認：</td>
                    <td><input type="password" size="20"
                    name="Password2"></td>
                </tr>
                <tr>
                    <td align="right">姓名：</td>
                    <td><input type="text" size="20" name="Name"></td>
                </tr>
                <tr>
                    <td align="right">Email：</td>
                    <td><input type="text" size="40" name="Email"></td>
                </tr>
            </table>
            <p><input type="submit" value=" 加 入 "> </p>
        </blockquote>
    </form>
</blockquote>
<hr>
<blockquote>
    <h5>我忘了使用者帳號及密碼</h5>
</blockquote>
<blockquote>
    <form action="forget-1.aspx" method="GET">
   <% 
        '<blockquote>
         '   <table border="0">
         '       <tr>
         '           <td>Email：</td>
         '           <td><input type="text" size="40" name="Email"></td>
          '      </tr>
          '  </table>
        ' </blockquote>
     %>   
        <blockquote>
            <p><input type="submit"
            value=" 請傳給我使用者帳號及密碼 "> </p>
        </blockquote>
    </form>
</blockquote>
<hr>
</body>
</html>
