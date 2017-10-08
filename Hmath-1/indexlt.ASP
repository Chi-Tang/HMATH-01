 <!-- #include virtual="/kjasp/func/DB.fun" -->

<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<title></title>
<base target="main">
</head>

<body background="../hmath/bg%20flowers1.bmp">

 
 <% 
   
 MySelf = Request.ServerVariables("PATH_INFO")

ITEM=Request("ITEM")
SECTM=Request("SECTM")
Dim URL
 URL = SECTM&ITEM
 'URL = SECTM&"login-1.asp" 
If Request("Send") <> Empty Then

     ' Response.Redirect "Test.asp?" & Request.QueryString
     'Response.Redirect "TestKa.asp?" & Request.QueryString
     Response.Redirect URL 
   ''' window.open URL  
    ''' window.open URL 
 End If     
   
  %>    

<h2 align="center">數學天地</h2>
<h4>請選年級: 　　　　 &項目:</h4>                                                        
<blockquote>
  <form action="<%=Myself%>" method="GET"> 
    <p><select name="SECTM" size="2">                        
            <option value="http://61.222.248.67/Hmath-1t/">高一(上)</option> 
            <option value="http://61.222.248.67/Hmath-3/">高二(上)</option> 
            <option value="http://61.222.248.67/Hmath-4/">高二(下)</option>
            <option value="1050501.html">END</option> 
        </select>    
      <select name="ITEM" size="3">                                             
         <option value="login-1r.asp">閱讀課本</option>
         <option value="login-1t.asp">逐題練習</option> 
          <option value="login-1.asp">平常測驗</option>
       </select> 
      <!.../p...>                                                                                                                                                                             
      <p><input type="submit" Name="Send" value=" 進入考場 "> </p> 
    </form> 
  </blockquote>  
<h2 align="center"><a href="http://ms.ho.net.tw/" >休閒園地</a></h2>
<h2 align="center"><a href="/Hmath/Rm-ch11/gform.htm" >留言版</a></h2>

<hr> 
<FONT Color=Red><%=Msg%></FONT> 
</boody> 
</html>




































