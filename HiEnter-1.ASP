 <!-- #include virtual="/kjasp/func/DB.fun" -->
<%
 MySelf = Request.ServerVariables("PATH_INFO")

SECTM=Request("SECTM")
Dim URL
 URL = SECTM 
If Request("Send") <> Empty Then

     ' Response.Redirect "Test.asp?" & Request.QueryString
     'Response.Redirect "TestKa.asp?" & Request.QueryString
     Response.Redirect URL 
     
 End If     
   
  %>    

<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<title></title>
<base target="main">
</head>

<body background="bg%20flowers1.bmp">

<h2 align="center">數學天地</h2>
<h4>請選年級</h4>
<blockquote>
    <form action="<%=Myself%>" method="GET"> 
          
     <p><select name="SECTM" multiple size="2"> 
            <option value="http://61.222.248.67/Hmath/login-1.asp">高一(上)</option> 
            <option value="http://61.222.248.67/Hmath/1010102.html">高一(下)</option> 
            <option value="1010103.html">1-1-3</option> 
            <option value="1010400">1-1-4</option> 
            <option value="1020100">1-2-1</option> 
            <option value="1020200">1-2-2</option> 
            <option value="1020300">1-2-3</option> 
            <option value="1030100">1-3-1</option> 
            <option value="1030200">1-3-2</option> 
            <option value="1030300">1-3-3</option> 
            <option value="1040100">1-4-1</option> 
            <option value="1040200">1-4-2</option> 
            <option value="1040300">1-4-3</option> 
            <option value="1040400">1-4-4</option> 
            <option value="1050100">1-5-1</option> 
            <option value="1050200">1-5-2</option> 
            <option value="1050300">1-5-3</option> 
            <option value="1050400">1-5-4</option> 
            <option value="1050500">1-5-5</option> 

      </select>
      <!.../p...>                                                                                                 
      <p><input type="submit" Name="Send" value=" 進入考場 "> </p> 
    </form> 
  </blockquote>  
<h2 align="center">休閒園地</h2>
<h2 align="center">留言版</h2>

<hr> 
<FONT Color=Red><%=Msg%></FONT> 
</boody> 
</html>




























