 <!-- #include virtual="/kjasp/func/DB.fun" -->
<%
 MySelf = Request.ServerVariables("PATH_INFO")
No = Request("No")
Name = Request("Name")
Lesson = Request("Lesson")

SECTM=Request("SECTM")
TNUM=Request("TNUM")
SECTMS=SPLIT(SECTM, ",")
TNUMS=SPLIT(TNUM, ",")
CNm=Request("NUM")
Dim URL
 URL = SECTM & ".HTML"
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
</head>

<body background="../b01.jpg">

<h2 align="center">國立鳳新高中網路考試</h2>

<hr>

<h2>選擇考試科目:</h2>

<blockquote>
    <form action="<%=Myself%>" method="GET"> 
        <p>科目：<select name="Lesson" size="1"> 
            <option IsSelected("ASP", Lesson)>ASP</option> 
            <option IsSelected("ASP", Lesson)>BCC</option> 
            <option IsSelected("ASP", Lesson)>VB</option> 
        </select></p> 
        <p>學號：<input type="text" size="20" name="No" Value="<%=No%>"></p>     
        姓名：<input type="text" size="20" name="Name"  Value="<%=Name%>"> 
       <p>冊-章-節：<select name="SECTM" multiple size="6"> 
            <option value="1010101">1-1-1</option> 
            <option value="1010102">link</option> 
            <option value="1010103">1-1-3</option> 
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
          題數：<select name="TNUM" multiple size="6">                                                              
            <option value="1">1</option> 
             <option value="2">2</option> 
             <option value="3">3</option> 
             <option value="4">4</option> 
             <option value="5">5</option> 
             <option value="6">6</option> 
             <option value="7">7</option> 
             <option value="8">8</option> 
            </select></p> 
 

         <p><input type="submit" Name="Send" value=" 進入考場 "> </p> 
         <p>　 </p> 
    </form> 
  </blockquote>  
 
<hr> 
<FONT Color=Red><%=Msg%></FONT> 
</boody> 
</html>

<%  
Function IsSelected( Which, Lesson ) 
   If Which = Lesson Then IsSelected = Selected 
End Function 
%>





























