 
 <!-- #include virtual="/kjasp/func/DB.fun" -->
 <%
  ''<!-- #include file="Login-1.asp" -->
  %>
<%
 MySelf = Request.ServerVariables("PATH_INFO")
No = Request("No")
Name = Request("Name")
Lesson = Request("Lesson")

SECTM=Request("SECTM")
TNUM=Request("TNUM")
 'SECTMS=SPLIT(SECTM, ",")
 'TNUMS=SPLIT(TNUM, ",")
 'CNm=Request("NUM")

Mssg = "ok"
''On Error Resume Next 

If Request("Send") <> Empty Then
  ' SQL = "Select * From ���Z�� " 
  ' SQL = SQL & "Where �Ǹ�=" & No & " And �m�W='" & Name & "'"
  ' Set rsScore = GetMdbRecordset( "Test.mdb", SQL )
    SQLL = "Select * From "&Lesson&"k" 
    Set rs = GetMdbRecordset( "Testac-1.mdb", SQLL )
     n=0
     TNum1=0
  RNDSN  SECTM,TNUM
    'For  k=0 to Ubound(SECTMS)
     ' RNDSN SECTMS(k),TNUMS(k)
    'NEXT
  
  On Error Resume Next 
  
  'Set conn = GetMdbConnection("Test1.mdb")
  Set cmd = Server.CreateObject( "ADODB.Command" )
  Set cmd.ActiveConnection = rs.ActiveConnection
  'SQLS ="Select * into ASP FROM ASPK"
  ' SQLS ="Select * into "&Lesson& " FROM " &Lesson&"K" �@
  ''SQLD ="Delete From "&Lesson
  ''cmd.CommandText = SQLD
    SQLS1 ="Select * into "&Lesson&Trim(Cstr(No))& " FROM " &Lesson 
    cmd.CommandText = SQLS1
     cmd.Execute
    SQLS2 ="Select * into "&Lesson&Trim(Cstr(No))&"A"& " FROM " &Lesson 
    cmd.CommandText = SQLS2
     cmd.Execute
   
    SQLD1 ="Delete From "&Lesson&Trim(Cstr(No))
  cmd.CommandText = SQLD1
  cmd.Execute
     SQLD2 ="Delete From "&Lesson&Trim(Cstr(No))&"A"
  cmd.CommandText = SQLD2
  cmd.Execute

     
  '' On Error Resume Next 
  ' If Err.Number = 0 Then 
  '    Response.Write Mssg
  '  Else
   '   Response.Write Err.Number
  '  End If
       
  'SQL1 ="Insert into ASP Select * FROM "&Lesson&"K"&"Where �аO=+1"
   SQL1 ="Insert into "&Lesson&Trim(Cstr(No))& " Select * From "&Lesson&"K"&" Where �аO=100"
   cmd.CommandText = SQL1
     cmd.Execute
   SQLU ="Update "&Lesson&"K"&" Set �аO=-1"  
    cmd.CommandText = SQLU
     cmd.Execute
     
     ' Response.Redirect "Test.asp?" & Request.QueryString
    Response.Redirect "Testkac-tt.asp?" & Request.QueryString

End If
%>
<%  '�H�������`�D��
  SUB RNDSN(SECTM, TNUM)
    ' TNUM1=30 
    ' TNUM1=TNUM1+TNUM 
   TNUM1=TNUM+0 
 Cstr(0)&"rn"=-2 
      ' 0&"rn"=-2
      '''�X�j�H���d�� TNUM+10
   'For j=1 to 30 
   For j=1 to 2 
     iRN=SECTM+TNUM1
      'iRN=SECTM+RN
      'iRN=SECTM+j 
     
     rs.MoveFirst
      While Not rs.EOF
        if rs("�y����")=iRN AND rs("�аO")=-1 then
              rs("�аO")=100
              rs("�D�ؽX")=rs("�y����")
              ' crsTN ="<A HREF =/Hmath-1t/"&Trim(Cstr(rs("�y����")))&".html"&">�Ե�����</A>"
              ' rs("�Ե�") = crsTN
              n=n+1
              rs("�D��")=iRN
              'rs("�D��")=n
              rs.Update
             'Response.Write rs("�D��")
         
             'if n>=TNUM1 then
              ' Exit For
             ' end if
          end if
         rs.MoveNext
       Wend  
          ' Response.Write int(iRN)            
      ''''     Response.Write int(iRN) 
    
    Next
   END SUB
  %>    

<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<title></title>
</head>

<body background="../b01.jpg">

<h2 align="center">.���갪���ƾǺ�������</h2>

<hr>

<h2>��ܦҸլ��:</h2>

<blockquote>
    <form action="<%=Myself%>" method="GET"> 
        <p>��ءG<select name="Lesson" size="1"> 
            <option IsSelected("MATH", Lesson)>MATH</option> 
            <option IsSelected("MATH", Lesson)>ENGLS</option> 
            <option IsSelected("MATH", Lesson)>CHINE</option> 
        </select></p> 
        <p>�Ǹ�(�����ҫ᤻�X)�G<input type="text" size="20" name="No" Value="<%=No%>"></p>     
        �m�W�G<input type="text" size="20" name="Name"  Value="<%=Name%>"> 
       <h5>�U-��-�`�@�@�@&�D��:</h5>

         <p>�G<select name="SECTM" size="6"> 
            <option value="1010100">1-1-1</option> 
            <option value="1010200">1-1-2</option> 
            <option value="1010300">1-1-3</option> 
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
          �G<select name="TNUM" size="6">                                                                                                                                                       
            <option value="1">1</option><option value="2">2</option> 
             <option value="3">3</option><option value="4">4</option> 
             <option value="5">5</option><option value="6">6</option> 
             <option value="7">7</option><option value="8">8</option> 
            
            </select></p> 
 

         <p><input type="submit" Name="Send" value=" �i�J�ҳ� "> </p> 
         <p>�@ </p> 
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





























