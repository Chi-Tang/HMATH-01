<!-- #include virtual="/kjasp/func/DB.fun" -->
<%
Lesson = Request("Lesson")
No = Request("No")
Name = Request("Name")
'Checkyn = Request("Checkyn")
'Set rs = GetMdbRecordset( "Testa.mdb", Lesson )
Set rs = GetMdbRecordset( "Testac-1.mdb", Lesson&Trim(CStr(No)))

%>
<HTML>
<BODY BgColor=White Background="../B01.jpg">
<H2>�Ҹզ��Z<HR></H2>
    
<%
Count = 0
Score = 0
While Not rs.EOF
  Sel = Request( "No" & rs("�D��") )
    rs("����") = Sel
    'rs("�Ե�") =" <A HREF = /KJASP/math/"&Trim(Cstr(rs("�y����")))&".html">&"�Ե�����</A>" 
   Ans = rs("�ѵ�")
   If Ans = Sel Then
     rs("�ֵ�") = "V"
     Score = Score + rs("����")
   Else
      rs("�ֵ�") = "X"
   End If
  rs.Update

  '  If Ans = Sel Then
  '  Score = Score + rs("����")
  'End If
 Count = Count+1 
  rs.MoveNext
Wend
 
SQL = "Select * From ���Z�� " 
'SQL = SQL & "Where �Ǹ�=" & No & " And �m�W='" & Name & "'"
'Set rsScore = GetMdbRecordset( "Test.mdb", SQL )
 'SQL = SQL & "Where �Ǹ�=" &  Trim(Cstr(No))
 'SQL = SQL & "Where �Ǹ�=" & No
 SQL = SQL & "Where �Ǹ�='" & No & "'"
  'Set rsScore = GetMdbRecordset( "Testac-1.mdb", SQL )
     'Set rsScore = GetSecuredMdbRecordset( mdbFile, SQL, mdbPassword )
  Set rsScore = GetSecuredMdbRecordset("/Hmath/UsersPwd.mdb", SQL, "kj6688" )

 rsScore(Lesson) = score
  rsScore.Update
 
  
%>

<BLOCKQUOTE>
�Ǹ��G<%=No%><BR>
�m�W�G<%=Name%><BR>
  <BLOCKQUOTE>
  <TABLE Border=1>
  <TR BgColor=Yellow><TD>���</TD><TD>�����D</TD><TD>�`�D��(����)</TD></TR>
  <TR><TD>MATH</TD><TD Align=Right><%=ScoreHTML(rsScore("MATH"))%></TD><TD><%="/"&Count%></TD></TR>
  <TR><TD>ENGLS</TD><TD Align=Right><%=ScoreHTML(rsScore("ENGLS"))%></TD><TD><%="/"&Count%></TD></TR>
  <TR><TD>CHINE</TD><TD Align=Right><%=ScoreHTML(rsScore("CHINE"))%></TD><TD><%="/"&Count%></TD></TR>
  </TABLE>
  </BLOCKQUOTE>
</BLOCKQUOTE>
<HR>
<!--<A HREF="Checkac-1.asp?No=<%=No%>&Name=<%=Name%>&Lesson=<%=Lesson%>">�ֹﵪ��</A>
     <!--<A HREF="Enter.asp?No=<%=No%>&Name=<%=Name%>">�ѥ[��L��ئҸ�</A-->
  <!--<INPUT Type="Submit" name="send" Value=" �ֹﵪ��(ú ��)  ">

 <% 
     '' Response.Redirect "Checkac-1.asp?" & Request.QueryString 
 %>
 
</BODY>
</HTML>-->
 
       
<%
Function ScoreHTML( Score )
   If Score = -1 Then 	' �|���Ҹ�
      ScoreHTML = "--"
   ELseIf Score < 60 Then
      ScoreHTML = "<FONT Color=Red>" & Score & "</FONT>"
   Else
      ScoreHTML = Score
   End If
End Function
%>
 


 <%

'Set rs = GetMdbRecordset( "Testa.mdb", Lesson )
Set rs = GetMdbRecordset( "Testac-1.mdb", Lesson&Trim(CStr(No)))
%>
<!--<HTML>
<BODY BgColor=White Background="../B01.jpg">-->
<H2>�Юֹﵪ��<HR></H2>
 
 <TABLE Border=1>
  <TR BgColor=Yellow><TD>�D��</TD><TD>�ֵ�</TD><TD>���� </TD>><TD>�ѵ�  </TD><TD>�D��</TD><TD>�D��</TD><TD>�Ե�</TD></TR>
<%
While Not rs.EOF
   crsTN ="<A HREF "&Trim(Cstr(rs("�y����")))&".html"&">�Ե�����</A>"
               rs("�Ե�") = crsTN
     ' crsTN ="<A HREF =/Hmath/"&Trim(Cstr(rs("�y����")))&".html"&">�Ե�����</A>"
              ' rs("�Ե�") = crsTN
     

%>
 <TR><TD Align=Right><%=rs("�D��")%></TD>
     <TD Align=Right><%=rs("�ֵ�")%></TD>
     <TD Align=Right><%=rs("����")%></TD>
     <TD Align=Right><%=rs("�ѵ�")%></TD>
   
  <% 
    crsTM =Trim(CStr(rs("�D�ؽX")))&".html"
     'crsTM ="math/"&Trim(CStr(rs("�D�ؽX")))&".html"
     ''crsTM ="1010104.html" & '�ƾǹϧ��ɶ��b�P�ؿ��~�|���
     ' Response.Write crsTM
   %>
       <TD Align=Left><%=HMCLD(crsTM) %> </TD> <TD   
         Align=Right><%=rs("����")%></TD><TD
         Align=Right><%=rs("�Ե�")%></TD>
    
<%
      
  Response.Write "</TR>"

  <!---Response.Write "</BLOCKQUOTE>"--->

  rs.MoveNext
 Wend

 %>
</TABLE>

 <%
  
    Set cmd = Server.CreateObject( "ADODB.Command" )
  Set cmd.ActiveConnection = rs.ActiveConnection
   'SQLU ="Update "&Lesson&Trim(CStr(No))&" Set �аO=-1"  
  SQL2 ="Insert into "&Lesson&Trim(Cstr(No))&"A"& " Select * From "&Lesson&Trim(Cstr(No))&" Where �аO=100"
   cmd.CommandText = SQL2
     cmd.Execute

  SQLD ="Delete From "&Lesson&Trim(CStr(No))
    cmd.CommandText = SQLD
    cmd.Execute
 %> 

<HR>
<A HREF="EnterKac-1.asp?No=<%=No%>&Name=<%=Name%>">�ѥ[��L��ئҸ�</A>
<!--A HREF="Enter.asp?No=<%=No%>&Name=<%=Name%>">�ѥ[��L��ئҸ�</A-->

</BODY>
</HTML>


<% '���հƵ{��2 
  FUNCTION HMCLD(rsTM) 
   Set fs = Server.CreateObject("Scripting.FileSystemObject")
     File = Server.MapPath(rsTM)
  Set txtf = fs.OpenTextFile( File )
If Not txtf.atEndOfStream Then	' ���T�w�٨S����F��������m
    Content = txtf.ReadAll	' Ū������ɮת����
   '' Lines = Replace(Content, vbCrLf, "<BR>" )
    ''Response.Write Lines
   ' Response.Write Content
End If
 HMCLD=Content
END FUNCTION

%>



 
 

































































































