<!-- #include virtual="/kjasp/func/DB.fun" -->
<%
Lesson = Request("Lesson")
No = Request("No")
Name = Request("Name")
'Set rs = GetMdbRecordset( "Testa.mdb", Lesson )
Set rs = GetMdbRecordset( "Testac-1.mdb", Lesson&Trim(CStr(No)))
%>
<HTML>
<BODY BgColor=White Background="../B01.jpg">
<H2>�Юֹﵪ��<HR></H2>

''<%
'Score = 0
'While Not rs.EOF
 ' Sel = Request( "No" & rs("�D��") )
  ' rs("����") = Sel
  'Ans = rs("�ѵ�")
   'If Ans = Sel Then
    ' rs("�ֵ�") = "V"
  ' Else
   '   rs("�ֵ�") = "X"
   'End If
  'rs.Update
  
  ''If Ans = Sel Then
  ''   Score = Score + rs("����")
  ''End If
  
  'rs.MoveNext
  
'Wend

''SQL = "Select * From ���Z�� " 
  'SQL = SQL & "Where �Ǹ�=" & No & " And �m�W='" & Name & "'"
  'Set rsScore = GetMdbRecordset( "Test.mdb", SQL )
''SQL = SQL & "Where �Ǹ�=" & No 
''Set rsScore = GetMdbRecordset( "Testac.mdb", SQL )
''rsScore(Lesson) = score
''rsScore.Update

' Set cmd = Server.CreateObject( "ADODB.Command" )
 ' Set cmd.ActiveConnection = rs.ActiveConnection
  'SQLU ="Update "&Lesson&Trim(CStr(No))&" Set �аO=-1"  
   ' cmd.CommandText = SQLU
    'cmd.Execute

' SQL = "Select * From " & Lesson&Trim(CStr(No))
 ' Set rs = GetMdbRecordset( "Testac-1.mdb", SQL )
 '%>
 
 <TABLE Border=1>
  <TR BgColor=Yellow><TD>�D��</TD><TD>�ֵ�</TD><TD>���� </TD>><TD>�ѵ�  </TD><TD>�D��</TD><TD>�D��</TD><TD>�Ե�</TD></TR>
<%
While Not rs.EOF
   crsTN ="<A HREF ="&Trim(Cstr(rs("�y����")))&".html"&">�Ե�����</A>"
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
<A HREF="EnterKac-tt.asp?No=<%=No%>&Name=<%=Name%>">�ѥ[��L��ئҸ�</A>
<!--A HREF="Enter.asp?No=<%=No%>&Name=<%=Name%>">�ѥ[��L��ئҸ�</A-->
<%
   Showmg "��²��"
 %> 

</BODY>
</HTML>

 <%
 Sub Showmg( msg )
   
  %>
<CENTER>
<H2><%=msg%><HR></H2>
<FORM>
  <INPUT Type=Button Value="�Юֵ���^�W�@��" OnClick='datachk();'></FORM>
</CENTER>
<% 
End Sub 
%>

<%
Function datachk()
Dim yn
yn=MsgBox("you summit?", 65, "caution!")
 if yn=2 then 
   Exit Function
  end if
     
%>

  <TABLE Border=1>
  <TR BgColor=Yellow>
  <TD Align=Left><%=HMCLD(crsTM) %> </TD> 
 
  </TABLE >
<%  
  End Function
%>

  
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























































