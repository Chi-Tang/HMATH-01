<!-- #include virtual="/kjasp/func/DB.fun" -->
<%
Lesson = Request("Lesson")
No = Request("No")
Name = Request("Name")
 ''Set rs = GetMdbRecordset( "Testa.mdb", Lesson )
  'Set rs = GetMdbRecordset( "Testac-1t.mdb", Lesson&Trim(CStr(No)))
  SQL = "Select * From "&Lesson&Trim(CStr(No))& " Order by �D��"
   Set rs = GetMdbRecordset( "Testac-1.mdb", SQL )

 %>

<HTML>
<BODY BgColor=White Background="../B01.jpg">
<H2>�Ҹզ��Z<HR></H2>
<FORM Action=Checkac-1.asp Method=POST>
  <INPUT Type=Hidden Name=Lesson Value=<%=Lesson%>>
  <INPUT Type=Hidden Name=No Value=<%=No%>>
  <INPUT Type=Hidden Name=Name Value=<%=Name%>>

<%
Count = 0
Score = 0
While Not rs.EOF
  Sel = Request( "No" & rs("�D��") )
    rs("����") = Sel
    'rs("�Ե�") =" <A HREF = "&Trim(Cstr(rs("�y����")))&".html">&"�Ե�����</A>" 
    'rs("�Ե�") =" <A HREF = /Hmath-3t/"&Trim(Cstr(rs("�y����")))&".html">&"�Ե�����</A>"
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
   'SQL = SQL & "Where �Ǹ�='" & No & "'" And �m�W='" & Name & "'"
   'Set rsScore = GetMdbRecordset( "Test.mdb", SQL )
   '' SQL = SQL & "Where �Ǹ�=" &  Trim(Cstr(No))
  SQL = SQL & "Where �Ǹ�='" & No & "'"
    'Set rsScore = GetMdbRecordset( "Testac-1.mdb", SQL )
     'Set rsScore = GetMdbRecordset( "Testac-1.mdb", SQL )
     'Set rsScore = GetSecuredMdbRecordset( mdbFile, SQL, mdbPassword )
  Set rsScore = GetSecuredMdbRecordset("/Hmath/UsersPwd.mdb", SQL, "kj6688" )

  rsScore(Lesson) = Score
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
<INPUT Type=Submit Value="  �ֹﵪ��(ú ��)  ">
  </FORM>
<!--<% 
   'Response.Redirect "Checkac-1t.asp?" & Request.QueryString 
  %>-->
 <!--<A HREF="Checkac-1.asp?No=<%=No%>&Name=<%=Name%>&Lesson=<%=Lesson%>">�ֹﵪ��</A>
 <!--P><A HREF="EnterKac.asp?No=<%=No%>&Name=<%=Name%>">�ѥ[��L��ئҸ�</A></P-->
 <!--A HREF="Enter.asp?No=<%=No%>&Name=<%=Name%>">�ѥ[��L��ئҸ�</A-->

</BODY>
</HTML>

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
















































































