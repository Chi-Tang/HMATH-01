<!-- #include virtual="/kjasp/func/DB.fun" -->
<%
Lesson = Request("Lesson")
No = Request("No")
Name = Request("Name")
'Set rs = GetMdbRecordset( "Test.mdb", Lesson )
Set rs = GetMdbRecordset( "Testa.mdb", Lesson )
%>

<HTML>
<BODY BgColor=White Background="../B01.jpg">
<H2>�Ҹզ��Z<HR></H2>
<%
Score = 0
While Not rs.EOF
  Sel = Request( "No" & rs("�D��") )
  Ans = rs("�ѵ�")
  If Ans = Sel Then
     Score = Score + rs("����")
  End If
  rs.MoveNext
Wend

SQL = "Select * From ���Z�� " 
'SQL = SQL & "Where �Ǹ�=" & No & " And �m�W='" & Name & "'"
'Set rsScore = GetMdbRecordset( "Test.mdb", SQL )
SQL = SQL & "Where �Ǹ�=" & No 
Set rsScore = GetMdbRecordset( "Testa.mdb", SQL )

rsScore(Lesson) = score
rsScore.Update
%>
<BLOCKQUOTE>
�Ǹ��G<%=No%><BR>
�m�W�G<%=Name%><BR>
  <BLOCKQUOTE>
  <TABLE Border=1>
  <TR BgColor=Yellow><TD>���</TD><TD>����</TD></TR>
  <TR><TD>ASP</TD><TD Align=Right><%=ScoreHTML(rsScore("ASP"))%></TD></TR>
  <TR><TD>BCC</TD><TD Align=Right><%=ScoreHTML(rsScore("BCC"))%></TD></TR>
  <TR><TD>VB</TD><TD Align=Right><%=ScoreHTML(rsScore("VB"))%></TD></TR>
  </TABLE>
  </BLOCKQUOTE>
</BLOCKQUOTE>
<HR>
<A HREF="EnterKa.asp?No=<%=No%>&Name=<%=Name%>">�ѥ[��L��ئҸ�</A>
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







