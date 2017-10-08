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
<H2>考試成績<HR></H2>
<%
Score = 0
While Not rs.EOF
  Sel = Request( "No" & rs("題號") )
  Ans = rs("解答")
  If Ans = Sel Then
     Score = Score + rs("分數")
  End If
  rs.MoveNext
Wend

SQL = "Select * From 成績單 " 
'SQL = SQL & "Where 學號=" & No & " And 姓名='" & Name & "'"
'Set rsScore = GetMdbRecordset( "Test.mdb", SQL )
SQL = SQL & "Where 學號=" & No 
Set rsScore = GetMdbRecordset( "Testa.mdb", SQL )

rsScore(Lesson) = score
rsScore.Update
%>
<BLOCKQUOTE>
學號：<%=No%><BR>
姓名：<%=Name%><BR>
  <BLOCKQUOTE>
  <TABLE Border=1>
  <TR BgColor=Yellow><TD>科目</TD><TD>分數</TD></TR>
  <TR><TD>ASP</TD><TD Align=Right><%=ScoreHTML(rsScore("ASP"))%></TD></TR>
  <TR><TD>BCC</TD><TD Align=Right><%=ScoreHTML(rsScore("BCC"))%></TD></TR>
  <TR><TD>VB</TD><TD Align=Right><%=ScoreHTML(rsScore("VB"))%></TD></TR>
  </TABLE>
  </BLOCKQUOTE>
</BLOCKQUOTE>
<HR>
<A HREF="EnterKa.asp?No=<%=No%>&Name=<%=Name%>">參加其他科目考試</A>
<!--A HREF="Enter.asp?No=<%=No%>&Name=<%=Name%>">參加其他科目考試</A-->

</BODY>
</HTML>

<%
Function ScoreHTML( Score )
   If Score = -1 Then 	' 尚未考試
      ScoreHTML = "--"
   ELseIf Score < 60 Then
      ScoreHTML = "<FONT Color=Red>" & Score & "</FONT>"
   Else
      ScoreHTML = Score
   End If
End Function
%>







