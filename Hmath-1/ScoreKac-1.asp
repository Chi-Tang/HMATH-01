<!-- #include virtual="/kjasp/func/DB.fun" -->
<%
Lesson = Request("Lesson")
No = Request("No")
Name = Request("Name")
 ''Set rs = GetMdbRecordset( "Testa.mdb", Lesson )
  'Set rs = GetMdbRecordset( "Testac-1t.mdb", Lesson&Trim(CStr(No)))
  SQL = "Select * From "&Lesson&Trim(CStr(No))& " Order by 題號"
   Set rs = GetMdbRecordset( "Testac-1.mdb", SQL )

 %>

<HTML>
<BODY BgColor=White Background="../B01.jpg">
<H2>考試成績<HR></H2>
<FORM Action=Checkac-1.asp Method=POST>
  <INPUT Type=Hidden Name=Lesson Value=<%=Lesson%>>
  <INPUT Type=Hidden Name=No Value=<%=No%>>
  <INPUT Type=Hidden Name=Name Value=<%=Name%>>

<%
Count = 0
Score = 0
While Not rs.EOF
  Sel = Request( "No" & rs("題號") )
    rs("做答") = Sel
    'rs("詳答") =" <A HREF = "&Trim(Cstr(rs("流水號")))&".html">&"詳答按此</A>" 
    'rs("詳答") =" <A HREF = /Hmath-3t/"&Trim(Cstr(rs("流水號")))&".html">&"詳答按此</A>"
    Ans = rs("解答")
   If Ans = Sel Then
     rs("核答") = "V"
     Score = Score + rs("分數")
   Else
      rs("核答") = "X"
   End If
  rs.Update

  '  If Ans = Sel Then
  '  Score = Score + rs("分數")
  'End If
  Count = Count+1
  rs.MoveNext
Wend

 SQL = "Select * From 成績單 " 
   'SQL = SQL & "Where 學號='" & No & "'" And 姓名='" & Name & "'"
   'Set rsScore = GetMdbRecordset( "Test.mdb", SQL )
   '' SQL = SQL & "Where 學號=" &  Trim(Cstr(No))
  SQL = SQL & "Where 學號='" & No & "'"
    'Set rsScore = GetMdbRecordset( "Testac-1.mdb", SQL )
     'Set rsScore = GetMdbRecordset( "Testac-1.mdb", SQL )
     'Set rsScore = GetSecuredMdbRecordset( mdbFile, SQL, mdbPassword )
  Set rsScore = GetSecuredMdbRecordset("/Hmath/UsersPwd.mdb", SQL, "kj6688" )

  rsScore(Lesson) = Score
  rsScore.Update
%>
<BLOCKQUOTE>
學號：<%=No%><BR>
姓名：<%=Name%><BR>
  <BLOCKQUOTE>
  <TABLE Border=1>
  <TR BgColor=Yellow><TD>科目</TD><TD>答對題</TD><TD>總題數(分數)</TD></TR>
  <TR><TD>MATH</TD><TD Align=Right><%=ScoreHTML(rsScore("MATH"))%></TD><TD><%="/"&Count%></TD></TR>
  <TR><TD>ENGLS</TD><TD Align=Right><%=ScoreHTML(rsScore("ENGLS"))%></TD><TD><%="/"&Count%></TD></TR>
  <TR><TD>CHINE</TD><TD Align=Right><%=ScoreHTML(rsScore("CHINE"))%></TD><TD><%="/"&Count%></TD></TR>
  </TABLE>
  </BLOCKQUOTE>
</BLOCKQUOTE>
<HR>
<INPUT Type=Submit Value="  核對答案(繳 卷)  ">
  </FORM>
<!--<% 
   'Response.Redirect "Checkac-1t.asp?" & Request.QueryString 
  %>-->
 <!--<A HREF="Checkac-1.asp?No=<%=No%>&Name=<%=Name%>&Lesson=<%=Lesson%>">核對答案</A>
 <!--P><A HREF="EnterKac.asp?No=<%=No%>&Name=<%=Name%>">參加其他科目考試</A></P-->
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
















































































