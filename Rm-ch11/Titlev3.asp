<!-- #include virtual="/kjasp/func/DB.fun" -->

<%
SQL = "Select * From Titles Where LastNewsDate > DateAdd(""d"",-14,Date()) Order By LastNewsDate Desc"
Set rs = GetMdbRecordset( "News.mdb", SQL )
If rs Is Nothing Then 
   Response.Write "呼叫 GetMdbRecordset 失敗!"
   Response.End
End If
%>

<HTML>
<BODY BACKGROUND="../../hmath/b01.jpg">
<H2 Align=Center>集思廣益 VB 討論群組 - 討論看板!<HR></H2>
<TABLE Border=0 Cellspacing=5>
<TR BgColor=#00FFFF>
<TD>日期</TD>
<TD NoWrapP>作  者</TD>
<TD>則數</TD>
<TD>主題</TD></TR>

<%
If Not rs.EOF Then rs.MoveFirst
While Not rs.EOF
   Day1   = Right("00" & Day(rs("CreateDate")), 2)
   Month1 = Right("00" & Month(rs("CreateDate")), 2)
   Day2   = Right("00" & Day(rs("LastNewsDate")), 2)
   Month2 = Right("00" & Month(rs("LastNewsDate")), 2)
   DateRange = Month1 & "/" & Day1 & "-" & Month2 & "/" & Day2
%>
   <TR Valign=TOP>
   <TD NoWrap><%=DateRange%></TD>
   <TD><%=rs("Name")%></TD>
   <TD Align=Right><%=rs("Number")%></TD>
   <TD><A HREF=Detail.asp?TitleID=<%=rs("TitleID")%>>
   <%=rs("Subject")%></A></TD>
   </TR>
<%
   rs.MoveNext
Wend  
%>
</TABLE>

<TABLE Width="100%"><TR><TD Align=Center bgcolor=#000080>
<FONT Color=#FFFFFF size=+1><B>發起討論主題</B></font>
</TD></TR></TABLE>
<CENTER>
<FORM Action=TitleNew.asp Method=POST>
<TABLE Border="0">

<TR><TD>姓名：</TD>
<TD><INPUT Type="text" Size="30" Name="Name" 
Value="<%=Session("Name")%>"></TD></TR>

<TR><TD>Email：</TD>
<TD><INPUT Type="text" Size="30" Name="Email" 
Value="<%=Session("Email")%>"></TD></TR>

<TR><TD>主題：</TD>
<TD><INPUT Type="text" Size="60" Name="Subject"
Value="<%=Session("Subject")%>"></TD></TR>

<TR><TD>內容：</TD>
<TD><TEXTAREA Name="Words" Rows="8" Cols="60">
<%=Session("Words")%></TEXTAREA></TD></TR>
</TABLE>
<INPUT Type="submit" Value="送出討論主題">
</FORM>
</CENTER><HR><FONT Color=Red><%=Request("Msg")%></FONT>
</BODY>
</HTML>
