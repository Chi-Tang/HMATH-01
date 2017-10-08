<!-- #include virtual="/kjasp/func/DB.fun" -->

<%
TitleID = CLng(Request("TitleID"))

SQL = "Select * From Titles Where TitleID=" & TitleID
Set rs = GetMdbRecordset( "News.mdb", SQL )

If rs Is Nothing Or rs.EOF Then
   Response.Redirect "Title.asp?Msg=此一討論主題已被刪除!"
End If
%>

<HTML>
<BODY BACKGROUND="../../hmath/b01.jpg">
<H2 Align=Center>集思廣益 VB 討論群組 - 討論看板!</H2>

<!------------------------------>
<!-- 以下為「討論主題」的顯示 -->
<!------------------------------>
<TABLE Width="100%"><TR><TD Align=Center bgColor=#000080>
<FONT Color=#FFFFFF Size=+1><B>討 論 主 題</B></FONT></TD></TR></TABLE>
<%
   Words = Replace( "" & rs("Words"), Chr(13), "<BR>" )
   Email = "<A HREF=mailto:" & rs("Email") & ">" & rs("Email") & "</A>"
%>

<CENTER><TABLE>
<TR><TD><B>作者:</B><%=rs("Name")%><B>　Email:</B><%=Email%>
<B>　日期:</B><%=Month(rs("CreateDate"))%>/<%=Day(rs("CreateDate"))%>
</TD></TR>
<TR><TD><B>主題:</B><%=rs("Subject")%></TD></TR>
<TR><TD BgColor=Cyan><%=Words%></TD></TR>
</Table></CENTER>

<!------------------------------>
<!-- 以下為「討論看板」的顯示 -->
<!------------------------------>
<% 
SQL = "Select * From Details Where TitleID=" & TitleID 
SQL = SQL & " Order By DetailID Desc"
Set rsDetail = GetMdbRecordset( "News.mdb", SQL )

If Not rsDetail.EOF Then 
%>
<TABLE Width="100%"><TR><TD Align=Center bgColor=#000080>
<FONT Color=#FFFFFF Size=+1><B>討 論 看 板</B></FONT></TD></TR></TABLE>
<%
   While Not rsDetail.EOF
      Words = Replace( "" & rsDetail("Words"), Chr(13), "<BR>" )
      Email = "<A HREF=mailto:" & rsDetail("Email") & ">" & rsDetail("Email") & "</A>"
%>
      <CENTER><TABLE>
      <TR><TD NOWARP><B>作者:</B><%=rsDetail("Name")%><B>　Email:</B><%=Email%>
      <B>　日期:</B><%=Month(rsDetail("NewsDate"))%>/<%=Day(rsDetail("NewsDate"))%>
      </TD></TR>
      <TR><TD><B>主題:</B><%=rsDetail("Subject")%></TD></TR>
      <TR><TD BgColor=Cyan><%=Words%></B></TD></TR>
      </Table></CENTER><hr>
<%
      rsDetail.MoveNext
   Wend
End If
%>

<!-------------------------------->
<!-- 以下為「參與討論」輸入表單 -->
<!-------------------------------->
<TABLE Width="100%"><TR><TD Align=Center bgColor=#000080>
<FONT Color=#FFFFFF Size=+1><B>參 與 討 論</B></FONT></TD></TR></TABLE>

<FORM action=DetNew.asp method=POST>
<Input Type=Hidden Name=TitleID Value=<%=Request("TitleID")%>>
<CENTER><TABLE Border="0">

<TR><TD>姓名：</TD>
<TD><INPUT Type=Text Size="30" name="Name" Value=<%=Session("Name")%>></TD></TR>

<TR><TD>Email：</TD>
<TD><INPUT Type=Text Size="30" name="Email" Value=<%=Session("Email")%>></TD></TR>

<TR><TD>主題：</TD>
<TD><INPUT Type=Text Size="60" name="Subject"></TD></TR>

<TR><TD>內容：</TD>
<TD><TEXTAREA name="Words" rows="8" cols="60"></TEXTAREA></TD></TR>
</TABLE>
<INPUT Type="submit" value=" 送出討論內容 "></CENTER>
</Form>
<HR><FONT Color=Red><%=Request("Msg")%></FONT>
　　　<A HREF="Title.asp">返回討論群組主畫面</A>
</BODY>
</HTML>
