<!-- #include virtual="/kjasp/func/DB.fun" -->

<%
TitleID = CLng(Request("TitleID"))

SQL = "Select * From Titles Where TitleID=" & TitleID
Set rs = GetMdbRecordset( "News.mdb", SQL )

If rs Is Nothing Or rs.EOF Then
   Response.Redirect "Title.asp?Msg=���@�Q�ץD�D�w�Q�R��!"
End If
%>

<HTML>
<BODY BACKGROUND="../../hmath/b01.jpg">
<H2 Align=Center>����s�q VB �Q�׸s�� - �Q�׬ݪO!</H2>

<!------------------------------>
<!-- �H�U���u�Q�ץD�D�v����� -->
<!------------------------------>
<TABLE Width="100%"><TR><TD Align=Center bgColor=#000080>
<FONT Color=#FFFFFF Size=+1><B>�Q �� �D �D</B></FONT></TD></TR></TABLE>
<%
   Words = Replace( "" & rs("Words"), Chr(13), "<BR>" )
   Email = "<A HREF=mailto:" & rs("Email") & ">" & rs("Email") & "</A>"
%>

<CENTER><TABLE>
<TR><TD><B>�@��:</B><%=rs("Name")%><B>�@Email:</B><%=Email%>
<B>�@���:</B><%=Month(rs("CreateDate"))%>/<%=Day(rs("CreateDate"))%>
</TD></TR>
<TR><TD><B>�D�D:</B><%=rs("Subject")%></TD></TR>
<TR><TD BgColor=Cyan><%=Words%></TD></TR>
</Table></CENTER>

<!------------------------------>
<!-- �H�U���u�Q�׬ݪO�v����� -->
<!------------------------------>
<% 
SQL = "Select * From Details Where TitleID=" & TitleID 
SQL = SQL & " Order By DetailID Desc"
Set rsDetail = GetMdbRecordset( "News.mdb", SQL )

If Not rsDetail.EOF Then 
%>
<TABLE Width="100%"><TR><TD Align=Center bgColor=#000080>
<FONT Color=#FFFFFF Size=+1><B>�Q �� �� �O</B></FONT></TD></TR></TABLE>
<%
   While Not rsDetail.EOF
      Words = Replace( "" & rsDetail("Words"), Chr(13), "<BR>" )
      Email = "<A HREF=mailto:" & rsDetail("Email") & ">" & rsDetail("Email") & "</A>"
%>
      <CENTER><TABLE>
      <TR><TD NOWARP><B>�@��:</B><%=rsDetail("Name")%><B>�@Email:</B><%=Email%>
      <B>�@���:</B><%=Month(rsDetail("NewsDate"))%>/<%=Day(rsDetail("NewsDate"))%>
      </TD></TR>
      <TR><TD><B>�D�D:</B><%=rsDetail("Subject")%></TD></TR>
      <TR><TD BgColor=Cyan><%=Words%></B></TD></TR>
      </Table></CENTER><hr>
<%
      rsDetail.MoveNext
   Wend
End If
%>

<!-------------------------------->
<!-- �H�U���u�ѻP�Q�סv��J��� -->
<!-------------------------------->
<TABLE Width="100%"><TR><TD Align=Center bgColor=#000080>
<FONT Color=#FFFFFF Size=+1><B>�� �P �Q ��</B></FONT></TD></TR></TABLE>

<FORM action=DetNew.asp method=POST>
<Input Type=Hidden Name=TitleID Value=<%=Request("TitleID")%>>
<CENTER><TABLE Border="0">

<TR><TD>�m�W�G</TD>
<TD><INPUT Type=Text Size="30" name="Name" Value=<%=Session("Name")%>></TD></TR>

<TR><TD>Email�G</TD>
<TD><INPUT Type=Text Size="30" name="Email" Value=<%=Session("Email")%>></TD></TR>

<TR><TD>�D�D�G</TD>
<TD><INPUT Type=Text Size="60" name="Subject"></TD></TR>

<TR><TD>���e�G</TD>
<TD><TEXTAREA name="Words" rows="8" cols="60"></TEXTAREA></TD></TR>
</TABLE>
<INPUT Type="submit" value=" �e�X�Q�פ��e "></CENTER>
</Form>
<HR><FONT Color=Red><%=Request("Msg")%></FONT>
�@�@�@<A HREF="Title.asp">��^�Q�׸s�եD�e��</A>
</BODY>
</HTML>
