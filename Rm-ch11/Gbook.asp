  <!--#include virtual="/kjasp/func/Db.fun" -->
<HTML>
<HEAD><TITLE>�X�ȯd���O</TITLE></HEAD>
<BODY TEXT="#000000" BGCOLOR="#FFFFFF" BACKGROUND="../../hmath/b01.jpg">
<H2 ALIGN=CENTER>�X �� �d �� �O<HR WIDTH="100%"></H2>

<%
SQL = "Select * From GuestBook Order By �ɶ� DESC"
Set rs = GetMdbStaticRecordset( "Gbook.mdb", SQL )

Page = CLng(Request("Page"))   ' CLng ���i�ٲ�
If Page < 1 Then Page = 1
If Page > rs.PageCount Then Page = rs.PageCount

ShowOnePage rs, Page
%>
<DIV ALIGN=right><P>
<FORM Action=gbook.asp Method=GET>
<A HREF="gform.htm">��^�d�����</A>�@
<% 
   If Page <> 1 Then
      Response.Write "<A HREF=gbook.asp?Page=1>�Ĥ@��</A>�@"
      Response.Write "<A HREF=gbook.asp?Page=" & (Page-1) & ">�W�@��</A>�@"
   End If
   If Page <> rs.PageCount Then
      Response.Write "<A HREF=gbook.asp?Page=" & (Page+1) & ">�U�@��</A>�@"
      Response.Write "<A HREF=gbook.asp?Page=" & rs.PageCount & ">�̫�@��</A>�@"
   End If
%>
����:<FONT COLOR="Red"><%=Page%>/<%=rs.PageCount%></FONT>
<P>
</BODY>
</HTML>

<!-- �H�U�� ShowOnePage �� RsToGbook �Ƶ{�� -->
<% 
Sub ShowOnePage( rs, Page )
  rs.AbsolutePage = Page
  For iPage = 1 To rs.PageSize
     RsToGbook rs
     rs.MoveNext
     If rs.EOF Then Exit For
  Next
End Sub

Sub RsToGbook( rs )
  memo=Replace(rs("�d��"), vbCrLf, "<BR>" )
  If Len(rs("�ϥ�")) <> 0 Then
     PicHTML = "<IMG SRC=" & rs("�ϥ�") & ">"
  Else
     PicHTML = ""
  End If
%>
<TABLE BOBDER=0  WIDTH="100%">
   <TR><TD WIDTH=110><%=PicHTML%></TD>
   <TD>
      <TABLE BORDER=0>
      <TR>
         <TD>�@ ��:<%=rs("�m�W")%></TD>
         <TD>Email:<A HREF="mailto:<%=rs("Email")%>">
         <%=rs("Email")%></TD>
      </TR>
      <TR>
         <TD COLSPAN=2>�D �D:<%=rs("�D�D")%></TD>
      </TR>
      <TR>
         <TD COLSPAN=2 BGCOLOR=#00FFFF><%=memo%></TD>
      </TR>
      <TR>
         <TD COLSPAN=2>�� ��:<%=rs("�ɶ�")%></TD>
      </TR>
      </TABLE>
   </TD></TR>
</TABLE>
<HR>
<%
End Sub
%>







