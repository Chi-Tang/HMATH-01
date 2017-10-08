  <!--#include virtual="/kjasp/func/Db.fun" -->
<HTML>
<HEAD><TITLE>訪客留言板</TITLE></HEAD>
<BODY TEXT="#000000" BGCOLOR="#FFFFFF" BACKGROUND="../../hmath/b01.jpg">
<H2 ALIGN=CENTER>訪 客 留 言 板<HR WIDTH="100%"></H2>

<%
SQL = "Select * From GuestBook Order By 時間 DESC"
Set rs = GetMdbStaticRecordset( "Gbook.mdb", SQL )

Page = CLng(Request("Page"))   ' CLng 不可省略
If Page < 1 Then Page = 1
If Page > rs.PageCount Then Page = rs.PageCount

ShowOnePage rs, Page
%>
<DIV ALIGN=right><P>
<FORM Action=gbook.asp Method=GET>
<A HREF="gform.htm">返回留言表單</A>　
<% 
   If Page <> 1 Then
      Response.Write "<A HREF=gbook.asp?Page=1>第一頁</A>　"
      Response.Write "<A HREF=gbook.asp?Page=" & (Page-1) & ">上一頁</A>　"
   End If
   If Page <> rs.PageCount Then
      Response.Write "<A HREF=gbook.asp?Page=" & (Page+1) & ">下一頁</A>　"
      Response.Write "<A HREF=gbook.asp?Page=" & rs.PageCount & ">最後一頁</A>　"
   End If
%>
頁次:<FONT COLOR="Red"><%=Page%>/<%=rs.PageCount%></FONT>
<P>
</BODY>
</HTML>

<!-- 以下為 ShowOnePage 及 RsToGbook 副程式 -->
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
  memo=Replace(rs("留言"), vbCrLf, "<BR>" )
  If Len(rs("圖示")) <> 0 Then
     PicHTML = "<IMG SRC=" & rs("圖示") & ">"
  Else
     PicHTML = ""
  End If
%>
<TABLE BOBDER=0  WIDTH="100%">
   <TR><TD WIDTH=110><%=PicHTML%></TD>
   <TD>
      <TABLE BORDER=0>
      <TR>
         <TD>作 者:<%=rs("姓名")%></TD>
         <TD>Email:<A HREF="mailto:<%=rs("Email")%>">
         <%=rs("Email")%></TD>
      </TR>
      <TR>
         <TD COLSPAN=2>主 題:<%=rs("主題")%></TD>
      </TR>
      <TR>
         <TD COLSPAN=2 BGCOLOR=#00FFFF><%=memo%></TD>
      </TR>
      <TR>
         <TD COLSPAN=2>時 間:<%=rs("時間")%></TD>
      </TR>
      </TABLE>
   </TD></TR>
</TABLE>
<HR>
<%
End Sub
%>







