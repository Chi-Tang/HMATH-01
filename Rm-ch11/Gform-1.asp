<!-- #include virtual="/kjasp/func/DB.fun" -->
<%
' Ū���U��쪺���
Name = Request("Name")
Email = Request("Email")
Subject = Request("Subject")
Memo = Request("Memo")
Pic = Request("Pic")

' �ˬd�U���O�_��J�����
If Name = "" Or Email = "" Or Subject = "" Or Memo = "" Then
   ShowMessage "���d�դ�����!"
   Response.End
End If

'Set rs = GetMdbRecordset( "gbook.mdb", "GuestBook" )
SQL = "Select * From GuestBook"
Set rs = GetMdbRecordset( "Gbook.mdb", SQL )

rs.AddNew
rs("�m�W") = Name
rs("Email") = Email
rs("�D�D") = Subject
rs("�d��") = Memo
If Pic <> "No" Then rs("�ϥ�") = Pic
rs.Update

' �N�����ɦܡu�[�ݯd���v������ gbook.asp
Response.Redirect "gbook.asp"
%>

<% 
Sub ShowMessage( msg ) 
%>
<CENTER>
<H2><%=msg%><HR></H2>
<FORM><INPUT Type=Button Value="��^�W�@��" OnClick="history.back();"></FORM>
</CENTER>
<% 
End Sub 
%>
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
<%
'' Sub ShowMessage( msg )
  Sub ShowMessage( msg,yn )
 
%>

   <CENTER>
   <H2><%=msg%><HR></H2>
    
   <FORM action="forget-1r.aspx" method= "get"  >
  <%
     If yn = "y" Then
  '' Response.Redirect "login-1r.asp?" & Request.QueryString

   %>
       <p><input type="submit" value="�жǵ��ڨϥΪ̱K�X "></p>
        <p>�A��J�b���G<input type="text" size="20" name="No" Value="<%=No%>"></p>
    <%
      End If
    %>
      
     <INPUT Type=Button Value="��^�W�@��" OnClick="history.back();">
        <!--<A HREF="Enterkac-1.asp">���s�i�J����</A>
     <A HREF="login-1t.asp">���s�i�J����[����update���\��form action]</A>--> 

   </FORM>
   </CENTER>
  
<%
   Response.End 
End Sub 
%>







