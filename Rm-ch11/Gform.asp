<!-- #include virtual="/kjasp/func/DB.fun" -->
<%
' Ū���U��쪺���
Name = Request("Name")
Email = Request("Email")
Subject = Request("Subject")
Memo = Request("Memo")
Pic = Request("Pic")

' �ˬd�U���O�_��J�����
If Name = "tang1206"  Then
   'ShowMessage "���d�դ�����!"
   Response.Redirect "gbook.asp"
End If 

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

'' �N�����ɦܡu�[�ݯd���v������ gbook.asp
  ''Response.Redirect "gbook.asp"
  ' Response.Redirect "Gform.htm"
 ShowMessage "�z�w�g�d�����\,����!"

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













