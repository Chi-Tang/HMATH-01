
<%@ Import Namespace="System.Web.Mail" %> 

<!-- #include virtual="/kjasp/func/DB.fun" -->
<%
mdbFile = "UsersPwd.mdb"
mdbPassword = "kj6688"

Email = Request("Email")
If Email = Empty Then
   ShowMessage "Email �|����J!"
End If

' �ˬd���@ Email �O�_���ӽбb��
SQL = "Select * From Users Where Email='" & Email & "'"
Set rs = GetSecuredMdbRecordset( mdbFile, SQL, mdbPassword )
If rs.EOF Then
   ShowMessage "�藍�_�A���@ Email �å��ӽбb��!"
End If

' �e�X���@ Email ���|�����
''<%@ Import Namespace="System.Web.Mail" %> 

<Html>
<Body BgColor="White">
<H2>ASP.NET Email �o�e�{��!<Hr></H2>

<Form runat="server">
<Table Border=1>
  <Tr><Td>����̡G</Td>
  <Td><asp:TextBox id="mailTo" Size=40 runat="server"/></Td></Tr>
  <Tr><Td>�H��̡G</Td>
  <Td><asp:TextBox id="mailFrom" Size=40 runat="server"/></Td></Tr>
  <Tr><Td>�D���G</Td>
  <Td><asp:TextBox id="mailSubject" Size=40 runat="server"/></Td></Tr>
  <Tr><Td>����G</Td>
  <Td><asp:TextBox runat="server" id="mailBody" TextMode="MultiLine" Rows=8 Cols=60 />
  </Td></Tr>
</Table>
<asp:Button runat="server" Text="�e �X" OnClick="Button_Click" />
</Form>

<Hr><asp:Label id="Msg" runat="server" ForeColor="Red" /><p>
<Font Size=-1 Color=Blue>�ϥΥ��d�Ҥ��e�A�Х��Ѿ\�ѥ��uSMTP Server �P�l�󪺶ǰe�v�q�����������A�]�w�n SmtpMail.SmtpServer �ݩʡC</Font>
</Body>
</Html>

<script Language="VB" runat="server">

   Sub Button_Click(sender As Object, e As EventArgs) 
      Dim mail As MailMessage = New MailMessage

      mail.To         = mailTo.Text
      mail.From       = mailFrom.Text
      mail.Subject    = mailSubject.Text
      mail.BodyFormat = MailFormat.Text 
      mail.Body       = mailBody.Text

      On Error Resume Next
      SmtpMail.SmtpServer = "msa.hinet.net"
      SmtpMail.Send(mail) 

      If Err.Number <> 0 Then
         Msg.Text = Err.Description
      Else
         Msg.Text = "�l��w�g�e�X!"
      End If
   End Sub

</script>
  '<%@ Import Namespace="System.Web.Mail" %> 

<Html>
<Body BgColor="White">
<H2>ASP.NET Email �o�e�{��!<Hr></H2>


<Hr><asp:Label id="Msg" runat="server" ForeColor="Red" /><p>
<Font Size=-1 Color=Blue>�ϥΥ��d�Ҥ��e�A�Х��Ѿ\�ѥ��uSMTP Server �P�l�󪺶ǰe�v�q�����������A�]�w�n SmtpMail.SmtpServer �ݩʡC</Font>
</Body>
</Html>

<script Language="VB" runat="server">

  ' Sub Button_Click(sender As Object, e As EventArgs) 
      Dim mail As MailMessage = New MailMessage

      mail.To         = Email
      mail.From       = "tech1206@www.yahoo.com.tw"
      mail.Subject    =  "�z���|�����"
      mail.BodyFormat = MailFormat.Text 
       Body = "�ϥΪ̦W�١G " & rs("UserID") & vbCrLf
       Body = Body & "�@�@�@�K�X�G " & rs("Password") & vbCrLf 
       Body = Body & "�@�@�@�m�W�G " & rs("Name") & vbCrLf & vbCrLf
       Body = Body & "ASP�����s�@�Х� �q�W"
      'mail.Body = Body
     mail.Body       = Body

      On Error Resume Next
      SmtpMail.SmtpServer = "msa.hinet.net"
      SmtpMail.Send(mail) 

      If Err.Number <> 0 Then
         Msg.Text = Err.Description
      Else
         Msg.Text = "�l��w�g�e�X!"
      End If
  ' End Sub

</script>


'Set kjAx = Server.CreateObject("kjActiveX.Agent")
'Set Outlook = kjAx.CreateObject("Outlook.Application")
'Set mail = Outlook.CreateItem(0)
'mail.Subject = "�z���|�����"
'Body = "�ϥΪ̦W�١G " & rs("UserID") & vbCrLf
'Body = Body & "�@�@�@�K�X�G " & rs("Password") & vbCrLf 
'Body = Body & "�@�@�@�m�W�G " & rs("Name") & vbCrLf & vbCrLf
'Body = Body & "ASP�����s�@�Х� �q�W"
'mail.Body = Body
'mail.Recipients.Add Email
'mail.Save
'mail.Send

ShowMessage "�z����Ƥw�g�Ǩ� " & Email & " �F�I"
%>

<% 
Sub ShowMessage( msg ) 
%>
   <CENTER>
   <H2><%=msg%><HR></H2>
   <FORM><INPUT Type=Button Value="��^�W�@��" OnClick="history.back();">
   </FORM>
   </CENTER>
<%
   Response.End 
End Sub 
%>

