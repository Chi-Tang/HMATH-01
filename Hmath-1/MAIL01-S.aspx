<%@ Import Namespace="System.Web.Mail" %> 

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
       'Dim mail = server.createobject("cdo.message")  
       'Dim iconf= server.createobject("cdo.configuration")  
             
       mail.To         = mailTo.Text
      mail.From       = mailFrom.Text
      mail.Subject    = mailSubject.Text
      mail.BodyFormat = MailFormat.Text 
      mail.Body       = mailBody.Text
      
      On Error Resume Next
      SmtpMail.SmtpServer = "seed.net.tw"
      'SmtpMail.SmtpServer = "msa.hinet.net"
      ' �]�����Ҥ覡�S�]�A�ЦbSmtpMail.Send(mail) �o�@�椧�e�A�[�W�@��:
    'mail.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate")=1
     
      SmtpMail.Send(mail) 

      If Err.Number <> 0 Then
         Msg.Text = Err.Description
      Else
         Msg.Text = "�l��w�g�e�X!"
      End If
   End Sub

</script>
