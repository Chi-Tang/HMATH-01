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
    'Imports System.Web.Mail 
    'Pubic Sub SendMailMessage() 
    'Dim _mm As new MailMessage() 
   Sub Button_Click(sender As Object, e As EventArgs) 
       Dim _mm As new MailMessage() 
        'Dim _mm = server.getobject("cdo.message")
        ''Dim _mm = server.createobject("cdo.message")
        Dim iconf = server.createobject("cdo.configuration")
       
       Dim flds =  iconf.fields
 with flds
  .Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate")=1
   .Update()
 End with
' Build HTML for message body. 
Dim strHTML As String = "<HTML><HEAD><BODY>Test</BODY></HTML>" 

' Apply the settings to the message. 
With _mm 
 ''.Configuration =iconf

.To = mailTo.Text
.From = mailFrom.Text 
.Subject = mailSubject.Text
.BodyFormat = MailFormat.Text  
.Body = mailBody.Text 
 
End With 
 On Error Resume Next

  'SmtpMail.SmtpServer = "seed.net.tw"

SmtpMail.Send(_mm) 
      
      'SmtpMail.SmtpServer = "msa.hinet.net"
      ' �]�����Ҥ覡�S�]�A�ЦbSmtpMail.Send(mail) �o�@�椧�e�A�[�W�@��:
    'mail.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate")=1
     
      'SmtpMail.Send(mail) 

      If Err.Number <> 0 Then
         Msg.Text = Err.Description
      Else
         Msg.Text = "�l��w�g�e�X!"
      End If
   End Sub

</script>











