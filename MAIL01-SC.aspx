<%@ Import Namespace="System.Web.Mail" %> 

<Html>
<Body BgColor="White">
<H2>ASP.NET Email 發送程式!<Hr></H2>

<Form runat="server">
<Table Border=1>
  <Tr><Td>收件者：</Td>
  <Td><asp:TextBox id="mailTo" Size=40 runat="server"/></Td></Tr>
  <Tr><Td>寄件者：</Td>
  <Td><asp:TextBox id="mailFrom" Size=40 runat="server"/></Td></Tr>
  <Tr><Td>主旨：</Td>
  <Td><asp:TextBox id="mailSubject" Size=40 runat="server"/></Td></Tr>
  <Tr><Td>內文：</Td>
  <Td><asp:TextBox runat="server" id="mailBody" TextMode="MultiLine" Rows=8 Cols=60 />
  </Td></Tr>
</Table>
<asp:Button runat="server" Text="送 出" OnClick="Button_Click" />
</Form>

<Hr><asp:Label id="Msg" runat="server" ForeColor="Red" /><p>
<Font Size=-1 Color=Blue>使用本範例之前，請先參閱書本「SMTP Server 與郵件的傳送」段落中的說明，設定好 SmtpMail.SmtpServer 屬性。</Font>
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
      ' 因為驗證方式沒設，請在SmtpMail.Send(mail) 這一行之前，加上一行:
    'mail.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate")=1
     
      'SmtpMail.Send(mail) 

      If Err.Number <> 0 Then
         Msg.Text = Err.Description
      Else
         Msg.Text = "郵件已經送出!"
      End If
   End Sub

</script>











