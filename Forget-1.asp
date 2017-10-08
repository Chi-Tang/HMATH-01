
<%@ Import Namespace="System.Web.Mail" %> 

<!-- #include virtual="/kjasp/func/DB.fun" -->
<%
mdbFile = "UsersPwd.mdb"
mdbPassword = "kj6688"

Email = Request("Email")
If Email = Empty Then
   ShowMessage "Email 尚未輸入!"
End If

' 檢查此一 Email 是否有申請帳號
SQL = "Select * From Users Where Email='" & Email & "'"
Set rs = GetSecuredMdbRecordset( mdbFile, SQL, mdbPassword )
If rs.EOF Then
   ShowMessage "對不起，此一 Email 並未申請帳號!"
End If

' 送出此一 Email 的會員資料
''<%@ Import Namespace="System.Web.Mail" %> 

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
         Msg.Text = "郵件已經送出!"
      End If
   End Sub

</script>
  '<%@ Import Namespace="System.Web.Mail" %> 

<Html>
<Body BgColor="White">
<H2>ASP.NET Email 發送程式!<Hr></H2>


<Hr><asp:Label id="Msg" runat="server" ForeColor="Red" /><p>
<Font Size=-1 Color=Blue>使用本範例之前，請先參閱書本「SMTP Server 與郵件的傳送」段落中的說明，設定好 SmtpMail.SmtpServer 屬性。</Font>
</Body>
</Html>

<script Language="VB" runat="server">

  ' Sub Button_Click(sender As Object, e As EventArgs) 
      Dim mail As MailMessage = New MailMessage

      mail.To         = Email
      mail.From       = "tech1206@www.yahoo.com.tw"
      mail.Subject    =  "您的會員資料"
      mail.BodyFormat = MailFormat.Text 
       Body = "使用者名稱： " & rs("UserID") & vbCrLf
       Body = Body & "　　　密碼： " & rs("Password") & vbCrLf 
       Body = Body & "　　　姓名： " & rs("Name") & vbCrLf & vbCrLf
       Body = Body & "ASP網頁製作教本 敬上"
      'mail.Body = Body
     mail.Body       = Body

      On Error Resume Next
      SmtpMail.SmtpServer = "msa.hinet.net"
      SmtpMail.Send(mail) 

      If Err.Number <> 0 Then
         Msg.Text = Err.Description
      Else
         Msg.Text = "郵件已經送出!"
      End If
  ' End Sub

</script>


'Set kjAx = Server.CreateObject("kjActiveX.Agent")
'Set Outlook = kjAx.CreateObject("Outlook.Application")
'Set mail = Outlook.CreateItem(0)
'mail.Subject = "您的會員資料"
'Body = "使用者名稱： " & rs("UserID") & vbCrLf
'Body = Body & "　　　密碼： " & rs("Password") & vbCrLf 
'Body = Body & "　　　姓名： " & rs("Name") & vbCrLf & vbCrLf
'Body = Body & "ASP網頁製作教本 敬上"
'mail.Body = Body
'mail.Recipients.Add Email
'mail.Save
'mail.Send

ShowMessage "您的資料已經傳到 " & Email & " 了！"
%>

<% 
Sub ShowMessage( msg ) 
%>
   <CENTER>
   <H2><%=msg%><HR></H2>
   <FORM><INPUT Type=Button Value="返回上一頁" OnClick="history.back();">
   </FORM>
   </CENTER>
<%
   Response.End 
End Sub 
%>

