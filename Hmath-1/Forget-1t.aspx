<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.OleDb" %>
<%@ Import Namespace="System.Web.Mail" %>

<html>
<body background="../B01.jpg" bgcolor="#FFFFFF">
<Form runat="server">
<h2>忘了帳號或密碼？<hr></h2>
<blockquote>
Email：<asp:TextBox runat="server" size="40" id="Email" />
   <asp:RequiredFieldValidator runat="server" Text="(必要欄位)"
        ControlToValidate="Email" EnableClientScript="False"
        Display="Dynamic" />
   <asp:RegularExpressionValidator runat="server"
        ControlToValidate="Email" Text="(Email 應含有 @ 符號)"
        ValidationExpression=".{1,}@.{3,}" 
        EnableClientScript="False" Display="Dynamic"/><Br>
<asp:Button runat="server" Text="請傳給我帳號及密碼" OnClick="Forget_Click" /><br>
<Font Size=-1 Color=Blue>請輸入您的 Email，然後按下「請傳給我帳號及密碼」</Font>
</blockquote>
<HR>
<asp:Label runat="server" id="Msg" ForeColor="Red" />
</Form>
</body>
</html>

<script Language="VB" runat="server">

   Sub Forget_Click(sender As Object, e As EventArgs)
      Msg.Text = ""
      If IsValid Then
         QueryDataAndSendTo()
      End If       
   End Sub

   Sub QueryDataAndSendTo()
      Dim Conn As OleDbConnection, Cmd As OleDbCommand
      Dim Rd As OleDbDataReader, SQL As String, Body As String

      Dim Provider = "Provider=Microsoft.Jet.OLEDB.4.0"
      
       'Dim Database = "Data Source=" & Server.MapPath( "../ch15/Users.mdb" )
      Dim Database = "Data Source=" & Server.MapPath( "UsersPwd.mdb" )
      Dim Dbpass = "Jet OLEDB:Database Password=kj6688"
      Conn = New OleDbConnection( Provider & ";" & DataBase & ";"&Dbpass )
      Conn.Open()

      ' 檢查 Email 是否存在
      SQL = "Select * From Users Where Email='" & Email.Text & "'"
      Cmd = New OleDbCommand( SQL, Conn )
      Rd = Cmd.ExecuteReader()
      If Rd.Read() Then ' 表示此一 Email 存在
         Dim mail As MailMessage = New MailMessage

         mail.Subject = "您的會員資料"
         mail.To = Rd.Item("Email")
         mail.From = "service@xxx.com.tw"   ' 改成系統維護者的 e-mail
         mail.BodyFormat = MailFormat.Text 
         Body = "使用者名稱：" & Rd.Item("UserID") & vbCrLf
         Body = Body & "　　　密碼：" & Rd.Item("Password") & vbCrLf 
         Body = Body & "　　　姓名：" & Rd.Item("Name") & vbCrLf & vbCrLf
         Body = Body & "ASP.NET 網頁製作教本 敬上"
         mail.Body = Body

         On Error Resume Next
         SmtpMail.SmtpServer = "msa.hinet.net"
          'SmtpMail.SmtpServer = "ms57.hinet.net"
         'SmtpMail.SmtpServer = "seed.net.tw"
          SmtpMail.Send(mail) 

         If Err.Number <> 0 Then
            Msg.Text = Err.Description & "link or change SmtpServer"
         Else
            Msg.Text = "帳號及密碼已經送出，請檢查您的信箱!"
         End If
      Else
         Msg.Text = "此一 Email 並未申請帳號!"
      End If  
      Conn.Close()
   End Sub

</script>