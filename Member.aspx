<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.OleDb" %>

<html>
<body background="../B01.jpg" bgcolor="#FFFFFF">
<Form runat="server">
<h2>會員登錄<hr></h2>

<blockquote>
<table border="0">
  <tr>
    <td align="right">使用者名稱：</td>
    <td><asp:TextBox runat="server" size="20" id="UserID" />
        <asp:RequiredFieldValidator runat="server" Text="(必要欄位)"
        ControlToValidate="UserID" EnableClientScript="False"
        Display="Dynamic" />
    </td>
  </tr>
  <tr>
    <td align="right">密碼：</td>
    <td><asp:TextBox runat="server" size="20" id="Password1" TextMode="Password" />
        <asp:RequiredFieldValidator runat="server" Text="(必要欄位)"
        ControlToValidate="Password1" EnableClientScript="False"
        Display="Dynamic" />
    </td>
  </tr>
  <tr>
    <td align="right">密碼再確認：</td>
    <td><asp:TextBox runat="server" size="20" id="Password2" TextMode="Password" />
        <asp:RequiredFieldValidator runat="server" Text="(必要欄位)"
        ControlToValidate="Password2" EnableClientScript="False"
        Display="Dynamic" />
    </td>
  </tr>
  <tr>
    <td align="right">姓名：</td>
    <td><asp:TextBox runat="server" size="20" id="Name" />
        <asp:RequiredFieldValidator runat="server" Text="(必要欄位)"
        ControlToValidate="Name" EnableClientScript="False"
        Display="Dynamic" />
    </td>
  </tr>
  <tr>
    <td align="right">Email：</td>
    <td><asp:TextBox runat="server" size="40" id="Email" />
        <asp:RequiredFieldValidator runat="server" Text="(必要欄位)"
        ControlToValidate="Email" EnableClientScript="False"
        Display="Dynamic" />
        <asp:RegularExpressionValidator runat="server"
        ControlToValidate="Email" Text="(Email 應含有 @ 符號)"
        ValidationExpression=".{1,}@.{3,}" 
        EnableClientScript="False" Display="Dynamic"/>
    </td>
  </tr>
</table>
<p><asp:Button runat="server" Text=" 加 入 " OnClick="Join_Click" /></p>
</blockquote>
<A HREF="Forget.aspx">舊會員，但是忘了密碼</A>
<hr>
<asp:Label runat="server" id="Msg" ForeColor="Red" />
</Form>
</body>
</html>

<script Language="VB" runat="server">

   Sub Join_Click(sender As Object, e As EventArgs)
      Msg.Text = ""
      If IsValid Then
         If Password1.Text <> Password2.Text Then
            Msg.Text = "兩次輸入的密碼不一樣!"
         Else
            WriteDataToDatabase()
         End If
      End If       
   End Sub

   Sub WriteDataToDatabase()
      Dim Conn As OleDbConnection, Cmd As OleDbCommand
      Dim Rd As OleDbDataReader, SQL As String

      Dim Provider = "Provider=Microsoft.Jet.OLEDB.4.0"
      ' Dim Database = "Data Source=" & Server.MapPath( "../ch15/Users.mdb" )
      Dim Database = "Data Source=" & Server.MapPath( "Users.mdb" )
      Conn = New OleDbConnection( Provider & ";" & DataBase )
      Conn.Open()

      ' 檢查 UserID 是否重複
      SQL = "Select * From Users Where UserID='" & UserID.Text & "'"
      Cmd = New OleDbCommand( SQL, Conn )
      Rd = Cmd.ExecuteReader()
      If Rd.Read() Then ' 表示此一 UserID 已經存在
         Msg.Text = "您所輸入「使用者名稱」已經被佔用, 請重新輸入!"
         Conn.Close()
         Exit Sub
      End If  
      Rd.Close()

      ' 檢查 Email 是否重複
      SQL = "Select * From Users Where Email='" & Email.Text & "'"
      Cmd = New OleDbCommand( SQL, Conn )
      Rd = Cmd.ExecuteReader()
      If Rd.Read() Then ' 表示此一 Email 已經存在
         Msg.Text = "此一 Email 已申請過帳號!"
         Conn.Close()
         Exit Sub
      End If  
      Rd.Close()
      
      ' 寫入資料
      SQL = "INSERT INTO Users (UserID, [Password], Name, Email) VALUES (?, ?, ?, ?)"
      Cmd = New OleDbCommand( SQL, Conn )
      Cmd.Parameters.Add( New OleDbParameter("@UserID", OleDbType.Char, 255))
      Cmd.Parameters.Add( New OleDbParameter("@Password", OleDbType.Char, 255))
      Cmd.Parameters.Add( New OleDbParameter("@Name", OleDbType.Char, 255))
      Cmd.Parameters.Add( New OleDbParameter("@Email", OleDbType.Char, 255))
      Cmd.Parameters(0).Value = UserID.Text
      Cmd.Parameters(1).Value = Password1.Text
      Cmd.Parameters(2).Value = Name.Text
      Cmd.Parameters(3).Value = Email.Text
      Cmd.ExecuteNonQuery()
      Msg.Text = "資料處理完畢，您已經成為本站的會員!"

      Conn.Close()
   End Sub

</script>