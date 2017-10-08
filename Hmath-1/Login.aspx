<%@ Import Namespace="System.Web.Security " %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.OleDb" %>

<Html>
<Body>
<Form runat="server">
<H3>請先登入您的帳號及密碼:<HR></H3>
<Blockquote>

帳號:<asp:TextBox runat="server" id="Account" /><p>
密碼:<asp:TextBox runat="server" id="Password" 
                  TextMode="Password" /><p>
記得我:<asp:CheckBox runat="server" id="RememberMe" /><p>

<asp:Button runat="server" text="登入" OnClick="Login_Click" /><p>
<asp:Label runat="server" id="Msg" ForeColor="Red" /><p>

<A HREF="Member.aspx">加入會員</A> 

</Blockquote><HR>
<Font Size=-1 Color=Blue>登入網頁前, 請確定瀏覽器的 Cookie 是開啟的.</Font>
</Form>
</Body>
</Html>

<script language="VB" runat="server">

   Sub Login_Click( sender As Object, e As EventArgs )

      If Verify( Account.Text, Password.Text ) Then
        
       ' FormsAuthentication.RedirectFromLoginPage(Account.Text, RememberMe.Checked)
       ' EXIT Sub  
       Response.Redirect("Enterkac.asp")
      Else
         Msg.Text = "帳號或密碼錯誤, 請重新輸入!"

      End If
   End Sub

   Function Verify( 帳號 As String, 密碼 As String) As Boolean
      Dim Conn As OleDbConnection, Cmd As OleDbCommand
      Dim Rd As OleDbDataReader, SQL  As String

      Dim Provider = "Provider=Microsoft.Jet.OLEDB.4.0"
      Dim Database = "Data Source=" & Server.MapPath( "Users.mdb" )
      Conn = New OleDbConnection( Provider & ";" & DataBase )
      Conn.Open()

      SQL = "Select * From Users Where " & _
                    "UserID='" & 帳號 & "'" & _
                    " And Password='" & 密碼 & "'"
      Cmd = New OleDbCommand( SQL, Conn )
      Rd = Cmd.ExecuteReader()
      If Rd.Read() Then ' 表示有找到 UserID 及 Password, 通過驗證
         Conn.Close()
         Return True
      Else
         Conn.Close()
         Msg.Text = "帳號或密碼錯誤, 請重新輸入!"
         Return False
      End If
   End Function

</script>