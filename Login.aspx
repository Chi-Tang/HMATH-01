<%@ Import Namespace="System.Web.Security " %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.OleDb" %>

<Html>
<Body>
<Form runat="server">
<H3>�Х��n�J�z���b���αK�X:<HR></H3>
<Blockquote>

�b��:<asp:TextBox runat="server" id="Account" /><p>
�K�X:<asp:TextBox runat="server" id="Password" 
                  TextMode="Password" /><p>
�O�o��:<asp:CheckBox runat="server" id="RememberMe" /><p>

<asp:Button runat="server" text="�n�J" OnClick="Login_Click" /><p>
<asp:Label runat="server" id="Msg" ForeColor="Red" /><p>

<A HREF="Member.aspx">�[�J�|��</A> 

</Blockquote><HR>
<Font Size=-1 Color=Blue>�n�J�����e, �нT�w�s������ Cookie �O�}�Ҫ�.</Font>
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
         Msg.Text = "�b���αK�X���~, �Э��s��J!"

      End If
   End Sub

   Function Verify( �b�� As String, �K�X As String) As Boolean
      Dim Conn As OleDbConnection, Cmd As OleDbCommand
      Dim Rd As OleDbDataReader, SQL  As String

      Dim Provider = "Provider=Microsoft.Jet.OLEDB.4.0"
      Dim Database = "Data Source=" & Server.MapPath( "Users.mdb" )
      Conn = New OleDbConnection( Provider & ";" & DataBase )
      Conn.Open()

      SQL = "Select * From Users Where " & _
                    "UserID='" & �b�� & "'" & _
                    " And Password='" & �K�X & "'"
      Cmd = New OleDbCommand( SQL, Conn )
      Rd = Cmd.ExecuteReader()
      If Rd.Read() Then ' ��ܦ���� UserID �� Password, �q�L����
         Conn.Close()
         Return True
      Else
         Conn.Close()
         Msg.Text = "�b���αK�X���~, �Э��s��J!"
         Return False
      End If
   End Function

</script>