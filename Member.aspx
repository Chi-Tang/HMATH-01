<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.OleDb" %>

<html>
<body background="../B01.jpg" bgcolor="#FFFFFF">
<Form runat="server">
<h2>�|���n��<hr></h2>

<blockquote>
<table border="0">
  <tr>
    <td align="right">�ϥΪ̦W�١G</td>
    <td><asp:TextBox runat="server" size="20" id="UserID" />
        <asp:RequiredFieldValidator runat="server" Text="(���n���)"
        ControlToValidate="UserID" EnableClientScript="False"
        Display="Dynamic" />
    </td>
  </tr>
  <tr>
    <td align="right">�K�X�G</td>
    <td><asp:TextBox runat="server" size="20" id="Password1" TextMode="Password" />
        <asp:RequiredFieldValidator runat="server" Text="(���n���)"
        ControlToValidate="Password1" EnableClientScript="False"
        Display="Dynamic" />
    </td>
  </tr>
  <tr>
    <td align="right">�K�X�A�T�{�G</td>
    <td><asp:TextBox runat="server" size="20" id="Password2" TextMode="Password" />
        <asp:RequiredFieldValidator runat="server" Text="(���n���)"
        ControlToValidate="Password2" EnableClientScript="False"
        Display="Dynamic" />
    </td>
  </tr>
  <tr>
    <td align="right">�m�W�G</td>
    <td><asp:TextBox runat="server" size="20" id="Name" />
        <asp:RequiredFieldValidator runat="server" Text="(���n���)"
        ControlToValidate="Name" EnableClientScript="False"
        Display="Dynamic" />
    </td>
  </tr>
  <tr>
    <td align="right">Email�G</td>
    <td><asp:TextBox runat="server" size="40" id="Email" />
        <asp:RequiredFieldValidator runat="server" Text="(���n���)"
        ControlToValidate="Email" EnableClientScript="False"
        Display="Dynamic" />
        <asp:RegularExpressionValidator runat="server"
        ControlToValidate="Email" Text="(Email ���t�� @ �Ÿ�)"
        ValidationExpression=".{1,}@.{3,}" 
        EnableClientScript="False" Display="Dynamic"/>
    </td>
  </tr>
</table>
<p><asp:Button runat="server" Text=" �[ �J " OnClick="Join_Click" /></p>
</blockquote>
<A HREF="Forget.aspx">�·|���A���O�ѤF�K�X</A>
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
            Msg.Text = "�⦸��J���K�X���@��!"
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

      ' �ˬd UserID �O�_����
      SQL = "Select * From Users Where UserID='" & UserID.Text & "'"
      Cmd = New OleDbCommand( SQL, Conn )
      Rd = Cmd.ExecuteReader()
      If Rd.Read() Then ' ��ܦ��@ UserID �w�g�s�b
         Msg.Text = "�z�ҿ�J�u�ϥΪ̦W�١v�w�g�Q����, �Э��s��J!"
         Conn.Close()
         Exit Sub
      End If  
      Rd.Close()

      ' �ˬd Email �O�_����
      SQL = "Select * From Users Where Email='" & Email.Text & "'"
      Cmd = New OleDbCommand( SQL, Conn )
      Rd = Cmd.ExecuteReader()
      If Rd.Read() Then ' ��ܦ��@ Email �w�g�s�b
         Msg.Text = "���@ Email �w�ӽйL�b��!"
         Conn.Close()
         Exit Sub
      End If  
      Rd.Close()
      
      ' �g�J���
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
      Msg.Text = "��ƳB�z�����A�z�w�g�����������|��!"

      Conn.Close()
   End Sub

</script>