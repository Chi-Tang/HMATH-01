<!-- #include virtual="/kjasp/func/DB.fun" -->
<%
mdbFile = "UsersPwd.mdb"
mdbPassword = "kj6688"

UserID = Request("UserID")
Password = Request("Password")
Password2 = Request("Password2")
Name = Request("Name")
Email = Request("Email")

' �ˬd���O�_�d��
If UserID = Empty Or Password = Empty Or Name = Empty Then
   ShowMessage "���d�աA������!"
End If

' �ˬd Email
If InStr( Email, "@" ) = 0 Or Right(Email, 1) = "@" Or Left(Email, 1) = "@" Then
   ShowMessage "Email ���~�A�Э��s��J!"
End If

' �ˬd�⦸��J���K�X�O�_�@�P
If Password <> Password2 Then
   ShowMessage "�⦸��J���K�X���@�P!"
End If

' �ˬd���@ UserID �w���L�H����
SQL = "Select * From Users Where UserID='" & UserID & "'"
Set rs = GetSecuredMdbRecordset( mdbFile, SQL, mdbPassword )
If Not rs.EOF Then 	' ���@ UserID �X�w�Q����
   ShowMessage "���@�u�ϥΪ̦W�١v�w�Q���ΡA�Э��s��J!"
End If

' �ˬd���@ Email �O�_�w�ӽбb��
SQL = "Select * From Users Where Email='" & Email & "'"
Set rs = GetSecuredMdbRecordset( mdbFile, SQL, mdbPassword )
If Not rs.EOF Then 	' ���@ Email �w�ӽбb��
   ShowMessage "���@ Email �w�ӽбb���A�Э��s��J!"
End If

rs.AddNew
rs("UserID") = UserID
rs("Password") = Password
rs("Name") = Name
rs("Email") = Email
rs.Update

SQL = "Select * From ���Z�� " 
'SQL = SQL & "Where �Ǹ�=" & No & " And �m�W='" & Name & "'"
Set rsScore = GetMdbRecordset( "Testac-1.mdb", SQL )
 rsScore.AddNew
  'rsScore("�Ǹ�") = CLng(UserID)
 'rsScore("�Ǹ�") = CLng(Trim(Right(UserID,6)))
 rsScore("�Ǹ�") = Trim(Right(UserID,6))
 rsScore("�m�W") =  Name
 rsScore("MATH") = -1
 rsScore("ENGLS") = -1
 rsScore("CHINE") = -1
rsScore.Update

ShowMessage "�z�w�g�����������|���F�I"
%>

<% 
Sub ShowMessage( msg ) 
%>
   <CENTER>
   <H2><%=msg%><HR></H2>
   <FORM><INPUT Type=Button Value="��^�W�@��" OnClick="history.back();">
   
    <!--<A HREF="Enterkac-1.asp">���s�i�J����</A>--> 
   <A HREF="login-1t.asp">���s�i�J����</A> 

   </FORM>
   </CENTER>
<%
   Response.End 
End Sub 
%>


























