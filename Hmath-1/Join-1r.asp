<!-- #include virtual="/kjasp/func/DB.fun" -->
<%
mdbFile = "/Hmath/UsersPwd.mdb"
 'mdbFile = "UsersPwd.mdb"
mdbPassword = "kj6688"

UserID = Request("UserID")
 'Password = Request("Password")
 'Password2 = Request("Password2")
Name = Request("Name")
Email = Request("Email")
Adress = Request("Adress")

'' �ˬd���O�_�d��
'If UserID = Empty Or Password = Empty Or Name = Empty Or Adress = Empty Then
If UserID = Empty Or Name = Empty Or Adress = Empty Then
  ShowMessage "���d�աA������!","n"
End If

' �ˬd Email
If InStr( Email, "@" ) = 0 Or Right(Email, 1) = "@" Or Left(Email, 1) = "@" Then
   ShowMessage "Email ���~�A�Э��s��J!","n"
End If

'' �ˬd�⦸��J���K�X�O�_�@�P
 'If Password <> Password2 Then
 '  ShowMessage "�⦸��J���K�X���@�P!","n"
 'End If

' �ˬd���@ UserID �w���L�H����
SQL = "Select * From Users Where UserID='" & UserID & "'"
Set rs = GetSecuredMdbRecordset( mdbFile, SQL, mdbPassword )
If Not rs.EOF Then 	' ���@ UserID �X�w�Q����
   ShowMessage "���@�u�ϥΪ̦W�١v�w�Q���ΡA�Э��s��J!","n"
End If

' �ˬd���@ Email �O�_�w�ӽбb��
SQL = "Select * From Users Where Email='" & Email & "'"
Set rs = GetSecuredMdbRecordset( mdbFile, SQL, mdbPassword )
If Not rs.EOF Then 	' ���@ Email �w�ӽбb��
   ShowMessage "���@ Email �w�ӽбb���A�Э��s��J!","n"
End If

rs.AddNew
rs("UserID") = UserID
'rs("Password") = Password
  Randomize
   Rn=Fix(Rnd*100)
  rs("Password") = Trim(Left(UserID,3))+CStr(Rn)+Trim(Left(Email, 2))
rs("Name") = Name
rs("Email") = Email
rs("Adress") = Adress
rs.Update

SQL = "Select * From ���Z�� " 
'SQL = SQL & "Where �Ǹ�=" & No & " And �m�W='" & Name & "'"
'Set rsScore = GetMdbRecordset( "Testac-1t.mdb", SQL )
Set rsScore = GetSecuredMdbRecordset( mdbFile, SQL, mdbPassword )

 rsScore.AddNew
  'rsScore("�Ǹ�") = CLng(UserID)
  'rsScore("�Ǹ�") =Trim(Right(UserID,6))
 rsScore("�Ǹ�") =Trim(UserID)
 rsScore("�m�W") =  Name
 rsScore("MATH") = -1
 rsScore("ENGLS") = -1
 rsScore("CHINE") = -1
rsScore.Update

ShowMessage "�z�w�g�����������|���F�I","y"

%>

<%
'' Sub ShowMessage( msg )
  Sub ShowMessage( msg,yn )
 
%>

   <CENTER>
   <H2><%=msg%><HR></H2>
    
   <FORM action="forget-1r.aspx" method= "get"  >
  <%
     If yn = "y" Then
  '' Response.Redirect "login-1r.asp?" & Request.QueryString

   %>
       <p><input type="submit" value="�жǵ��ڨϥΪ̱K�X "></p>
    <%
      End If
    %>
      
     <INPUT Type=Button Value="��^�W�@��" OnClick="history.back();">
        <!--<A HREF="Enterkac-1.asp">���s�i�J����</A>
     <A HREF="login-1t.asp">���s�i�J����[����update���\��form action]</A>--> 

   </FORM>
   </CENTER>
  
<%
   Response.End 
End Sub 
%>

















































































