<!-- #include virtual="/kjasp/func/DB.fun" -->
<%
mdbFile = "UsersPwd.mdb"
mdbPassword = "kj6688"

UserID = Request("UserID")
Password = Request("Password")
Password2 = Request("Password2")
Name = Request("Name")
Email = Request("Email")

' 檢查欄位是否留白
If UserID = Empty Or Password = Empty Or Name = Empty Then
   ShowMessage "欄位留白，不接受!"
End If

' 檢查 Email
If InStr( Email, "@" ) = 0 Or Right(Email, 1) = "@" Or Left(Email, 1) = "@" Then
   ShowMessage "Email 錯誤，請重新輸入!"
End If

' 檢查兩次輸入的密碼是否一致
If Password <> Password2 Then
   ShowMessage "兩次輸入的密碼不一致!"
End If

' 檢查此一 UserID 已有他人佔用
SQL = "Select * From Users Where UserID='" & UserID & "'"
Set rs = GetSecuredMdbRecordset( mdbFile, SQL, mdbPassword )
If Not rs.EOF Then 	' 此一 UserID 碼已被佔用
   ShowMessage "此一「使用者名稱」已被佔用，請重新輸入!"
End If

' 檢查此一 Email 是否已申請帳號
SQL = "Select * From Users Where Email='" & Email & "'"
Set rs = GetSecuredMdbRecordset( mdbFile, SQL, mdbPassword )
If Not rs.EOF Then 	' 此一 Email 已申請帳號
   ShowMessage "此一 Email 已申請帳號，請重新輸入!"
End If

rs.AddNew
rs("UserID") = UserID
rs("Password") = Password
rs("Name") = Name
rs("Email") = Email
rs.Update

SQL = "Select * From 成績單 " 
'SQL = SQL & "Where 學號=" & No & " And 姓名='" & Name & "'"
Set rsScore = GetMdbRecordset( "Testac-1.mdb", SQL )
 rsScore.AddNew
  'rsScore("學號") = CLng(UserID)
 'rsScore("學號") = CLng(Trim(Right(UserID,6)))
 rsScore("學號") = Trim(Right(UserID,6))
 rsScore("姓名") =  Name
 rsScore("MATH") = -1
 rsScore("ENGLS") = -1
 rsScore("CHINE") = -1
rsScore.Update

ShowMessage "您已經成為本站的會員了！"
%>

<% 
Sub ShowMessage( msg ) 
%>
   <CENTER>
   <H2><%=msg%><HR></H2>
   <FORM><INPUT Type=Button Value="返回上一頁" OnClick="history.back();">
   
    <!--<A HREF="Enterkac-1.asp">重新進入測試</A>--> 
   <A HREF="login-1t.asp">重新進入測試</A> 

   </FORM>
   </CENTER>
<%
   Response.End 
End Sub 
%>


























