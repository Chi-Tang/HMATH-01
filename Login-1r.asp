
<!-- #include virtual="/Hmath/func/DB.fun" -->

<%
  ITEM=Request("ITEM")
  SECTM=Request("SECTM")
   Dim URL  
   URL=SECTM&ITEM
  

If mdbFile = Empty Then mdbFile = "/Hmath/UsersPwd.mdb"
 'If mdbFile = Empty Then mdbFile = "UsersPwd.mdb"
If mdbPassword = Empty Then mdbPassword = "kj6688"
If UserID = Empty Then UserID = Request("UserID")
If Password = Empty Then Password = Request("Password")

'''Passed = Session( mdbFile )
If Passed <> "Passed" And UserID <> Empty And Password <> Empty Then
   Passed = CheckPassword( UserID, Password, mdbFile, mdbPassword )
End If

 

If Passed <> "Passed" Then
%>

   <HTML>
   <BODY BGCOLOR="#FFFFFF">
   <H2 ALIGN=CENTER>請輸入數學帳號(任一英文字和學號)及密碼：</H2>
   <CENTER>
   <FORM Action=<%=Request.ServerVariables("PATH_INFO")%> Method=POST>
     <INPUT Type=Hidden Name=SECTM Value=<%=SECTM%>>
     <INPUT Type=Hidden Name=ITEM Value=<%=ITEM%>>
        
   <TABLE BORDER=1 CELLSPACING=0 >
   <TR>
   <TD ALIGN=RIGHT>使用者帳號:</TD>
   <TD><Input Type=Text Name=UserID Size=12 Value=<%=UserID%>></TD>
   </TR>

   <TR><TD ALIGN=RIGHT>密碼:</TD>
   <TD><Input Type=Password Name=Password Size=12 Value=<%=Password%>></TD>
   </TR>
   </TABLE>
    <P> <Input Type=Submit Value=" 確 定 "></P>
   
        <P>  <FONT Color=Red>第一次使用者請先加入會員,即可由EMAIL取得密碼使用！</FONT></P>
        <P>  <FONT Color=Red>使用者帳號或密碼錯誤，請重新輸入！</FONT></P> 
    
        <P>   <A HREF="Member-1r.asp">加入新會員(或忘了帳號,密碼)</A></P> 
       
   
   </FORM>
   </CENTER>
   <HR>
   <% If Passed = "NotPassed" Then  ' 不是第一次進入 %>
    
   <%  '' Response.redirect "Member-1r.asp"  %>
   <%    Response.redirect "Login-1r.asp"  %>
      
   <%End If%> 
  
  
   
   </BODY>
   </HTML>
<% 
      Response.End
      
   Else
        
        Response.redirect URL

         
        ' Response.redirect("TBKI-3/TBKI-3H.htm")   
          ' Response.redirect("DNT3/DNT-3H.HTM") 
End If
%><%
Function CheckPassword( UserID, Password, mdbFile, mdbPassword )
   If Session( mdbFile ) = "Passed" Then
       CheckPassword = "Passed"
     
       Exit Function ' 不必再檢查資料庫，以節省時間
  
    End If

   SQL = "Select * From Users Where UserID='" & UserID & "'"
   SQL = SQL & " And Password = '" & Password & "'"
   If Len(mdbPassword) = 0 Then
      Set rs = GetMdbRecordset( mdbFile, SQL )
   Else
      Set rs = GetSecuredMdbRecordset( mdbFile, SQL, mdbPassword )
   End If
   
   If Not rs.EOF Then 	' 此組密碼存在，故通過
       CheckPassword = "Passed"
       Session( mdbFile ) = "Passed"
   Else
       CheckPassword = "NotPassed"
   End If
End Function
%>