
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
   <H2 ALIGN=CENTER>�п�J�ƾǱb��(���@�^��r�M�Ǹ�)�αK�X�G</H2>
   <CENTER>
   <FORM Action=<%=Request.ServerVariables("PATH_INFO")%> Method=POST>
     <INPUT Type=Hidden Name=SECTM Value=<%=SECTM%>>
     <INPUT Type=Hidden Name=ITEM Value=<%=ITEM%>>
        
   <TABLE BORDER=1 CELLSPACING=0 >
   <TR>
   <TD ALIGN=RIGHT>�ϥΪ̱b��:</TD>
   <TD><Input Type=Text Name=UserID Size=12 Value=<%=UserID%>></TD>
   </TR>

   <TR><TD ALIGN=RIGHT>�K�X:</TD>
   <TD><Input Type=Password Name=Password Size=12 Value=<%=Password%>></TD>
   </TR>
   </TABLE>
    <P> <Input Type=Submit Value=" �T �w "></P>
   
        <P>  <FONT Color=Red>�Ĥ@���ϥΪ̽Х��[�J�|��,�Y�i��EMAIL���o�K�X�ϥΡI</FONT></P>
        <P>  <FONT Color=Red>�ϥΪ̱b���αK�X���~�A�Э��s��J�I</FONT></P> 
    
        <P>   <A HREF="Member-1r.asp">�[�J�s�|��(�ΧѤF�b��,�K�X)</A></P> 
       
   
   </FORM>
   </CENTER>
   <HR>
   <% If Passed = "NotPassed" Then  ' ���O�Ĥ@���i�J %>
    
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
     
       Exit Function ' �����A�ˬd��Ʈw�A�H�`�ٮɶ�
  
    End If

   SQL = "Select * From Users Where UserID='" & UserID & "'"
   SQL = SQL & " And Password = '" & Password & "'"
   If Len(mdbPassword) = 0 Then
      Set rs = GetMdbRecordset( mdbFile, SQL )
   Else
      Set rs = GetSecuredMdbRecordset( mdbFile, SQL, mdbPassword )
   End If
   
   If Not rs.EOF Then 	' ���ձK�X�s�b�A�G�q�L
       CheckPassword = "Passed"
       Session( mdbFile ) = "Passed"
   Else
       CheckPassword = "NotPassed"
   End If
End Function
%>