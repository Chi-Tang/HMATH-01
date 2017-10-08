<!-- #include virtual="/kjasp/func/DB.fun" -->
<%
' 讀取各欄位的資料
Name = Request("Name")
Email = Request("Email")
Subject = Request("Subject")
Memo = Request("Memo")
Pic = Request("Pic")

' 檢查各欄位是否輸入有資料
If Name = "" Or Email = "" Or Subject = "" Or Memo = "" Then
   ShowMessage "欄位留白不接受!"
   Response.End
End If

'Set rs = GetMdbRecordset( "gbook.mdb", "GuestBook" )
SQL = "Select * From GuestBook"
Set rs = GetMdbRecordset( "Gbook.mdb", SQL )

rs.AddNew
rs("姓名") = Name
rs("Email") = Email
rs("主題") = Subject
rs("留言") = Memo
If Pic <> "No" Then rs("圖示") = Pic
rs.Update

' 將網頁導至「觀看留言」的網頁 gbook.asp
Response.Redirect "gbook.asp"
%>

<% 
Sub ShowMessage( msg ) 
%>
<CENTER>
<H2><%=msg%><HR></H2>
<FORM><INPUT Type=Button Value="返回上一頁" OnClick="history.back();"></FORM>
</CENTER>
<% 
End Sub 
%>
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
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
       <p><input type="submit" value="請傳給我使用者密碼 "></p>
        <p>再輸入帳號：<input type="text" size="20" name="No" Value="<%=No%>"></p>
    <%
      End If
    %>
      
     <INPUT Type=Button Value="返回上一頁" OnClick="history.back();">
        <!--<A HREF="Enterkac-1.asp">重新進入測試</A>
     <A HREF="login-1t.asp">重新進入測試[不能update成功改form action]</A>--> 

   </FORM>
   </CENTER>
  
<%
   Response.End 
End Sub 
%>







