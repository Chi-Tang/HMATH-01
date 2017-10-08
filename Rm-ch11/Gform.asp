<!-- #include virtual="/kjasp/func/DB.fun" -->
<%
' 讀取各欄位的資料
Name = Request("Name")
Email = Request("Email")
Subject = Request("Subject")
Memo = Request("Memo")
Pic = Request("Pic")

' 檢查各欄位是否輸入有資料
If Name = "tang1206"  Then
   'ShowMessage "欄位留白不接受!"
   Response.Redirect "gbook.asp"
End If 

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

'' 將網頁導至「觀看留言」的網頁 gbook.asp
  ''Response.Redirect "gbook.asp"
  ' Response.Redirect "Gform.htm"
 ShowMessage "您已經留言成功,謝謝!"

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













