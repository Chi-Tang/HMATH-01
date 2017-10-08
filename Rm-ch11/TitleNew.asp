<!-- #include virtual="/kjasp/func/DB.fun" -->

<%
Name=Request("Name")
Email=Request("Email")
Subject=Request("Subject")
Words=Request("Words")

' 保留下來，以免重新輸入
Session("Name") = Name
Session("Email") = Email

If Subject=Empty Or Words=Empty Or Name=Empty Or Email=Empty Then
   Msg = "欄位留白，不接受！"
   Session("Subject") = Subject
   Session("Words") = Words
Else
   Set rs = GetMdbRecordset( "news.mdb", "Titles" )
   rs.AddNew
   rs("Name") = Name
   rs("Email") = Email
   rs("Subject") = Subject
   rs("Words") = Words
   rs("Number") = 0
   rs.Update
   Msg = "您的討論主題已加入討論群組中！"
   Session("Subject") = ""
   Session("Words") = ""
End If

Response.Redirect "Title.asp?Msg=" & Msg 
%>