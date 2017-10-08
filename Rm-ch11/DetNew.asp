<!--#include virtual="/kjasp/func/DB.fun" -->

<%
TitleID=Request("TitleID")
Name=Request("Name")
Email=Request("Email")
Subject=Request("Subject")
Words=Request("Words")

Session("Name") = Name
Session("Email") = Email

If Subject="" Or Words="" Or Name="" Or Email="" Then 
   Msg = "欄位留白，不接受！"
   Session("Subject") = Subject
   Session("Words") = Words
Else
   Set rs = GetMdbRecordset( "news.mdb", "Details" )
   rs.AddNew
   rs("Name") = Name
   rs("Email") = Email
   rs("Subject") = Subject
   rs("Words") = Words
   rs("TitleID") = TitleID
   rs.Update
   Msg = "您的意見已加入討論主題中！"
   Session("Subject") = ""
   Session("Words") = ""

   Set cmd = Server.CreateObject( "ADODB.Command")
   Set cmd.ActiveConnection = rs.ActiveConnection
   SQL = "Update Titles Set LastNewsDate=Now(), [Number]=[Number]+1 "
   SQL = SQL & " Where TitleID=" & TitleID
   cmd.CommandText = SQL
   cmd.Execute
End If

Response.Redirect "Detail.asp?TitleID=" & TitleID & "&Msg=" & Msg
%>
