<!-- #include virtual="/kjasp/func/DB.fun" -->

<%
Name=Request("Name")
Email=Request("Email")
Subject=Request("Subject")
Words=Request("Words")

' �O�d�U�ӡA�H�K���s��J
Session("Name") = Name
Session("Email") = Email

If Subject=Empty Or Words=Empty Or Name=Empty Or Email=Empty Then
   Msg = "���d�աA�������I"
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
   Msg = "�z���Q�ץD�D�w�[�J�Q�׸s�դ��I"
   Session("Subject") = ""
   Session("Words") = ""
End If

Response.Redirect "Title.asp?Msg=" & Msg 
%>