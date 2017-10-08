<%
Application.Lock
QueryString = Request.ServerVariables("QUERY_STRING")
If Len(QueryString) > 0 Then
    page = Request("_page_")
    pos = InStr(QueryString, "&" )
    If pos > 0 Then
        QueryString = "?" & Mid(QueryString, pos + 1)
    Else
        QueryString = ""
    End If
End If

If Len(page) = 0 Then
    Response.Write "此為檢查 Cookies 之網頁!"
    Response.End
End If

If Not Response.IsClientConnected Then 
   Application("kjVisit") = 0
   Application("kjIP") = ""
End If

Response.Redirect page & QueryString
%>