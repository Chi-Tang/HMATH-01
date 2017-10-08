<HTML>
<BODY BgColor=White>
<H2>清除檔案的唯讀屬性<HR></H2>
<UL>
<%
On Error Resume Next
Set fs = Server.CreateObject("Scripting.FileSystemObject")
If Err.Number <> 0 Then
   Response.Write "請先安裝好 IIS 或 PWS 在執行本程式!"
   Response.End
End If

ClearReadonlyAttr fs.GetFolder(Server.MapPath("."))

Sub ClearReadonlyAttr( fd )
   For each f In fd.Files
       f.Attributes = f.Attributes And &HFFFFFFFE
   Next

   For each sfd In fd.SubFolders
      ClearReadonlyAttr sfd
   Next

   Response.Write "<LI>" & fd.Path & " ... 已清除!<P>"
End Sub
%>
</UL><HR>
</BODY>
</HTML>