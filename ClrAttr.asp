<HTML>
<BODY BgColor=White>
<H2>�M���ɮת���Ū�ݩ�<HR></H2>
<UL>
<%
On Error Resume Next
Set fs = Server.CreateObject("Scripting.FileSystemObject")
If Err.Number <> 0 Then
   Response.Write "�Х��w�˦n IIS �� PWS �b���楻�{��!"
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

   Response.Write "<LI>" & fd.Path & " ... �w�M��!<P>"
End Sub
%>
</UL><HR>
</BODY>
</HTML>