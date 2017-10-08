<!-- #include virtual="/kjasp/func/DB.fun" -->
<%
Lesson = Request("Lesson")
No = Request("No")
Name = Request("Name")
'Set rs = GetMdbRecordset( "Testa.mdb", Lesson )
Set rs = GetMdbRecordset( "Testac-1.mdb", Lesson&Trim(CStr(No)))
%>
<HTML>
<BODY BgColor=White Background="../B01.jpg">
<H2>請核對答案<HR></H2>

''<%
'Score = 0
'While Not rs.EOF
 ' Sel = Request( "No" & rs("題號") )
  ' rs("做答") = Sel
  'Ans = rs("解答")
   'If Ans = Sel Then
    ' rs("核答") = "V"
  ' Else
   '   rs("核答") = "X"
   'End If
  'rs.Update
  
  ''If Ans = Sel Then
  ''   Score = Score + rs("分數")
  ''End If
  
  'rs.MoveNext
  
'Wend

''SQL = "Select * From 成績單 " 
  'SQL = SQL & "Where 學號=" & No & " And 姓名='" & Name & "'"
  'Set rsScore = GetMdbRecordset( "Test.mdb", SQL )
''SQL = SQL & "Where 學號=" & No 
''Set rsScore = GetMdbRecordset( "Testac.mdb", SQL )
''rsScore(Lesson) = score
''rsScore.Update

' Set cmd = Server.CreateObject( "ADODB.Command" )
 ' Set cmd.ActiveConnection = rs.ActiveConnection
  'SQLU ="Update "&Lesson&Trim(CStr(No))&" Set 標記=-1"  
   ' cmd.CommandText = SQLU
    'cmd.Execute

' SQL = "Select * From " & Lesson&Trim(CStr(No))
 ' Set rs = GetMdbRecordset( "Testac-1.mdb", SQL )
 '%>
 
 <TABLE Border=1>
  <TR BgColor=Yellow><TD>題號</TD><TD>核答</TD><TD>做答 </TD>><TD>解答  </TD><TD>題目</TD><TD>題型</TD><TD>詳答</TD></TR>
<%
While Not rs.EOF
   crsTN ="<A HREF ="&Trim(Cstr(rs("流水號")))&".html"&">詳答按此</A>"
               rs("詳答") = crsTN
 ' crsTN ="<A HREF =/Hmath/"&Trim(Cstr(rs("流水號")))&".html"&">詳答按此</A>"
              ' rs("詳答") = crsTN
            

%>
 <TR><TD Align=Right><%=rs("題號")%></TD>
     <TD Align=Right><%=rs("核答")%></TD>
     <TD Align=Right><%=rs("做答")%></TD>
     <TD Align=Right><%=rs("解答")%></TD>
   
 
  <% 
    crsTM =Trim(CStr(rs("題目碼")))&".html"
     'crsTM ="math/"&Trim(CStr(rs("題目碼")))&".html"
     ''crsTM ="1010104.html" & '數學圖形檔須在同目錄才會顯示
     ' Response.Write crsTM
   %>
       <TD Align=Left><%=HMCLD(crsTM) %> </TD> <TD   
         Align=Right><%=rs("類型")%></TD><TD
         Align=Right><%=rs("詳答")%></TD>
      
    
<%
      
  Response.Write "</TR>"

  <!---Response.Write "</BLOCKQUOTE>"--->

  rs.MoveNext
 Wend

 

 %>
</TABLE>

 <%
 
  
    Set cmd = Server.CreateObject( "ADODB.Command" )
  Set cmd.ActiveConnection = rs.ActiveConnection
   'SQLU ="Update "&Lesson&Trim(CStr(No))&" Set 標記=-1"  
  SQL2 ="Insert into "&Lesson&Trim(Cstr(No))&"A"& " Select * From "&Lesson&Trim(Cstr(No))&" Where 標記=100"
   cmd.CommandText = SQL2
     cmd.Execute

  SQLD ="Delete From "&Lesson&Trim(CStr(No))
    cmd.CommandText = SQLD
    cmd.Execute
 %> 

<HR>
<A HREF="EnterKac-tt.asp?No=<%=No%>&Name=<%=Name%>">參加其他科目考試</A>
<!--A HREF="Enter.asp?No=<%=No%>&Name=<%=Name%>">參加其他科目考試</A-->
<%
   Showmg "看簡答"
 %> 

</BODY>
</HTML>

 <%
 Sub Showmg( msg )
   
  %>
<CENTER>
<H2><%=msg%><HR></H2>
<FORM>
  <INPUT Type=Button Value="請核答返回上一頁" OnClick='datachk();'></FORM>
</CENTER>
<% 
End Sub 
%>

<%
Function datachk()
Dim yn
yn=MsgBox("you summit?", 65, "caution!")
 if yn=2 then 
   Exit Function
  end if
     
%>

  <TABLE Border=1>
  <TR BgColor=Yellow>
  <TD Align=Left><%=HMCLD(crsTM) %> </TD> 
 
  </TABLE >
<%  
  End Function
%>

  
<%
Function ScoreHTML( Score )
   If Score = -1 Then 	' 尚未考試
      ScoreHTML = "--"
   ELseIf Score < 60 Then
      ScoreHTML = "<FONT Color=Red>" & Score & "</FONT>"
   Else
      ScoreHTML = Score
   End If
End Function
%>

<% '測試副程式2 
  FUNCTION HMCLD(rsTM) 
   Set fs = Server.CreateObject("Scripting.FileSystemObject")
     File = Server.MapPath(rsTM)
  Set txtf = fs.OpenTextFile( File )
If Not txtf.atEndOfStream Then	' 先確定還沒有到達結尾的位置
    Content = txtf.ReadAll	' 讀取整個檔案的資料
   '' Lines = Replace(Content, vbCrLf, "<BR>" )
    ''Response.Write Lines
   ' Response.Write Content
End If
 HMCLD=Content
END FUNCTION

%>























































