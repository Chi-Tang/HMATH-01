


<!-- #include virtual="/kjasp/func/DB.fun" -->
<%
 
Lesson = Request("Lesson")
No =  Request("No")
Name = Request("Name")

 'SQL = "Select * From " & Lesson & " Order By 題號"
 'Set rs = GetMdbRecordset( "Test.mdb", SQL )
 SQL = "Select * From " & Lesson&Trim(CStr(No))
 Set rs = GetMdbRecordset( "Testac-1.mdb", SQL )

%>

<HTML>
 <BODY BgColor=White Background="../B01.jpg">
'<BODY BgColor=White >
<H2>考試科目 -- <%=Lesson%><HR></H2>
<FORM Action=ScoreKac-1c.asp Method=POST>
<INPUT Type=Hidden Name=Lesson Value=<%=Lesson%>>
<INPUT Type=Hidden Name=No Value=<%=No%>>
<INPUT Type=Hidden Name=Name Value=<%=Name%>>
  
 <TABLE Border=1>
  <TR BgColor=Yellow><TD>題號</TD><TD>選項</TD><TD>  </TD><TD>  </TD><TD>  </TD><TD>題目</TD><TD>題型</TD><TD>分數</TD></TR>
<%
While Not rs.EOF
%>
 <TR><TD Align=Right><%=rs("題號")%></TD>
   <%
     For I = 1 To 4 
      If rs("類型") = "單選" Then
       TestType = "Radio"
      Else
       TestType = "CheckBox"
      End If
   
  %>

   <TD Align=Right><%=I%> <INPUT Type=<%=TestType%> Name=No<%=rs("題號")%> 
     Value=<%=I%>></TD>
    
   
  <%   Next 
    crsTM =Trim(CStr(rs("題目碼")))&".html"
    'crsTM ="math/"&Trim(CStr(rs("題目碼")))&".html"
    ''crsTM ="1010104.html" & '數學圖形檔須在同目錄才會顯示
    '''Response.Write crsTM
   %>
       <TD Align=Left><%=HMCLD(crsTM) %> </TD> <TD   
         Align=Right><%=rs("類型")%></TD><TD
         Align=Right><%=rs("分數")%></TD>
    
<%
  '' Next
   
    
  Response.Write "</TR>"

  <!---</TR>--->
  <!---Response.Write "</BLOCKQUOTE>"--->

  rs.MoveNext
 Wend

%>

</TABLE>


<INPUT Type=Submit Value="  繳  卷  ">
</FORM>
<HR>
</BODY>
   <script>
    document.onselectionchange=__OnSelectionChange;
       var running=false;
     function __OnSelectionChange()
       { 
       if (running==true) return;
          running=true;
       document.selection.empty();
       running=false;       
        }
  </script>
</HTML>

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
 
   
   
<% '測試副程式1
   FUNCTION HMCOD(rsTM) 
    '巨集WORD建立
    gher = ""
     ' mypos1 = 1
     ' mypos2 = 1
      mysear1 = "{"
      mysear2 = "}"      
     mytext=rsTM
     mylen = Len(mytext)
    For j= 0 TO mylen
      mypos1 = InStr(1, mytext, mysear1, 1)  
      mypos2 = InStr(1, mytext, mysear2, 1)  
     If mypos1 <> 0 Then
          textf = Mid(mytext, 1, mypos1-1)
          textm = Mid(mytext,mypos1+1, mypos2-1-mypos1)
          textb = Mid(mytext,mypos2+1, mylen)
        gher=gher+textf
          'RESPONSE.WRITE textf
          If textm <> "" Then
               textmf = Rs(textm)
              gher=gher+textmf
              'RESPONSE.WRITE textmf
           End If
             textm = ""
             mytext = Textb
               mylen = Len(mytext)
       Else
          Textb = mytext
             mylen = Len(Textb)
         gher=gher+textb
           'RESPONSE.WRITE textb
          Exit For
      End If     
   Next
  HMCOD=gher
END FUNCTION
%>      
    

    
 





















