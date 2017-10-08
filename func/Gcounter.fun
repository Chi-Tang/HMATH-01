<%
Function GCounter( counter )
   Dim S, i, G
   S = CStr( counter ) ' 先將數值轉成字串 S

   ' 逐一取字串S的每一個字元, 然後串成 <IMG SRC=?.gif> 的圖形標示
   For i = 1 to Len(S)
      G = G & "<IMG SRC=/kjasp/func/" & Mid(S, i, 1) & ".gif Align=TextTop>"
   Next
   GCounter = G
End Function
%>