<%
Function GCounter( counter )
   Dim S, i, G
   S = CStr( counter ) ' ���N�ƭ��ন�r�� S

   ' �v�@���r��S���C�@�Ӧr��, �M��ꦨ <IMG SRC=?.gif> ���ϧμХ�
   For i = 1 to Len(S)
      G = G & "<IMG SRC=/kjasp/func/" & Mid(S, i, 1) & ".gif Align=TextTop>"
   Next
   GCounter = G
End Function
%>