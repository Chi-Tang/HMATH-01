<%
' 若傳回 True，表示信用卡卡號正確
' 若傳回 False，表示信用卡卡號錯誤
Function CheckID( ID )
  Dim intgrtmp,intgr,intgreven,intgrodd,intgrtotal,char

  If Len(ID) <> 16 Then
     CheckID = False
     Exit Function
  End If

  If Not IsNumeric(ID) Then
     CheckID = False
     Exit Function
  End If

  ChkNum = Right(ID, 1)
  Number = Left(ID, 15)

  For intgr = len(Number) -1 To 2 Step -2
      char = Mid(Number,intgr,1)
      If IsNumeric(char) then
         intgreven = intgreven + CInt(char)
      End If
  Next 
            
  For intgr = Len(Number) To 1 Step -2
      char = Mid(Number,intgr,1)
      If IsNumeric(char) then
         intgrtmp = CInt(Char) * 2
         If intgrtmp > 9 then
            intgrodd = intgrodd + ( intgrtmp \ 10 ) + (intgrtmp - 10)
         Else
            intgrodd = intgrodd + intgrtmp
         End If
      End If
  Next 

  intgrtotal = intgreven + intgrodd
  intgrtmp = 10 - (intgrtotal mod 10)
  if intgrtmp < 10 then
     CheckID = intgrtmp
  else
     CheckID = 0
  end if

  If CInt(ChkNum) = CInt(CheckID) then
     CheckID = True
  else     
     CheckID = False
  end if

End function
%>