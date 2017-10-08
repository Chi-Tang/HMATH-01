<% ' FileCnt.fun: 訪客計數器, 檔案版本
Function kjFileCounter( counter_file )
    Dim fs, file, txt
    Application.Lock
    Set fs = Server.CreateObject("Scripting.FileSystemObject")
    file = Server.MapPath( counter_file )

    Set txt = fs.OpenTextFile( file, 1, True )
    If Not txt.atEndOfStream Then kjFileCounter = CLng(txt.ReadLine)
    kjFileCounter = kjFileCounter + 1
    txt.Close

    On Error Resume Next
    Set txt = fs.CreateTextFile( file, True )
    If Err.Number = 70 Then
        kjFileCounter = "請先將 " & file & " 的唯讀屬性去除!"
    ElseIf Err.Number <> 0 Then
        kjFileCounter = Err.Description
    Else
       txt.WriteLine kjFileCounter
       txt.Close
    End If
    Application.UnLock
End Function

' 此一函數只是讀出訪客計數值, 不會將訪客計數器加一
Function kjReadFileCounter( counter_file )
    Dim fs, file, txt
    Set fs = Server.CreateObject("Scripting.FileSystemObject")
    file = Server.MapPath( counter_file )

    Set txt = fs.OpenTextFile( file, 1, True )
    If Not txt.atEndOfStream Then 
        kjReadFileCounter = CLng(txt.ReadLine)
    Else
        kjReadFileCounter = 0
    End If
    txt.Close
End Function
%>