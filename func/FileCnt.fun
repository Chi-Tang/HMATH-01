<% ' FileCnt.fun: �X�ȭp�ƾ�, �ɮת���
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
        kjFileCounter = "�Х��N " & file & " ����Ū�ݩʥh��!"
    ElseIf Err.Number <> 0 Then
        kjFileCounter = Err.Description
    Else
       txt.WriteLine kjFileCounter
       txt.Close
    End If
    Application.UnLock
End Function

' ���@��ƥu�OŪ�X�X�ȭp�ƭ�, ���|�N�X�ȭp�ƾ��[�@
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