Attribute VB_Name = "modMain"
Option Explicit


Public Enum Packet
    P1024 = 1024
    P2048 = 2048
    P4096 = 4096
    P8192 = 8192
End Enum





Public Function SocketState(iSocket As Integer) As String

    Select Case iSocket
    Case 1: SocketState = "Open"
    Case 2: SocketState = "Listening"
    Case 3: SocketState = "Connection Pending"
    Case 4: SocketState = "Resolving Host"
    Case 5: SocketState = "Host Resolved"
    Case 6: SocketState = "Connecting"
    Case 7: SocketState = "Connected"
    Case 8: SocketState = "Closing"
    Case 9: SocketState = "Error"
    End Select

End Function
'====================================
'   FormatBytes
'====================================
Public Function FormatBytes(iBytes As Long) As String

    If iBytes < 1024 Then
        FormatBytes = iBytes & " b"
    ElseIf iBytes < 1048576 Then
        FormatBytes = Format(iBytes / 1024, "0.0") & " kb"
    Else 'If iBits < 1000000000 Then
        FormatBytes = Format(iBytes / 1048576, "0.00") & " mb"
    End If

End Function
'====================================
'   GrabFilename
'====================================
Public Function GrabFilename(FullPath As String)

    ' Pulls the filename from a full path and filename.
    ' Returns filename.
    GrabFilename = Right(FullPath, Len(FullPath) - InStrRev(FullPath, "\"))

End Function
