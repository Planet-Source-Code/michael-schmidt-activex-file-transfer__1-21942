VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.UserControl FileIN 
   ClientHeight    =   300
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   315
   Picture         =   "FileIN.ctx":0000
   ScaleHeight     =   300
   ScaleWidth      =   315
   ToolboxBitmap   =   "FileIN.ctx":04F2
   Begin MSWinsockLib.Winsock Socket 
      Left            =   1440
      Top             =   360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "FileIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
Private FileNumber As Long      ' Holds File Handle
Private BytesTransfered As Long ' Holds BytesTransfered (BytesTransfered / TotalBytes)
Private BitsSecond As Long      ' Holds bits transmitted. (call every second to find bps)
'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
Public Event Connected()                        ' Tell remote we're connected.

Public Event SockError(ErrorStats As String)    ' Tell remote we have error.

Public Event Transfered(Percent As Long, Bytes As String)
                                                ' Every time bytes are written to
                                                ' a file, we raise this
                                                ' event, telling the user how much
                                                ' has been transmitted and
                                                ' what percent of the total file.
                                                
Public Event FileComplete()                     ' Tell remote we've finished the file transfer.

Public Event Canceled()                         ' Tell remote and local that user has canceled.
'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
Private mvarLocalPort As Long 'local copy
Private mvarRemoteIP As Variant 'local copy
Private mvarFileSize As Variant 'local copy
Private mvarLocalFile As Variant 'local copy
'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||



Public Function BPS() As Long
' Call this sub ever second to see bps
    
    BPS = BitsSecond
    BitsSecond = 0

End Function


Public Sub Cancel()
' Tell remote we have canceled and
' raise local event canceled.
    On Error Resume Next
    ' if error, it's cause sock closed
    
    Socket.SendData "040" & NewCom
    DoEvents
    Socket.Close
    RaiseEvent Canceled
    Close #FileNumber

End Sub


Public Function GetState()
' Return socket state.

    GetState = Socket.State

End Function

Public Sub Disconnect()
' Disconnect.

    Socket.Close
    
End Sub


Public Sub Listen()
' Listen for connection.
' Set class port to whatever the socket pulls, if
' user entered 0 as port, then it would pull a random
' port, so we must reset our class port to the socket port.

    Socket.Close
    Socket.LocalPort = LocalPort
    Socket.Listen
    LocalPort = Socket.LocalPort
    

End Sub


Private Sub Socket_ConnectionRequest(ByVal requestID As Long)
' Accept connection. Tell remote we're ready
' for data transfer. Remote will then raise
' the (connected) event.

    Socket.Close
    Socket.Accept requestID
    
    Socket.SendData "010" & NewCom
    RemoteIP = Socket.RemoteHostIP
    
End Sub

Private Sub Socket_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Dim ErrorStats As String

    ErrorStats = Number & " : " & Description
    RaiseEvent SockError(ErrorStats)

End Sub


Private Sub Socket_DataArrival(ByVal bytesTotal As Long)
Dim iPacket As String
Dim iCom As String

    BitsSecond = BitsSecond + bytesTotal

    ' Parse Packet
    Socket.GetData iPacket
    iCom = Word(iPacket, 1, NewCom)
    iPacket = Right(iPacket, Len(iPacket) - Len(iCom) - 1)

    Select Case iCom
    Case "010": ReadyFile (iPacket)
    Case "020": WriteFileData (iPacket)
    Case "030": Close #FileNumber
                RaiseEvent FileComplete
    Case "040": ' Canceled
                Close #FileNumber
                RaiseEvent Canceled
                Me.Disconnect
    'Case "999": Ready = True
    End Select

End Sub
Private Sub ReadyFile(iData As String)

    ' Set Packet Data
    LocalFile = Word(iData, 1, NewCom)
    FileSize = Word(iData, 2, NewCom)

    ' Raise Event Transfered
    BytesTransfered = 0
    RaiseEvent Connected
    RaiseEvent Transfered(0, Format(BytesTransfered) & " / " & FormatBytes(FileSize))

    ' File Operations
    FileNumber = FreeFile
    Open LocalFile For Binary Access Write As #FileNumber
    
    ' Tell remote send next packet.
    Socket.SendData "020" & NewCom

End Sub


Private Sub WriteFileData(iData As String)

    Put FileNumber, , iData

    BytesTransfered = BytesTransfered + Len(iData)
    
    If BytesTransfered <> 0 And FileSize <> 0 Then _
    RaiseEvent Transfered(CInt(BytesTransfered / FileSize * 100), FormatBytes(BytesTransfered) & " / " & FormatBytes(FileSize)) Else
    RaiseEvent Transfered(100, 0)

    Socket.SendData "999" & NewCom

End Sub


Public Property Let LocalFile(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.LocalFile = 5
    mvarLocalFile = vData
End Property


Public Property Set LocalFile(ByVal vData As Variant)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.LocalFile = Form1
    Set mvarLocalFile = vData
End Property


Public Property Get LocalFile() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.LocalFile
    If IsObject(mvarLocalFile) Then
        Set LocalFile = mvarLocalFile
    Else
        LocalFile = mvarLocalFile
    End If
End Property



Public Property Let FileSize(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FileSize = 5
    mvarFileSize = vData
End Property


Public Property Set FileSize(ByVal vData As Variant)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.FileSize = Form1
    Set mvarFileSize = vData
End Property


Public Property Get FileSize() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FileSize
    If IsObject(mvarFileSize) Then
        Set FileSize = mvarFileSize
    Else
        FileSize = mvarFileSize
    End If
End Property



Public Property Let RemoteIP(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RemoteIP = 5
    mvarRemoteIP = vData
End Property


Public Property Set RemoteIP(ByVal vData As Variant)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.RemoteIP = Form1
    Set mvarRemoteIP = vData
End Property


Public Property Get RemoteIP() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.RemoteIP
    If IsObject(mvarRemoteIP) Then
        Set RemoteIP = mvarRemoteIP
    Else
        RemoteIP = mvarRemoteIP
    End If
End Property


Public Property Let LocalPort(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.LocalPort = 5
    mvarLocalPort = vData
End Property


Public Property Get LocalPort() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.LocalPort
    LocalPort = mvarLocalPort
End Property


Private Sub UserControl_Resize()
    UserControl.Width = 300
    UserControl.Height = 300
End Sub
