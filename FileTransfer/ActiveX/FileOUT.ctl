VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.UserControl FileOUT 
   ClientHeight    =   300
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   300
   Picture         =   "FileOUT.ctx":0000
   ScaleHeight     =   300
   ScaleWidth      =   300
   ToolboxBitmap   =   "FileOUT.ctx":04F2
   Begin MSWinsockLib.Winsock Socket 
      Left            =   1920
      Top             =   300
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "FileOUT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
Private Ready As Boolean        ' Controls if data is sent or not.
Private BitsSecond As Long      ' Bits per second (call every second)
'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
Public Event Connected()
                        ' When remote accepts our connection, it sends us a packet,
                        ' when we receive it, we raise event -connected-.
Public Event FileComplete()
                        ' When file transfer is complete, raise event -filecomplete-
Public Event SockError(ErrorStats As String)
                        ' When winsock generates an error, we simply pass it on
                        ' to the controller, by raising the event and passing.
Public Event Transfered(Percent As Long, Bytes As String)
                        ' Every time we send file data, we raise an event telling
                        ' percent completed and bytes of total file sent.
Public Event Canceled() ' User canceled. Tell remote to cancel and take care
                        ' of closing file.
'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
Private mvarLocalFile As String
Private mvarRemoteFile As String
Private mvarRemoteIP As String
Private mvarRemotePort As Long
Private mvarPacketSize As Integer
Private mvarFileSize As Variant
'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||


Public Sub Cancel()
' Tell remote to cancel and close file.
' Generate local event canceled.
    
    On Error Resume Next
    ' if error, it's because sock is closed...

    Socket.SendData "040" & NewCom
    DoEvents
    Socket.Close
    RaiseEvent Canceled

End Sub


Public Function BPS() As Long
' Call this sub ever second to see bps
    
    BPS = BitsSecond
    BitsSecond = 0

End Function


Public Property Get FileSize() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FileSize
    If IsObject(mvarFileSize) Then
        Set FileSize = mvarFileSize
    Else
        FileSize = mvarFileSize
    End If
End Property


Public Sub Disconnect()
' Disconnect socket.

    Socket.Close
    
End Sub


Public Function GetState()
' Return state of socket.

    GetState = Socket.State

End Function


Public Sub Connect()
' Connect socket to remote.

    If Not isValidFile(LocalFile) Then
        RaiseEvent SockError("Invalid File!")
        Exit Sub
    End If

    Socket.Connect RemoteIP, RemotePort

End Sub


Private Sub Socket_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Dim ErrorStats As String

    ErrorStats = Number & " : " & Description
    RaiseEvent SockError(ErrorStats)

End Sub


Private Sub Socket_DataArrival(ByVal bytesTotal As Long)
Dim iPacket As String
Dim iCom As String

    Socket.GetData iPacket
    iCom = Word(iPacket, 1, NewCom)

    Select Case iCom
    Case "010": SendFileStats
    Case "020": SendFileData
    Case "040": RaiseEvent Canceled
    Case "999": Ready = True
    End Select

End Sub


Private Sub SendFileStats()
Dim iPacket As String

    ' We are connected.
    RaiseEvent Connected

    ' Set up packet. Tell remote about file stats.
    iPacket = "010" & NewCom
    iPacket = iPacket & RemoteFile & NewCom
    iPacket = iPacket & GetFileSize(LocalFile) & NewCom
    
    ' Send Packet
    Socket.SendData iPacket

End Sub


'====================================
'   Send File
'====================================
Private Sub SendFileData()
Dim iPacket As String
Dim BytesSent As Long
Dim SentSize As Long
Dim Buffer As Long
Dim FileNumber
Dim DelLastByte As Boolean

    On Error GoTo ErrSub
    
    FileSize = GetFileSize(LocalFile)

    BitsSecond = 0
    SentSize = 0
    ' Buffer size size will be BUFFER - 4bits because every packet
    ' has 4 bits of command line added to it (040XRAWDATA...
    Buffer = PacketSize - 4

    ' - At the end of the file, we usually end up sending
    ' - the EOF character, meaning each sent file gets 1byte
    ' - larger than it should, so now we subtract at the end...
    DelLastByte = False
    ' - Grab Local File Handle (Random = FreeFile)
    ' - Open it. Loop through while not at end of file.
    FileNumber = FreeFile
    Open LocalFile For Binary Access Read As #FileNumber
    Do While Not EOF(FileNumber)
     
    ' end of file...size buffer appropriately
    If FileSize - Loc(FileNumber) <= Buffer Then
        Buffer = FileSize - Loc(FileNumber) + 1
        DelLastByte = True
    End If

        iPacket = ""
        iPacket = Space$(Buffer)
        Get FileNumber, , iPacket
        
        ' end of file, our packet is 1 bit to large.
        If DelLastByte = True And FileSize <> 0 Then iPacket = Left(iPacket, Len(iPacket) - 1)
        
        ' Set bit trackers...
        BytesSent = BytesSent + Len(iPacket)
        BitsSecond = BitsSecond + Len(iPacket) + 4
        
        Socket.SendData "020" & NewCom & iPacket
        
        If FileSize = 0 Then RaiseEvent Transfered(100, 0)
        If FileSize <> 0 Then RaiseEvent Transfered(CInt(BytesSent / FileSize * 100), FormatBytes(BytesSent) & " / " & FormatBytes(FileSize))
        
        WaitForRemote

    Loop

    Close #FileNumber
    
    ' Tell Remote Finished!
    Socket.SendData "030" & NewCom
    RaiseEvent FileComplete
 
 Exit Sub

ErrSub:
Select Case Err.Number
    Case 40006 ' No Connection Detected
        Close #FileNumber

    Case Else
            MsgBox "Error Number: " & Err.Number & vbCrLf & Err.Description, vbCritical, "ActiveX Error (Data Arrival)"
            Exit Sub 'Resume Next
End Select
End Sub


Private Sub WaitForRemote()

    Ready = False
    While Ready = False
    DoEvents
    Wend
    

End Sub


Public Property Let PacketSize(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.PacketSize = 5
    mvarPacketSize = vData
End Property


Public Property Get PacketSize() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PacketSize
    PacketSize = mvarPacketSize
End Property


Public Property Let RemotePort(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RemotePort = 5
    mvarRemotePort = vData
End Property


Public Property Get RemotePort() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.RemotePort
    RemotePort = mvarRemotePort
End Property



Public Property Let RemoteIP(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RemoteIP = 5
    mvarRemoteIP = vData
End Property


Public Property Get RemoteIP() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.RemoteIP
    RemoteIP = mvarRemoteIP
End Property



Public Property Let RemoteFile(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RemoteFile = 5
    mvarRemoteFile = vData
End Property


Public Property Get RemoteFile() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.RemoteFile
    RemoteFile = mvarRemoteFile
End Property



Public Property Let LocalFile(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.LocalFile = 5
    mvarLocalFile = vData
End Property


Public Property Get LocalFile() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.LocalFile
    LocalFile = mvarLocalFile
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


Private Sub UserControl_Resize()
    UserControl.Width = 300
    UserControl.Height = 300
End Sub


