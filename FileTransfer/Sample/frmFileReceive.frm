VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{B65D1865-FA43-433A-9247-3B005D55C695}#1.0#0"; "NetTransfer.ocx"
Begin VB.Form frmFileReceive 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "File Receive"
   ClientHeight    =   1875
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4035
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmFileReceive.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1875
   ScaleWidth      =   4035
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin NetTransfer.FileIN FileIN1 
      Height          =   300
      Left            =   600
      TabIndex        =   8
      Top             =   120
      Width           =   300
      _ExtentX        =   529
      _ExtentY        =   529
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   315
      Left            =   3000
      TabIndex        =   1
      Top             =   1260
      Width           =   975
   End
   Begin VB.Timer timBPS 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3600
      Top             =   2040
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   1620
      Width           =   4035
      _ExtentX        =   7117
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   413
            MinWidth        =   413
            Picture         =   "frmFileReceive.frx":000C
            Key             =   "icon"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   2646
            MinWidth        =   2646
            Text            =   "0.0.0.0"
            TextSave        =   "0.0.0.0"
            Key             =   "ip"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   1196
            MinWidth        =   1196
            Text            =   "0"
            TextSave        =   "0"
            Key             =   "port"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   2725
            Key             =   "state"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList imgConnect 
      Left            =   3360
      Top             =   2040
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   14
      ImageHeight     =   14
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileReceive.frx":03C4
            Key             =   "off"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileReceive.frx":077C
            Key             =   "on"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ProgressBar ProgressBar 
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   2835
      _ExtentX        =   5001
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   90
      Picture         =   "frmFileReceive.frx":08DE
      Top             =   60
      Width           =   240
   End
   Begin VB.Label lblTransfered 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0 / 0"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1080
      TabIndex        =   7
      Top             =   630
      Width           =   2895
   End
   Begin VB.Label lblBPS 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0 / bps"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1080
      TabIndex        =   6
      Top             =   930
      Width           =   2895
   End
   Begin VB.Label lblFile 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "(Not Available)"
      Height          =   255
      Left            =   1080
      TabIndex        =   5
      Top             =   180
      Width           =   2895
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Completed:"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   630
      Width           =   825
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Speed:"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   240
      TabIndex        =   3
      Top             =   960
      Width           =   510
   End
End
Attribute VB_Name = "frmFileReceive"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdCancel_Click()
    FileIN1.Cancel
    Unload Me
End Sub


Private Sub FileIN1_Canceled()
    Me.Caption = "Canceled"
End Sub


Private Sub FileIN1_Connected()
    
    StatusBar.Panels("icon").Picture = imgConnect.ListImages("on").Picture
    StatusBar.Panels("ip").Text = FileIN1.RemoteIP
    StatusBar.Panels("port").Text = FileIN1.LocalPort
    StatusBar.Panels("state").Text = SocketState(FileIN1.GetState)
    
    timBPS.Enabled = True
    lblFile = GrabFilename(FileIN1.LocalFile)

End Sub

Private Sub FileIN1_FileComplete()

    Me.Caption = "File Complete!"
    FileIN1.Disconnect

End Sub


Private Sub FileIN1_SockError(ErrorStats As String)
    MsgBox ErrorStats, vbCritical, "Connection Failed"
End Sub


Private Sub FileIN1_Transfered(Percent As Long, Bytes As String)

    ProgressBar.Value = Percent
    lblTransfered = Bytes

End Sub


Private Sub Form_Load()
On Error GoTo ErrSub

    FileIN1.LocalPort = frmMain.txtPort2
    FileIN1.Listen

    StatusBar.Panels("icon").Picture = imgConnect.ListImages("off").Picture
    StatusBar.Panels("ip").Text = FileIN1.RemoteIP
    StatusBar.Panels("port").Text = FileIN1.LocalPort
    StatusBar.Panels("state").Text = SocketState(FileIN1.GetState)


Exit Sub
ErrSub:
    MsgBox Err.Number & ":" & Err.Description & ":" & Err.Source, vbCritical
End Sub


Private Sub Form_Unload(Cancel As Integer)

        FileIN1.Cancel

End Sub


Private Sub timBPS_Timer()
    lblBPS = FormatBytes(FileIN1.BPS) & "ps"
End Sub
