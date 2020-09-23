VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Transfer"
   ClientHeight    =   1560
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4470
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1560
   ScaleWidth      =   4470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtPort2 
      Height          =   285
      Left            =   3810
      TabIndex        =   6
      Text            =   "333"
      Top             =   1140
      Width           =   585
   End
   Begin VB.TextBox txtPort1 
      Height          =   285
      Left            =   3810
      TabIndex        =   5
      Text            =   "333"
      Top             =   780
      Width           =   585
   End
   Begin VB.TextBox txtReceiveFile 
      Height          =   285
      Left            =   1410
      TabIndex        =   4
      Text            =   "C:\AUTOEXEC.BACKUP"
      Top             =   1140
      Width           =   2295
   End
   Begin VB.TextBox txtSendFile 
      Height          =   255
      Left            =   1410
      TabIndex        =   3
      Text            =   "C:\AUTOEXEC.BAT"
      Top             =   780
      Width           =   2295
   End
   Begin VB.CommandButton cmdReceive 
      Caption         =   "Receive File"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   90
      TabIndex        =   1
      Top             =   1140
      Width           =   1245
   End
   Begin VB.CommandButton cmdSendFile 
      Caption         =   "Send File"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   90
      TabIndex        =   0
      Top             =   750
      Width           =   1245
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H00C00000&
      BackStyle       =   1  'Opaque
      Height          =   405
      Left            =   2460
      Shape           =   4  'Rounded Rectangle
      Top             =   210
      Width           =   525
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00FF8080&
      BackStyle       =   1  'Opaque
      Height          =   435
      Left            =   2760
      Shape           =   3  'Circle
      Top             =   90
      Width           =   375
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00800000&
      BackStyle       =   1  'Opaque
      Height          =   435
      Left            =   2370
      Shape           =   4  'Rounded Rectangle
      Top             =   90
      Width           =   525
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Transfer Test"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   270
      TabIndex        =   2
      Top             =   180
      Width           =   1905
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00000000&
      FillColor       =   &H00800000&
      FillStyle       =   0  'Solid
      Height          =   885
      Left            =   0
      Top             =   660
      Width           =   4455
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00000000&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   675
      Left            =   0
      Top             =   0
      Width           =   4455
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Michael Schmidt
' Freeware / Sourceware from
' www.planetsourcecode.com/vb
'
' <<<<< SAMPLE >>>>
' Update to my December 1st code. If you like, please vote for me on PSC.
' OCX removes the WINSOCK referance problem that DLL's have when you distribute your
' code to others who do not have VB installed, or the ActiveX data object library.
' Hence you can distribute an OCX instead of the bloated library :) Plus OCX for this
' kind of thing seems more sensible.
'
' <<<<< Notes >>>>>
' OCX Controls are threaded seperate from the VB program you write using them.
' I create new instances, or multiple file transfers to/on one machine, and it
' will only transfer one at a time? If you know how to get around this, please
' write to me at mikes@mtdmarketing.com
'
' 03/26/01 - More notes in OCX

Private Sub cmdReceive_Click()
Dim NewForm As frmFileReceive
    
    ' New Instance
    Set NewForm = New frmFileReceive
    NewForm.Show

End Sub

Private Sub cmdSendFile_Click()
Dim NewForm As frmFileSend
    
    ' New Instance
    Set NewForm = New frmFileSend
    NewForm.Show
End Sub
