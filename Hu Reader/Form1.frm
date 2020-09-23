VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5565
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5490
   LinkTopic       =   "Form1"
   ScaleHeight     =   5565
   ScaleWidth      =   5490
   StartUpPosition =   3  'Windows Default
   Begin MSCommLib.MSComm MSComm 
      Left            =   4800
      Top             =   4800
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Close Port"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2880
      TabIndex        =   4
      Top             =   4440
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Open Port"
      Height          =   375
      Left            =   1800
      TabIndex        =   3
      Top             =   4440
      Width           =   1095
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Form1.frx":0000
      Left            =   4080
      List            =   "Form1.frx":0010
      TabIndex        =   2
      Text            =   "COM1"
      Top             =   4440
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Get Card Info."
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   4440
      Width           =   1575
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   3975
      Left            =   120
      ScaleHeight     =   3945
      ScaleWidth      =   5265
      TabIndex        =   0
      Top             =   360
      Width           =   5295
   End
   Begin VB.Label Label3 
      Caption         =   "Current Status:"
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   5280
      Width           =   5415
   End
   Begin VB.Label Label2 
      Caption         =   "ATR:"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   0
      Width           =   5295
   End
   Begin VB.Label Label1 
      Caption         =   "Atmel Code:"
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   4920
      Width           =   2175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public regComPort

Private Sub Command1_Click()
Picture1.Cls
Picture1.Print ""
Picture1.FontSize = 10
Picture1.FontBold = True
Picture1.ForeColor = &HFFFFFF
Picture1.Print " Welcome to HU Reader."
Picture1.FontSize = 8
Picture1.FontBold = False
Picture1.ForeColor = &HFF00&
Picture1.Print " Version 1.0"
Picture1.FontSize = 7
Picture1.ForeColor = &HC000&

If MSComm.PortOpen = True Then
Label1.Caption = "Atmel Code: " & GetAtmelVersion

If IsCardPresent = True Then
        Picture1.Cls
        Picture1.FontBold = True
        Call ResetForATR
        Label2.Caption = "ATR: " & GetATR
        Picture1.ForeColor = &HFFFFFF
        Picture1.Print ""
        Picture1.Print GetSTARTUPINFO
    Else
        Picture1.Print " Please insert card in loader."
    End If


Else
Label3.Caption = "Current Status: No port is open."
End If
End Sub

Private Sub Command2_Click()
Command1.Enabled = True
Command3.Enabled = True
regComPort = Combo1.Text
With MSComm
 .CommPort = Val(Right(regComPort, 1))
 .Settings = "115200,n,8,1"
 .Handshaking = comNone
 .RTSEnable = True
 .DTREnable = False
 .RThreshold = 1
 .InputLen = 0
 End With
 
If MSComm.PortOpen = True Then
Label3.Caption = "Current Status: " & regComPort & " is already Open."
Else
MSComm.PortOpen = True
Label3.Caption = "Current Status: " & regComPort & " opened ok."
End If
End Sub

Private Sub Command3_Click()
Label2.Caption = "ATR:"
Label1.Caption = "Atmel Code:"
Picture1.Cls
Picture1.Print ""
Picture1.FontSize = 10
Picture1.FontBold = True
Picture1.ForeColor = &HFFFFFF
Picture1.Print " Welcome to HU Reader."
Picture1.FontSize = 8
Picture1.FontBold = False
Picture1.ForeColor = &HFF00&
Picture1.Print " Version 1.0"
Picture1.FontSize = 7
MSComm.PortOpen = False
Label3.Caption = "Current Status: " & regComPort & " is closed."
Command1.Enabled = False
Command3.Enabled = False
End Sub

Private Sub Form_Load()
Picture1.Cls
Picture1.Print ""
Picture1.FontSize = 10
Picture1.FontBold = True
Picture1.ForeColor = &HFFFFFF
Picture1.Print " Welcome to HU Reader."
Picture1.FontSize = 8
Picture1.FontBold = False
Picture1.ForeColor = &HFF00&
Picture1.Print " Version 1.0"
Picture1.FontSize = 7
End Sub
Private Sub Form_Unload(Cancel As Integer)
If MSComm.PortOpen = True Then
WriteHU ("020200")
Call GreenOff
Call RedOff
MSComm.PortOpen = False
Unload Me
Else
Unload Me
End If
End Sub
