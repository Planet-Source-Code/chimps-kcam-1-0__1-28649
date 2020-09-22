VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7530
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   7425
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7530
   ScaleWidth      =   7425
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   240
      Top             =   3960
   End
   Begin VB.Frame frmDown 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   1800
      TabIndex        =   28
      Top             =   340
      Visible         =   0   'False
      Width           =   1815
      Begin VB.Image Image15 
         Height          =   180
         Left            =   0
         Picture         =   "Main.frx":030A
         Top             =   120
         Width           =   210
      End
      Begin VB.Label lblStatus 
         BackStyle       =   0  'Transparent
         Caption         =   "Downloading..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   29
         Top             =   120
         Width           =   1455
      End
   End
   Begin VB.Frame frmMain 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   3495
      Left            =   2040
      TabIndex        =   25
      Top             =   960
      Width           =   4575
      Begin MSComDlg.CommonDialog CommonDialog2 
         Left            =   480
         Top             =   2880
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         Filter          =   "BMP (*.bmp)|*.bmp"
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   0
         Top             =   2880
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         Filter          =   "KCam files (*.kcf)|*.kcf"
      End
      Begin VB.Image Image13 
         Height          =   1185
         Left            =   960
         Picture         =   "Main.frx":055C
         Top             =   240
         Width           =   2715
      End
      Begin VB.Image Image14 
         Height          =   480
         Left            =   480
         Picture         =   "Main.frx":AD7E
         Top             =   1680
         Width           =   480
      End
      Begin VB.Label Label23 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Click here to open a webcam database"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   855
         Left            =   1080
         MouseIcon       =   "Main.frx":B088
         MousePointer    =   99  'Custom
         TabIndex        =   27
         Top             =   1680
         Width           =   2895
      End
      Begin VB.Label Label24 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Show Cameras window"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   495
         Left            =   1560
         MouseIcon       =   "Main.frx":B952
         MousePointer    =   99  'Custom
         TabIndex        =   26
         Top             =   2760
         Visible         =   0   'False
         Width           =   1815
      End
   End
   Begin VB.Frame frmExtra 
      BorderStyle     =   0  'None
      Height          =   1800
      Left            =   4440
      TabIndex        =   23
      Top             =   5520
      Visible         =   0   'False
      Width           =   1300
      Begin VB.Label Label27 
         Caption         =   "Download more cameras"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         MouseIcon       =   "Main.frx":C21C
         MousePointer    =   99  'Custom
         TabIndex        =   24
         Top             =   240
         Width           =   1095
      End
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   135
      Left            =   120
      TabIndex        =   22
      ToolTipText     =   "Update meter"
      Top             =   5040
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Frame frmHelp 
      BorderStyle     =   0  'None
      Height          =   1800
      Left            =   1560
      TabIndex        =   15
      Top             =   5520
      Visible         =   0   'False
      Width           =   1300
      Begin VB.Label Label20 
         Caption         =   "&About"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         MouseIcon       =   "Main.frx":CAE6
         MousePointer    =   99  'Custom
         TabIndex        =   19
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Line Line9 
         BorderColor     =   &H00FFFFFF&
         X1              =   1300
         X2              =   0
         Y1              =   1320
         Y2              =   1320
      End
      Begin VB.Line Line8 
         BorderColor     =   &H00808080&
         X1              =   1300
         X2              =   0
         Y1              =   1335
         Y2              =   1335
      End
      Begin VB.Label Label19 
         Caption         =   "&History"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         MouseIcon       =   "Main.frx":D3B0
         MousePointer    =   99  'Custom
         TabIndex        =   18
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label18 
         Caption         =   "&Read me"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         MouseIcon       =   "Main.frx":DC7A
         MousePointer    =   99  'Custom
         TabIndex        =   17
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label15 
         Caption         =   "&Help"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         MouseIcon       =   "Main.frx":E544
         MousePointer    =   99  'Custom
         TabIndex        =   16
         Top             =   960
         Width           =   1095
      End
   End
   Begin VB.Frame frmViews 
      BorderStyle     =   0  'None
      Height          =   1800
      Left            =   3000
      TabIndex        =   10
      Top             =   5520
      Visible         =   0   'False
      Width           =   1300
      Begin VB.Line Line7 
         BorderColor     =   &H00FFFFFF&
         X1              =   1300
         X2              =   0
         Y1              =   1320
         Y2              =   1320
      End
      Begin VB.Line Line6 
         BorderColor     =   &H00808080&
         X1              =   1300
         X2              =   0
         Y1              =   1335
         Y2              =   1335
      End
      Begin VB.Label Label14 
         Caption         =   "&Full Screen"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         MouseIcon       =   "Main.frx":EE0E
         MousePointer    =   99  'Custom
         TabIndex        =   14
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label13 
         Caption         =   "&Small"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         MouseIcon       =   "Main.frx":F6D8
         MousePointer    =   99  'Custom
         TabIndex        =   13
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label12 
         Caption         =   "&Normal"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         MouseIcon       =   "Main.frx":FFA2
         MousePointer    =   99  'Custom
         TabIndex        =   12
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label11 
         Caption         =   "&Large"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         MouseIcon       =   "Main.frx":1086C
         MousePointer    =   99  'Custom
         TabIndex        =   11
         Top             =   960
         Width           =   1095
      End
   End
   Begin VB.Frame frmCameras 
      BorderStyle     =   0  'None
      Height          =   1800
      Left            =   120
      TabIndex        =   6
      Top             =   5520
      Visible         =   0   'False
      Width           =   1300
      Begin VB.Label Label17 
         Caption         =   "&Open new list"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         MouseIcon       =   "Main.frx":11136
         MousePointer    =   99  'Custom
         TabIndex        =   31
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Line Line13 
         BorderColor     =   &H00808080&
         X1              =   1300
         X2              =   0
         Y1              =   975
         Y2              =   975
      End
      Begin VB.Line Line12 
         BorderColor     =   &H00FFFFFF&
         X1              =   1300
         X2              =   0
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Label Label6 
         Caption         =   "&Modify cam"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         MouseIcon       =   "Main.frx":11A00
         MousePointer    =   99  'Custom
         TabIndex        =   9
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "&Stop Updating"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         MouseIcon       =   "Main.frx":122CA
         MousePointer    =   99  'Custom
         TabIndex        =   8
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "&Add cam"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         MouseIcon       =   "Main.frx":12B94
         MousePointer    =   99  'Custom
         TabIndex        =   7
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3855
      Left            =   1800
      ScaleHeight     =   3855
      ScaleWidth      =   5175
      TabIndex        =   30
      Top             =   720
      Width           =   5175
      Begin VB.Image imgCam 
         Height          =   3135
         Left            =   525
         Stretch         =   -1  'True
         Top             =   405
         Width           =   4095
      End
   End
   Begin VB.Image imgTempSave 
      Height          =   495
      Left            =   6120
      Stretch         =   -1  'True
      Top             =   5520
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   120
      MouseIcon       =   "Main.frx":1345E
      MousePointer    =   99  'Custom
      TabIndex        =   32
      Top             =   2445
      Width           =   1095
   End
   Begin VB.Image Image3 
      Height          =   345
      Left            =   120
      Picture         =   "Main.frx":13D28
      Top             =   2400
      Width           =   1110
   End
   Begin VB.Line Line11 
      BorderColor     =   &H00FFFFFF&
      X1              =   1300
      X2              =   0
      Y1              =   4920
      Y2              =   4920
   End
   Begin VB.Line Line10 
      BorderColor     =   &H00808080&
      X1              =   1300
      X2              =   0
      Y1              =   4935
      Y2              =   4935
   End
   Begin VB.Image Image11 
      Height          =   480
      Left            =   1440
      MouseIcon       =   "Main.frx":1518A
      MousePointer    =   99  'Custom
      Picture         =   "Main.frx":15A54
      ToolTipText     =   "Modify current cam"
      Top             =   4800
      Width           =   480
   End
   Begin VB.Line Line3 
      X1              =   1320
      X2              =   1320
      Y1              =   240
      Y2              =   5280
   End
   Begin VB.Label Label21 
      BackColor       =   &H00808080&
      Caption         =   "chimps@whoever.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5760
      MouseIcon       =   "Main.frx":15E96
      MousePointer    =   99  'Custom
      TabIndex        =   20
      ToolTipText     =   "Email the author"
      Top             =   5040
      Width           =   1695
   End
   Begin VB.Image Image6 
      Height          =   480
      Left            =   5160
      MouseIcon       =   "Main.frx":16760
      MousePointer    =   99  'Custom
      Picture         =   "Main.frx":1702A
      ToolTipText     =   "Email the author"
      Top             =   4800
      Width           =   480
   End
   Begin VB.Label Label22 
      BackColor       =   &H00808080&
      Caption         =   "Modify this camera"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2040
      MouseIcon       =   "Main.frx":17334
      MousePointer    =   99  'Custom
      TabIndex        =   21
      ToolTipText     =   "Modify current cam"
      Top             =   5040
      Width           =   1455
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   255
      Left            =   5040
      Top             =   5040
      Width           =   2415
   End
   Begin VB.Image Image10 
      Height          =   225
      Left            =   4680
      MouseIcon       =   "Main.frx":17BFE
      MousePointer    =   99  'Custom
      Picture         =   "Main.frx":184C8
      ToolTipText     =   "Copy camera image"
      Top             =   5055
      Width           =   240
   End
   Begin VB.Image Image9 
      Height          =   225
      Left            =   4320
      MouseIcon       =   "Main.frx":189FA
      MousePointer    =   99  'Custom
      Picture         =   "Main.frx":192C4
      ToolTipText     =   "Save camera image"
      Top             =   5055
      Width           =   240
   End
   Begin VB.Image Image8 
      Height          =   225
      Left            =   3960
      MouseIcon       =   "Main.frx":197F6
      MousePointer    =   99  'Custom
      Picture         =   "Main.frx":1A0C0
      ToolTipText     =   "Print camera image"
      Top             =   5055
      Width           =   240
   End
   Begin VB.Image Image7 
      Height          =   225
      Left            =   3600
      MouseIcon       =   "Main.frx":1A5F2
      MousePointer    =   99  'Custom
      Picture         =   "Main.frx":1AEBC
      ToolTipText     =   "Add new cam"
      Top             =   5055
      Width           =   240
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      Height          =   255
      Left            =   3480
      Top             =   5040
      Width           =   1575
   End
   Begin VB.Line Line5 
      X1              =   7440
      X2              =   1320
      Y1              =   5040
      Y2              =   5040
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   255
      Left            =   1320
      Top             =   5040
      Width           =   2175
   End
   Begin VB.Line Line4 
      X1              =   7440
      X2              =   1320
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      X1              =   1300
      X2              =   0
      Y1              =   3015
      Y2              =   3015
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   1300
      X2              =   0
      Y1              =   3000
      Y2              =   3000
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Extra"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   120
      MouseIcon       =   "Main.frx":1B3EE
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   1965
      Width           =   1095
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Help"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   120
      MouseIcon       =   "Main.frx":1BCB8
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   1485
      Width           =   1095
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Views"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   120
      MouseIcon       =   "Main.frx":1C582
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   1005
      Width           =   1095
   End
   Begin VB.Image Image5 
      Height          =   345
      Left            =   120
      Picture         =   "Main.frx":1CE4C
      Top             =   1920
      Width           =   1110
   End
   Begin VB.Image Image4 
      Height          =   345
      Left            =   120
      Picture         =   "Main.frx":1E2AE
      Top             =   1440
      Width           =   1110
   End
   Begin VB.Image Image2 
      Height          =   345
      Left            =   120
      Picture         =   "Main.frx":1F710
      Top             =   960
      Width           =   1110
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Cameras"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   120
      MouseIcon       =   "Main.frx":20B72
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   525
      Width           =   1095
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "x"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7200
      MouseIcon       =   "Main.frx":2143C
      MousePointer    =   99  'Custom
      TabIndex        =   2
      ToolTipText     =   "Close"
      Top             =   0
      Width           =   255
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00E0E0E0&
      BorderStyle     =   0  'Transparent
      BorderWidth     =   3
      Height          =   250
      Left            =   7200
      Top             =   0
      Width           =   255
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C00000&
      Caption         =   " KCam 1.0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7455
   End
   Begin VB.Image Image1 
      Height          =   345
      Left            =   120
      Picture         =   "Main.frx":21D06
      Top             =   480
      Width           =   1110
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      BorderStyle     =   0  'Transparent
      Height          =   5055
      Left            =   0
      Top             =   240
      Width           =   1335
   End
   Begin VB.Menu mnuViews 
      Caption         =   "&Views"
      Visible         =   0   'False
      Begin VB.Menu mnuViewsSmall 
         Caption         =   "&Small"
      End
      Begin VB.Menu mnuViewsNormal 
         Caption         =   "&Normal"
      End
      Begin VB.Menu mnuViewsLarge 
         Caption         =   "&Large"
      End
      Begin VB.Menu mnuViewsSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewsNewWindow 
         Caption         =   "&Full Screen"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim yFrame As Long            'Top position of frames
Dim xFrame As Long            'Left position of frames
Dim m_cIni As New cInifile    'ini file
Public SavedTitle As String   'Save current title to be used in other forms
Public Interval As Long       'Update interval
Public Address As String      'Address of web cam
Public TempPath As String     'Path where to download image
Public Path As String         'Path to load cam database


Private Sub Form_Load()
    'Resize Form
    Form1.Height = 5370
    Form1.Width = 7515
    
    'Set frame position
    xFrame = -1200
    yFrame = 3120
    
    'Arrange Frames
    frmCameras.Top = yFrame
    frmCameras.Left = xFrame

    frmViews.Top = yFrame
    frmViews.Left = xFrame

    frmHelp.Top = yFrame
    frmHelp.Left = xFrame
    
    frmExtra.Top = yFrame
    frmExtra.Left = xFrame
    
    'Center imgCam
    imgCam.Top = Picture1.Height / 2 - imgCam.Height / 2
    imgCam.Left = Picture1.Width / 2 - imgCam.Width / 2
    
    'Asign value to temporary file
    Form1.TempPath = App.Path & "\Temp.dat"
    
    'Stablish settings for CommonDialog1
    CommonDialog1.InitDir = App.Path
    CommonDialog1.DialogTitle = "Open KCam database"

    CommonDialog2.InitDir = App.Path & "\Saved Images"
    CommonDialog2.DialogTitle = "Save Camera image"
End Sub

Private Sub Form_Terminate()
    'Delete temporary files when the app is closed
    On Error Resume Next
    Kill App.Path & "\Temp.dat"
    Kill App.Path & "\Downloads.ini"
    Kill App.Path & "\TempDownloads.ini"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Delete temporary files when the app is closed
    On Error Resume Next
    Kill App.Path & "\Temp.dat"
    Kill App.Path & "\Downloads.ini"
    Kill App.Path & "\TempDownloads.ini"
End Sub

Private Sub Image10_Click()
    'Make sure there is an image loaded
    If frmMain.Visible = False Then
        'Save image to clipboard
        Clipboard.SetData imgCam.Picture
        'Inform user that the image has been saved
        MsgBox "The Image has been copied into the clipboard", vbInformation
    Else
        Exit Sub
    End If
End Sub

Private Sub Image11_Click()
    'Make sure the camera window is opened
    If Label24.Visible = False Then
        'Call ModifyCam Sub
        Call ModifyCam
        Form3.Modify = True
    Else
        Exit Sub
    End If
    
    'Stablish which image to show
    Form3.imgWhat.Picture = Form3.imgModify.Picture
End Sub

Private Sub Image6_Click()
    'Show email client app
    ShellExecute Me.hwnd, vbNullString, "mailto:chimps@whoever.com", vbNullString, "C:\", SW_SHOWNORMAL
End Sub

Private Sub Image7_Click()
    'Make sure a database is loaded
    If Form1.Path = "" Then
        Exit Sub
    Else
    Form3.Show
    
    'Show -new- image
    Form3.imgWhat.Picture = Form3.imgNew.Picture
    End If
End Sub

Private Sub Image8_Click()
    'Temporarily load it into another image control
    imgTempSave.Picture = imgCam.Picture
    
    'Print
    Printer.PaintPicture imgTempSave.Picture, 100, 100
End Sub

Private Sub Image9_Click()
    On Error Resume Next
    CommonDialog1.CancelError = True
    
    'Copy the image onto another image control so
    'it saves the image that was loaded at the moment
    imgTempSave.Picture = imgCam.Picture
    
    'Show Save dialog box
    CommonDialog2.ShowSave
    
    'When then user presses the cancel button
    If Err.Number = cdlCancel Then
        Exit Sub
    Else
        'Save camera
        SavePicture imgTempSave.Picture, CommonDialog2.FileName
    End If
End Sub

Private Sub imgCam_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Make sure a camera is selected
    If Form2.ListView1.SelectedItem <> "" Then
    
        'Popup views menu
        If Button = vbRightButton Then
            Form1.PopupMenu mnuViews
        End If
    End If
End Sub

Private Sub Label1_Click()

    'Make sure a database is loaded
    If Form1.Path = "" Then
        Exit Sub
    Else
    
    'Make the form go bak
    Call GoBack
    
    'Bring this frame forward forward
    frmCameras.ZOrder (vbOnTop)
    frmCameras.Visible = True
    
    'Slide camera
    Do While frmCameras.Left <> 0
        frmCameras.Left = frmCameras.Left + 1
        DoEvents  'Slow down for a while so the app won't freeze
    Loop
    End If
End Sub

Private Sub Label10_Click()
   
    'Make frames slid back in
    Call GoBack
    
    'Bring frame to front
    frmExtra.ZOrder (vbOnTop)
    frmExtra.Visible = True
    
    'Slide frame until its left property is 0
    Do While frmExtra.Left <> 0
        frmExtra.Left = frmExtra.Left + 1
        DoEvents
    Loop
End Sub

Private Sub Label11_Click()
    'Resize image
    imgCam.Height = Picture1.Height
    imgCam.Width = Picture1.Width
    
    'Center it
    imgCam.Top = Picture1.Height / 2 - imgCam.Height / 2
    imgCam.Left = Picture1.Width / 2 - imgCam.Width / 2
End Sub

Private Sub Label12_Click()
    'Resize
    imgCam.Height = 3135
    imgCam.Width = 4095
    
    'Center it
    imgCam.Top = Picture1.Height / 2 - imgCam.Height / 2
    imgCam.Left = Picture1.Width / 2 - imgCam.Width / 2
End Sub

Private Sub Label13_Click()
    'Resize image
    imgCam.Height = Picture1.Height / 2
    imgCam.Width = Picture1.Width / 2
    
    'Center it
    imgCam.Top = Picture1.Height / 2 - imgCam.Height / 2
    imgCam.Left = Picture1.Width / 2 - imgCam.Width / 2
End Sub

Private Sub Label14_Click()
    'make sura a camera is loaded
    If frmMain.Visible = False Then
        Form5.Show
        Form5.imgFScreen.Picture = Form1.imgCam.Picture 'Load image from one form to another
        Form5.Timer1.Enabled = True 'Make the timer in form2 start
    End If
End Sub

Private Sub Label15_Click()
    Form7.Show
    
    'Clear contents of Text1
    Form7.Text1.Text = ""
    
    'Open file
    Open App.Path & "\Help.txt" For Input As #1
    
    'Read each line until end of file
    Do Until EOF(1)
        Line Input #1, Data
        
        'write each line to Text1
        Form7.Text1.Text = Form7.Text1.Text & Data & vbCrLf
    Loop
    
    Close #1
End Sub

Private Sub Label17_Click()
    On Error Resume Next
    
    'When the user presses the cancel button
    CommonDialog1.CancelError = True
    
    CommonDialog1.ShowOpen
    
    'When then user presses the cancel button
    If Err.Number = cdlCancel Then
        Exit Sub
    Else
        'Apply path to the variable
        Form1.Path = CommonDialog1.FileName
        
        'Show camera window
        Form2.Show
        
        'Load New file
        Call LoadFile
        
        'Stick one form to another
        Call MoveCamsForm
        
        'Status
        Label24.Visible = False
        
        'Stop Timer 1 (updating)
        Timer1.Enabled = False
        ProgressBar1.Value = 0  'Stablish normal position
    End If
End Sub

Private Sub Label18_Click()
    Form7.Show
    
    'Clean text field
    Form7.Text1.Text = ""
    
    'Open file
    Open App.Path & "\Readme.txt" For Input As #1
    Do Until EOF(1)
        Line Input #1, Data
        Form7.Text1.Text = Form7.Text1.Text & Data & vbCrLf
    Loop
    
    Close #1
End Sub

Private Sub Label19_Click()
    Form7.Show
    Form7.Text1.Text = ""
    Open App.Path & "\History.txt" For Input As #1
    Do Until EOF(1)
        Line Input #1, Data
        Form7.Text1.Text = Form7.Text1.Text & Data & vbCrLf
    Loop
    
    Close #1
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Move form
    FormDrag Me
    
    Call MoveCamsForm
End Sub

Private Sub Label20_Click()
    Form4.Show
End Sub

Private Sub Label21_Click()
    'Show email client app
    ShellExecute Me.hwnd, vbNullString, "mailto:chimps@whoever.com", vbNullString, "C:\", SW_SHOWNORMAL
End Sub

Private Sub Label22_Click()
    'Make sure the camera window is opened
    If Label24.Visible = False Then
        Call ModifyCam
        Form3.Modify = True
    Else
        Exit Sub
    End If
    
    'Show the modify icon
    Form3.imgWhat.Picture = Form3.imgModify.Picture
End Sub

Private Sub Label23_Click()
    On Error Resume Next
    CommonDialog1.CancelError = True
    
    CommonDialog1.ShowOpen
    
    'When then user presses the cancel button
    If Err.Number = cdlCancel Then
        Exit Sub
    Else
        Form1.Path = CommonDialog1.FileName
        Call LoadFile
        Call MoveCamsForm
        
        'Show camera windows
        Label24_Click
    End If
End Sub

Private Sub Label24_Click()
    If Form1.Path = "" Then
        Exit Sub
    Else
        Form2.Show
        Call MoveCamsForm
        Label24.Visible = False
    End If
End Sub

Private Sub Label27_Click()
    Form6.Show
End Sub

Private Sub Label3_Click()
    End
End Sub

Private Sub Label4_Click()
    Form3.Show
    Form3.imgWhat.Picture = Form3.imgNew.Picture
End Sub

Private Sub Label5_Click()
    'Restablish normal position
    Timer1.Enabled = False
    ProgressBar1.Value = 0
End Sub

Private Sub Label6_Click()
    'Make sure the camera window is opened
    If Label24.Visible = False Then
        Call ModifyCam
        Form3.Modify = True
    Else
        Exit Sub
    End If
    
    Form3.imgWhat.Picture = Form3.imgModify.Picture
End Sub

Private Sub Label7_Click()
    'Make sure a database is loaded
    If Form1.Path = "" Then
        Exit Sub
    Else
    Call GoBack
    frmViews.ZOrder (vbOnTop)
    frmViews.Visible = True
    Do While frmViews.Left <> 0
        frmViews.Left = frmViews.Left + 1
        DoEvents
    Loop
    End If
End Sub



Private Sub Label8_Click()
    End
End Sub

Private Sub Label9_Click()
    Call GoBack
    frmHelp.ZOrder (vbOnTop)
    frmHelp.Visible = True
    Do While frmHelp.Left <> 0
        frmHelp.Left = frmHelp.Left + 1
        DoEvents
    Loop
End Sub

Private Sub GoBack()
    'Make the frames go back in
    frmCameras.Left = xFrame
    frmViews.Left = xFrame
    frmHelp.Left = xFrame
    frmExtra.Left = xFrame
End Sub

Private Sub MoveCamsForm()
    'Glue form 2 to form1
    Form2.Left = Form1.Left - Form2.Width
    Form2.Top = Form1.Top
End Sub

Private Sub mnuViewsLarge_Click()
    'Resize image
    imgCam.Height = Picture1.Height
    imgCam.Width = Picture1.Width
    
    'Center it
    imgCam.Top = Picture1.Height / 2 - imgCam.Height / 2
    imgCam.Left = Picture1.Width / 2 - imgCam.Width / 2
End Sub

Private Sub mnuViewsNewWindow_Click()
    If frmMain.Visible = False Then
        Form5.Show
        Form5.imgFScreen.Picture = Form1.imgCam.Picture
        Form5.Timer1.Enabled = True
    End If
End Sub

Private Sub mnuViewsNormal_Click()
    'Resize
    imgCam.Height = 3135
    imgCam.Width = 4095
    
    'Center it
    imgCam.Top = Picture1.Height / 2 - imgCam.Height / 2
    imgCam.Left = Picture1.Width / 2 - imgCam.Width / 2
End Sub

Private Sub mnuViewsSmall_Click()
    'Resize image
    imgCam.Height = Picture1.Height / 2
    imgCam.Width = Picture1.Width / 2
    
    'Center it
    imgCam.Top = Picture1.Height / 2 - imgCam.Height / 2
    imgCam.Left = Picture1.Width / 2 - imgCam.Width / 2
End Sub

Private Sub Timer1_Timer()
    On Error Resume Next
    'Increment the value of ProgressBar1 every second (1000 ms)
    Form1.ProgressBar1.Value = Form1.ProgressBar1.Value + 1
    
    'If the value has reached its maximum amount, call reload
    If ProgressBar1.Value >= ProgressBar1.Max Then
        Call Reload
    End If
End Sub

Public Sub Reload()
        'Status
        Form1.lblStatus.Caption = "Downloading!"
                
        'Delete file
        Kill App.Path & "\Temp.dat"
        
        'Refresh camera
        imgCam.Refresh
        
        'Download image from the web
        DoEvents
        Call DownloadFile(Address, Form1.TempPath)
        DoEvents
        
        'Display downloaded image in imgCam and refresh it
        imgCam.Picture = LoadPicture(Form1.TempPath)
        imgCam.Refresh
        
        'Set value to 0 and start counting until interval
        ProgressBar1.Value = 0
        Form1.lblStatus.Caption = "Ready!"
End Sub
