VERSION 5.00
Begin VB.Form Form5 
   BackColor       =   &H00808080&
   BorderStyle     =   0  'None
   ClientHeight    =   2910
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4380
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2910
   ScaleWidth      =   4380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   360
      Top             =   2280
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Press ESC to exit."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1335
   End
   Begin VB.Image imgFScreen 
      Height          =   2895
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4335
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyPress(KeyAscii As Integer)
    'Unload form when ESC is pressed
    If KeyAscii = 27 Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    'Resize image
    imgFScreen.Height = Form5.Height
    imgFScreen.Width = Form5.Width
End Sub

Private Sub Form_Resize()
    'Resize image
    imgFScreen.Height = Form5.Height
    imgFScreen.Width = Form5.Width
End Sub

Private Sub Timer1_Timer()
    'Copy image from imgCam every 1 second
    imgFScreen.Picture = Form1.imgCam.Picture
End Sub
