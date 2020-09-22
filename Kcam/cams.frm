VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form2 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5280
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   2055
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7530
   ScaleMode       =   0  'User
   ScaleWidth      =   2055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1440
      Top             =   4440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   15
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "cams.frx":0000
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   4382
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   7726
      View            =   2
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.TextBox txtComments 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   4414
      Width           =   2055
   End
   Begin VB.Line Line2 
      X1              =   -360
      X2              =   2040
      Y1              =   6267.869
      Y2              =   6267.869
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   2040
      Y1              =   7187.727
      Y2              =   7187.727
   End
   Begin VB.Label Label2 
      Caption         =   "Modify"
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
      Left            =   1440
      MouseIcon       =   "cams.frx":0324
      MousePointer    =   99  'Custom
      TabIndex        =   2
      ToolTipText     =   "Modify this camera"
      Top             =   5040
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Delete cam"
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
      Left            =   240
      MouseIcon       =   "cams.frx":0BEE
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   5040
      Width           =   855
   End
   Begin VB.Image Image2 
      Height          =   240
      Left            =   1175
      MouseIcon       =   "cams.frx":14B8
      MousePointer    =   99  'Custom
      Picture         =   "cams.frx":1D82
      Stretch         =   -1  'True
      ToolTipText     =   "Modify this camera"
      Top             =   5040
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   225
      Left            =   0
      MouseIcon       =   "cams.frx":21C4
      MousePointer    =   99  'Custom
      Picture         =   "cams.frx":2A8E
      Top             =   5040
      Width           =   240
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      Height          =   255
      Left            =   0
      Top             =   5040
      Width           =   2175
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim lstxList1 As ListItem
Dim m_cIni As New cInifile

Private Sub Form_Load()
    On Error Resume Next
    Call LoadFile
    
    'Load file and show comments
    With m_cIni
        .Path = Form1.Path
        .Section = ListView1.SelectedItem.Text
        .Key = "Comments"
        txtComments.Text = .Value
    End With
End Sub


Private Sub Image1_Click()
    Call Delete
End Sub

Private Sub Image2_Click()
    'Make sure the camera window is opened
    If Form1.Label24.Visible = False Then
        Call ModifyCam
        Form3.Modify = True
    Else
        Exit Sub
    End If
    
    Form3.imgWhat.Picture = Form3.imgModify.Picture
End Sub

Private Sub Label1_Click()
    Call Delete
End Sub

Private Sub Label2_Click()
    'Make sure the camera window is opened
    If Form1.Label24.Visible = False Then
        Call ModifyCam
        Form3.Modify = True
    Else
        Exit Sub
    End If
    
    Form3.imgWhat.Picture = Form3.imgModify.Picture
End Sub

Private Sub ListView1_DblClick()
    'Status
    Form1.lblStatus.Caption = "Downloading!"
    DoEvents
    Form1.frmMain.Visible = False
    Form1.frmDown.Visible = True
    
    'Get address from .dat file and download camera
    With m_cIni
        .Path = Form1.Path
        .Section = ListView1.SelectedItem.Text
        
        .Key = "Address"
        
        'Download camera
        DoEvents
        Down = DownloadFile(.Value, Form1.TempPath)
        DoEvents
        Form1.Address = .Value
        
        'Load camera
        Form1.imgCam.Picture = LoadPicture(App.Path & "\Temp.dat")
        Form1.lblStatus.Caption = "Ready!"
        
        'Stop current updating
        Form1.Timer1.Enabled = False
        Form1.ProgressBar1.Value = 0
        
        'Get interval from file
        .Key = "Interval"
        Form1.Interval = .Value
        
        Form1.Timer1.Enabled = True
        Form1.ProgressBar1.Max = Form1.Interval + 1
    End With
End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    'Show comments of selected item
    With m_cIni
        .Path = Form1.Path
        .Section = ListView1.SelectedItem.Text
        .Key = "Comments"
        txtComments.Text = .Value
    End With
    
End Sub

Private Sub Delete()
'Ask if the user wants to delete
If MsgBox("Are you sure you want to delete this webcam?", vbYesNo) = vbYes Then

    'Delete section
    With m_cIni
        .Path = Form1.Path
        .Section = ListView1.SelectedItem.Text
        .DeleteSection
        Call LoadFile
    End With
Else
    Exit Sub
End If
End Sub
