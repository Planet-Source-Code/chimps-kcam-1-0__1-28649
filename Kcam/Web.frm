VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form6 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4065
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   5370
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4065
   ScaleWidth      =   5370
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   720
      Top             =   1440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "KCam files (*.kcf)|*.kcf"
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   1935
      Left            =   480
      TabIndex        =   2
      ToolTipText     =   "Double click to download"
      Top             =   2040
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   3413
      View            =   3
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      GridLines       =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Camera"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Description"
         Object.Width           =   6703
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "URL"
         Object.Width           =   18
      EndProperty
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   2880
      TabIndex        =   4
      Top             =   1590
      Width           =   1410
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "&Start"
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
      Left            =   1680
      MouseIcon       =   "Web.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   3
      ToolTipText     =   "Click here to see if there any new cameras available to download"
      Top             =   1605
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   1185
      Left            =   1560
      Picture         =   "Web.frx":08CA
      Top             =   360
      Width           =   2715
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
      Left            =   5160
      MouseIcon       =   "Web.frx":B0EC
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   0
      Width           =   255
   End
   Begin VB.Line Line1 
      X1              =   240
      X2              =   4200
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00E0E0E0&
      BorderStyle     =   0  'Transparent
      BorderWidth     =   3
      Height          =   255
      Left            =   5160
      Top             =   0
      Width           =   255
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C00000&
      Caption         =   "Download more cameras"
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
      TabIndex        =   1
      Top             =   0
      Width           =   5535
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      Height          =   4095
      Left            =   0
      Top             =   0
      Width           =   375
   End
   Begin VB.Image Image5 
      Height          =   345
      Left            =   1680
      Picture         =   "Web.frx":B9B6
      Top             =   1560
      Width           =   1110
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim lstLst As ListItem
Dim sSections() As String
Dim iSectionCount As Long
Dim m_cIni As New cInifile

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Kill App.Path & "\TempDownloads.ini"
End Sub

Private Sub Label11_Click()
    On Error Resume Next
    'Status
    Label1.Caption = "Checking..."
    Call DownloadFile("http://greendayband.tripod.com/Downloads.ini", App.Path & "\TempDownloads.ini")
    DoEvents

    'Clear ListVIew before adding any new items
    ListView1.ListItems.Clear
    
    With m_cIni
        .Path = App.Path & "\TempDownloads.ini"   'Open file
        .EnumerateAllSections sSections(), iSectionCount
        For iSection = 1 To iSectionCount
        
            'ListView1.AddItem "[" & sSections(iSection) & "]"
            .Section = sSections(iSection)
            
            'Put new items into listview
            Set lstLst = ListView1.ListItems.Add(, , Trim(sSections(iSection)))
        
            'Check Description Key and add it to ListView1
            .Key = "Description"
            lstLst.SubItems(1) = .Value
            
            'Check URL Key and add it to ListView1
            .Key = "URL"
            lstLst.SubItems(2) = .Value
        Next iSection
    End With
    
    Label1.Caption = "Done!"
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    FormDrag Me
End Sub

Private Sub Label3_Click()
    Unload Me
End Sub
Private Sub ListView1_DblClick()
    On Error Resume Next
    CommonDialog1.CancelError = True
        
    'Show Save dialog box
    CommonDialog1.ShowSave
    
    'When then user presses the cancel button
    If Err.Number = cdlCancel Then
        Exit Sub
    Else
        DoEvents
        'Download and save camera file
        Dow = URLDownloadToFile(0, Form6.ListView1.SelectedItem.SubItems(2), CommonDialog1.FileName, 0, 0)
        DoEvents
    End If
End Sub
