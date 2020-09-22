VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4065
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   4545
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4065
   ScaleWidth      =   4545
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   960
      TabIndex        =   10
      Text            =   "10"
      Top             =   3120
      Width           =   975
   End
   Begin VB.TextBox txtComments 
      Height          =   525
      Left            =   480
      MultiLine       =   -1  'True
      TabIndex        =   7
      Top             =   2160
      Width           =   3855
   End
   Begin VB.TextBox txtAddress 
      Height          =   285
      Left            =   480
      TabIndex        =   5
      Top             =   1440
      Width           =   3855
   End
   Begin VB.TextBox txtTitle 
      Height          =   285
      Left            =   480
      TabIndex        =   3
      Top             =   720
      Width           =   3855
   End
   Begin VB.Image imgWhat 
      Height          =   225
      Left            =   80
      Stretch         =   -1  'True
      Top             =   360
      Width           =   240
   End
   Begin VB.Image imgModify 
      Height          =   255
      Left            =   480
      Picture         =   "camopt.frx":0000
      Stretch         =   -1  'True
      Top             =   3600
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "&Cancel"
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
      Left            =   2040
      MouseIcon       =   "camopt.frx":0442
      MousePointer    =   99  'Custom
      TabIndex        =   13
      Top             =   3645
      Width           =   1095
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFFFFF&
      Caption         =   "seconds"
      Height          =   255
      Left            =   2040
      TabIndex        =   12
      Top             =   3150
      Width           =   735
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Every"
      Height          =   255
      Left            =   480
      TabIndex        =   11
      Top             =   3150
      Width           =   615
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Update:"
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
      Left            =   480
      TabIndex        =   9
      Top             =   2880
      Width           =   735
   End
   Begin VB.Image imgNew 
      Height          =   225
      Left            =   840
      Picture         =   "camopt.frx":0D0C
      Stretch         =   -1  'True
      Top             =   3600
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "&Okay"
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
      Left            =   3240
      MouseIcon       =   "camopt.frx":123E
      MousePointer    =   99  'Custom
      TabIndex        =   8
      Top             =   3645
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   345
      Left            =   3240
      Picture         =   "camopt.frx":1B08
      Top             =   3600
      Width           =   1110
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Comments:"
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
      Left            =   480
      TabIndex        =   6
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Address:"
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
      Left            =   480
      TabIndex        =   4
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Camera title:"
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
      Left            =   480
      TabIndex        =   2
      Top             =   480
      Width           =   1215
   End
   Begin VB.Line Line1 
      X1              =   240
      X2              =   4560
      Y1              =   240
      Y2              =   240
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
      Left            =   4320
      MouseIcon       =   "camopt.frx":2F6A
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   0
      Width           =   255
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00E0E0E0&
      BorderStyle     =   0  'Transparent
      BorderWidth     =   3
      Height          =   255
      Left            =   4320
      Top             =   0
      Width           =   255
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C00000&
      Caption         =   " Add / Modify camera"
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
      Width           =   4575
   End
   Begin VB.Image Image3 
      Height          =   345
      Left            =   2040
      Picture         =   "camopt.frx":3834
      Top             =   3600
      Width           =   1110
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      Height          =   4095
      Left            =   0
      Top             =   0
      Width           =   375
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim m_cIni As New cInifile
Public Modify As Boolean

Private Sub Form_Load()
    'Add Items to Combo1
    With Combo1
        .AddItem "5"
        .AddItem "10"
        .AddItem "15"
        .AddItem "20"
        .AddItem "25"
        .AddItem "30"
        .AddItem "35"
        .AddItem "40"
        .AddItem "45"
        .AddItem "50"
        .AddItem "55"
        .AddItem "60"
    End With
End Sub

Private Sub Label10_Click()
    Unload Me
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    FormDrag Me
End Sub

Private Sub Label3_Click()
    Unload Me
End Sub

Private Sub Label6_Click()
If txtTitle.Text = "" Or txtAddress.Text = "" Or txtComments.Text = "" Then
    MsgBox "Please fill in all fields."
Else
'Not modyfing...adding new camera
If Modify = False Then
    With m_cIni
        .Path = Form1.Path
        .Section = txtTitle.Text
        .Key = "Cam Title"
        .Value = txtTitle
        
        .Key = "Address"
        .Value = txtAddress.Text
        
        .Key = "Comments"
        .Value = txtComments.Text
        
        .Key = "Interval"
        .Value = Combo1.Text
        
        If Not (.Success) Then
            MsgBox "There is no database available to save the camera into."
        End If
    End With
    
    Call LoadFile
    Unload Me
End If

'Modifying camera
If Modify = True Then
    With m_cIni
    
        'Delete section and write it again (modifying)
        .Path = Form1.Path
        .Section = Form1.SavedTitle
        .DeleteSection
        
        .Section = txtTitle.Text
        
        .Key = "Cam Title"
        .Value = txtTitle.Text
        
        .Key = "Address"
        .Value = txtAddress.Text
        
        .Key = "Comments"
        .Value = txtComments.Text
        
        .Key = "Interval"
        .Value = Combo1.Text
        
        If Not (.Success) Then
            MsgBox "There is no database available to save the camera into."
        End If
    End With
    
    'Load camera database into ListView
    Call LoadFile
    Unload Me
End If
End If
End Sub
