Attribute VB_Name = "Module1"
Dim m_cIni As New cInifile

Declare Function ShellExecute Lib "shell32.dll" Alias _
"ShellExecuteA" (ByVal hwnd As Long, _
ByVal lpOperation As String, _
ByVal lpFile As String, _
ByVal lpParameters As String, _
ByVal lpDirectory As String, _
ByVal nShowCmd As Long) As Long

Const SW_SHOWNORMAL = 1

'Download image from the internet and save it into a file
Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" _
(ByVal pCaller As Long, _
ByVal szURL As String, _
ByVal szFileName As String, _
ByVal dwReserved As Long, _
ByVal lpfnCB As Long) As Long

'Drag Form
Declare Function ReleaseCapture Lib "user32" () As Long
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Const WM_NCLBUTTONDOWN = &HA1
Const HTCAPTION = 2

'Function that uses the URLDownloadToFile API
Public Function DownloadFile(Url As String, LocalFileName As String) As Boolean
    Dim Value As Long
    Form1.lblStatus.Caption = "Downloading!"
    Value = URLDownloadToFile(0, Url, LocalFileName, 0, 0)
    DoEvents   'A brief pause so the app won't freeze
End Function

Public Sub FormDrag(TheForm As Form)
    ReleaseCapture
    Call SendMessage(TheForm.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
End Sub

Public Sub LoadFile()
    'Declare variables
    Dim sSections() As String
    Dim iSectionCount As Long
    
    'Clear ListVIew before adding any new items
    Form2.ListView1.ListItems.Clear
    
    With m_cIni
        .Path = Form1.Path   'Open file
        .EnumerateAllSections sSections(), iSectionCount
        For iSection = 1 To iSectionCount
            'lstIni.AddItem "[" & sSections(iSection) & "]"
            .Section = sSections(iSection)
            .Key = "Address" 'Get items from this key only
            
            'Put new items into listview
            Set lstxList1 = Form2.ListView1.ListItems.Add(, , Trim(sSections(iSection)), , 1)
        Next iSection
    End With
End Sub

Public Sub ModifyCam()
    Form3.Show
    
    'Open same section as selected item in ListView and show its contents
    With m_cIni
        .Path = Form1.Path
        .Section = Form2.ListView1.SelectedItem.Text
        Form1.SavedTitle = Form2.ListView1.SelectedItem.Text
        
        .Key = "Cam Title"
        Form3.txtTitle.Text = .Value
        
        .Key = "Address"
        Form3.txtAddress.Text = .Value
        
        .Key = "Comments"
        Form3.txtComments.Text = .Value
        
        .Key = "Interval"
        Form3.Combo1.Text = .Value
        
    End With
End Sub
