Attribute VB_Name = "basFunctions"
Option Explicit

Public gstrCommentStyle As String
Const NoRecentlyOpenedFiles = 5

'Always on Top
Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
Const SWP_NOMOVE = &H2
Const SWP_NOSIZE = &H1
Const SWP_NOACTIVATE = &H10
Const SWP_SHOWWINDOW = &H40
Const TOPMOST_FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Public Const WM_USER As Long = &H400
Public Const EM_FORMATRANGE As Long = WM_USER + 57

Declare Function SetWindowPos Lib "user32" _
    (ByVal hwnd As Long, _
    ByVal hWndInsertAfter As Long, _
    ByVal X As Long, _
    ByVal Y As Long, _
    ByVal cx As Long, _
    ByVal cy As Long, _
    ByVal wFlags As Long) As Long
    
Private Declare Function FindExecutable _
    Lib "shell32.dll" Alias "FindExecutableA" _
    (ByVal lpFile As String, _
    ByVal lpDirectory As String, _
    ByVal lpResult As String) As Long
    
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
     (ByVal hwnd As Long, ByVal msg As Long, ByVal wp As Long, Ip As Any) As Long
     
Public Function RunApp(strAppName As String) As Long
Dim strResult   As String
Dim lngResult   As Long
Dim i           As Variant
    
        strResult = String(255, 0)
        lngResult = FindExecutable(strAppName, vbNullString, strResult)
        strResult = Trim(Replace(strResult, "/dde", "", 1))
        
        'Run the file and not an .exe file
        i = Shell(Trim(Replace(strResult, vbNullChar, "", 1)) & " " & strAppName, 1)
    
End Function
     
Public Function AlwaysOnTop(hwnd As Long, blnAlwaysOnTop As Boolean)

    On Error Resume Next

    If blnAlwaysOnTop Then
        SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, TOPMOST_FLAGS
    Else
        SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, TOPMOST_FLAGS
    End If

End Function

'Public Function SaveToFile() As Boolean
'Dim rs              As ADODB.Recordset
'Dim strFilename     As String
'
'    On Error GoTo SaveToFile_Error
'
'    frmSQLMDI.CommonDialog1.FileName = "*.txt"
'    frmSQLMDI.CommonDialog1.DefaultExt = "txt"
'    frmSQLMDI.CommonDialog1.Filter = "Text (*.txt)|*.txt"
'    frmSQLMDI.CommonDialog1.ShowSave
'
'    frmSQLMDI.StatusBar1.Panels(1).Text = "Saving..."
'
'    If frmSQLMDI.CommonDialog1.FileName <> "*.txt" Then
'        Screen.MousePointer = vbHourglass
'        Set rs = frmSQLResults.DataGrid1.DataSource
'
'        If Not rs Is Nothing Then
'            rs.Filter = ""
'            strFilename = frmSQLMDI.CommonDialog1.FileName
'            rs.Save strFilename
'        End If
'    End If
'
'    frmSQLMDI.StatusBar1.Panels(1).Text = "Ready"
'
'SaveToFile_Exit:
'    Screen.MousePointer = vbNormal
'    Exit Function
'
'SaveToFile_Error:
'    MsgBox Err.Description, vbCritical, "SQL Writer"
'    Resume SaveToFile_Exit
'
'End Function
'
'Public Function OpenFile() As Boolean
'Dim rs      As New ADODB.Recordset
'
'    On Error GoTo OpenFile_Error
'
'    With frmSQLMDI.CommonDialog1
'        .DefaultExt = "*.*"
'        .FileName = "*.*"
'        .Filter = "ADO Recordset (*.*)"
'        .ShowOpen
'
'        frmSQLMDI.StatusBar1.Panels(1).Text = "Opening..."
'
'        If .FileName <> "*.*" Then
'            Screen.MousePointer = vbHourglass
'            rs.Open .FileName
'        End If
'
'    End With
'
'    Set frmSQLResults.DataGrid1.DataSource = rs
'    frmSQLResults.DataGrid1.Refresh
'
'    frmSQLMDI.StatusBar1.Panels(1).Text = "Ready"
'
'OpenFile_Exit:
'    Screen.MousePointer = vbNormal
'    Exit Function
'
'OpenFile_Error:
'    MsgBox Err.Description, vbCritical, "SQL Writer"
'    Resume OpenFile_Exit
'
'End Function
'
'Public Function RunApp(strAppName As String) As Long
'Dim strResult   As String
'Dim lngResult   As Long
'Dim i           As Variant
'
'        strResult = String(255, 0)
'        lngResult = FindExecutable(strAppName, vbNullString, strResult)
'        strResult = Trim(Replace(strResult, "/dde", "", 1))
'
'        'Run the file and not an .exe file
'        i = Shell(Trim(Replace(strResult, vbNullChar, "", 1)) & " " & strAppName, 1)
'
'End Function


