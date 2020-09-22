VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.MDIForm frmParent 
   BackColor       =   &H8000000C&
   Caption         =   "SQL Writer Pro"
   ClientHeight    =   9510
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   10140
   Icon            =   "frmParent.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   9255
      Width           =   10140
      _ExtentX        =   17886
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   8996
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   4048
            Text            =   "Total number of connections: 0"
            TextSave        =   "Total number of connections: 0"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   1402
            MinWidth        =   1411
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1402
            MinWidth        =   1411
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   1402
            MinWidth        =   1411
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   360
      Top             =   5400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "New Connection"
      End
      Begin VB.Menu mnuFileSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileOpenQuery 
         Caption         =   "Open Query"
      End
      Begin VB.Menu mnuFileSaveQuery 
         Caption         =   "Save Query"
      End
      Begin VB.Menu mnuFileSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "Print"
      End
      Begin VB.Menu mnuFilePrintPreview 
         Caption         =   "Print Preview"
      End
      Begin VB.Menu mnuFilePageSetup 
         Caption         =   "Page Setup"
      End
      Begin VB.Menu mnuFileSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMRUFiles 
         Caption         =   "Recent File 1"
         Index           =   1
      End
      Begin VB.Menu mnuMRUFiles 
         Caption         =   "Recent File 2"
         Index           =   2
      End
      Begin VB.Menu mnuMRUFiles 
         Caption         =   "Recent File 3"
         Index           =   3
      End
      Begin VB.Menu mnuMRUFiles 
         Caption         =   "Recent File 4"
         Index           =   4
      End
      Begin VB.Menu mnuMRUFiles 
         Caption         =   "Recent File 5"
         Index           =   5
      End
      Begin VB.Menu mnuMRUFiles 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditSelectAll 
         Caption         =   "Select All"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuEditSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditFind 
         Caption         =   "Find..."
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuEditSep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditAdvanced 
         Caption         =   "Advanced"
         Begin VB.Menu mnuEditAdvancedSelUpperCase 
            Caption         =   "Convert Selection To UpperCase"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuEditAdvancedSelLowerCase 
            Caption         =   "Convert Selection To LowerCase"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuEditAdvancedSep1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuEditAdvancedIncreaseIndent 
            Caption         =   "Increase Indent                                 TAB"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuEditAdvancedDecreaseIndent 
            Caption         =   "Decrease Indent                                SHIFT+TAB"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuEditAdvancedSep2 
            Caption         =   "-"
         End
         Begin VB.Menu mnuEditAdvancedComment 
            Caption         =   "Comment Selection"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuEditAdvancedUncomment 
            Caption         =   "Uncomment Selection"
            Enabled         =   0   'False
         End
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuViewSchema 
         Caption         =   "Schema"
      End
      Begin VB.Menu mnuViewResults 
         Caption         =   "Results"
      End
      Begin VB.Menu mnuViewSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewProperties 
         Caption         =   "Properties"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnuQuery 
      Caption         =   "&Query"
      Begin VB.Menu mnuQueryExecute 
         Caption         =   "Execute"
         Shortcut        =   {F5}
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuODBCManager 
         Caption         =   "Microsoft ODBC Manager"
      End
      Begin VB.Menu mnuToolsSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuToolsOptions 
         Caption         =   "Options"
      End
   End
   Begin VB.Menu mnuWin 
      Caption         =   "&Window"
      WindowList      =   -1  'True
      Begin VB.Menu mnuWinCascade 
         Caption         =   "Cascasde"
      End
      Begin VB.Menu mnuWinTileHorizontally 
         Caption         =   "Tile Horizontally"
      End
      Begin VB.Menu mnuWinTileVertically 
         Caption         =   "Tile Vertically"
      End
      Begin VB.Menu mnuWinArrangeIcons 
         Caption         =   "Arrange Icons"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpAlwaysOnTop 
         Caption         =   "Always On Top"
      End
      Begin VB.Menu mnuHelpADO 
         Caption         =   "ADO & SQL Help"
      End
      Begin VB.Menu mnuHelpSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "frmParent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private QueryPad() As frmChild
Private mintChildNo As Integer

Public Property Let NumberOfChildren(intValue As Integer)

    mintChildNo = intValue
    
    Me.StatusBar1.Panels(2).Text = "Total number of connections: " & mintChildNo
    
    If mintChildNo = 0 Then
        SetMenuNoChildren
    Else
        SetMenuChildren
    End If

End Property

Public Property Get NumberOfChildren() As Integer

    NumberOfChildren = mintChildNo

End Property

Private Sub MDIForm_Load()
Dim strStartupCnstr As String

    On Error Resume Next
    
    modMRUF.Number = 5
    modMRUF.Load App.Title
    
    strStartupCnstr = GetSetting(App.Title, "Settings", "Startup DSN", "")
    
    If strStartupCnstr <> "" Then
        OpenNewChild strStartupCnstr
    Else
        SetMenuNoChildren
        modMRUF.Update Me
    End If

    LoadFormProperties
    
End Sub

Private Sub LoadFormProperties()
Dim strTop      As String
Dim strLeft     As String
Dim strWidth    As String
Dim strHeight   As String
Dim strMaximized As String

    'Grab from registry
    strTop = GetSetting(App.Title, "Settings", "Top")
    strLeft = GetSetting(App.Title, "Settings", "Left")
    strWidth = GetSetting(App.Title, "Settings", "Width")
    strHeight = GetSetting(App.Title, "Settings", "Height")
    strMaximized = GetSetting(App.Title, "Settings", "Maximized")
    frmOptions.cboCommentSel.ListIndex = Val(GetSetting(App.Title, "Settings", "Comment Style", ""))
    
    
    'All must be set to use
    If (strTop <> "" And strLeft <> "" And strWidth <> "" And strHeight <> "") Or strMaximized = "1" Then
        
        If strMaximized = "1" Then
            Me.WindowState = vbMaximized
        Else
            'Apply settings
            If strTop > 0 Then Me.Top = strTop
            If strLeft > 0 Then Me.Left = strLeft
            Me.Width = strWidth
            Me.Height = strHeight
        End If
        
    End If

End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim intCount    As Integer

    'Unload All child windows
    For intCount = 1 To mintChildNo
        Set QueryPad(mintChildNo) = Nothing
    Next
    
    'Save any settings
    If Me.WindowState = vbMaximized Then
        SaveSetting App.Title, "Settings", "Maximized", "1"
    Else
        SaveSetting App.Title, "Settings", "Maximized", "0"
        SaveSetting App.Title, "Settings", "Top", Me.Top
        SaveSetting App.Title, "Settings", "Left", Me.Left
        SaveSetting App.Title, "Settings", "Width", Me.Width
        SaveSetting App.Title, "Settings", "Height", Me.Height
    End If
    
    'Unload All forms
    Unload frmOptions
    Unload frmFind
    Unload frmPrintPreview
    Unload frmAbout
    
End Sub

Private Sub mnuEditAdvancedComment_Click()

    Me.ActiveForm.CommentText

End Sub

Private Sub mnuEditAdvancedDecreaseIndent_Click()

    Me.ActiveForm.DecreaseIndent
    
End Sub

Private Sub mnuEditAdvancedIncreaseIndent_Click()

    Me.ActiveForm.IncreaseIndent

End Sub

Private Sub mnuEditAdvancedSelLowerCase_Click()

    Me.ActiveForm.MakeLowercase
    
End Sub

Private Sub mnuEditAdvancedSelUpperCase_Click()
    
    Me.ActiveForm.MakeUppercase

End Sub

Private Sub mnuEditAdvancedUncomment_Click()

    Me.ActiveForm.UnCommentText

End Sub

Private Sub mnuEditCopy_Click()

    Me.ActiveForm.CopySQLText

End Sub

Private Sub mnuEditCut_Click()

    Me.ActiveForm.CutSQLText

End Sub

Private Sub mnuEditFind_Click()

    frmFind.Show vbModeless

End Sub

Private Sub mnuEditPaste_Click()

    Me.ActiveForm.PasteSQLText

End Sub

Private Sub mnuFileExit_Click()

    End

End Sub

Private Sub mnuFileNew_Click()

    OpenNewChild

End Sub

Public Sub OpenNewChild(Optional strConnection As String = "")

    On Error Resume Next
    
        'Open child window
        mintChildNo = mintChildNo + 1
        ReDim QueryPad(mintChildNo)

        Set QueryPad(mintChildNo) = New frmChild
        
        'Grab connection string
        If strConnection = "" Then
            strConnection = QueryPad(mintChildNo).DAL.ConnectionWizard
        End If
        
        QueryPad(mintChildNo).mstrConnection = strConnection
        QueryPad(mintChildNo).Show
        QueryPad(mintChildNo).Caption = Me.ActiveForm.conn.Properties("Data Source") & " (" & Me.ActiveForm.conn.Properties("DBMS Name") & ") - " & "Untitled" & mintChildNo
    
        SetMenuChildren
        Me.StatusBar1.Panels(2).Text = "Total number of connections: " & mintChildNo

End Sub

Private Sub mnuFileOpenQuery_Click()
Dim strFilename As String

    With CommonDialog1
        .FileName = "*.qry"
        .Filter = "Queries | *.qry"
        .DefaultExt = "qry"
        .ShowOpen
        strFilename = .FileName
    End With

    If strFilename <> "*.qry" Then
        Me.ActiveForm.txtSQL.LoadFile strFilename
        Me.ActiveForm.Caption = strFilename
    End If

End Sub

Private Sub mnuFilePageSetup_Click()

    ViewPageSetup

End Sub

Private Sub mnuFilePrint_Click()

    Me.ActiveForm.PrintSQL

End Sub

Private Sub mnuFilePrintPreview_Click()

    Me.ActiveForm.PreviewSQL

End Sub

Private Sub mnuFileSaveQuery_Click()
Dim strFilename As String

    With CommonDialog1
        .FileName = "*.qry"
        .Filter = "Queries | *.qry"
        .DefaultExt = "qry"
        .ShowSave
        strFilename = .FileName
    End With
    
    If strFilename <> "*.qry" Then
        Me.ActiveForm.txtSQL.SaveFile strFilename
        Me.ActiveForm.Caption = strFilename
    End If

End Sub

Private Sub mnuHelpADO_Click()
    
    RunApp App.Path & "\Ado210.chm"
    
End Sub

Private Sub mnuHelpAlwaysOnTop_Click()

    mnuHelpAlwaysOnTop.Checked = Not mnuHelpAlwaysOnTop.Checked
    AlwaysOnTop Me.hwnd, mnuHelpAlwaysOnTop.Checked
    
End Sub

Private Sub mnuMRUFiles_Click(Index As Integer)
Dim strConnectionString As String

    Screen.MousePointer = vbHourglass
        strConnectionString = modMRUF.ITem(Right(mnuMRUFiles(Index).Caption, Len(mnuMRUFiles(Index).Caption) - 2))
        OpenNewChild strConnectionString
    Screen.MousePointer = vbNormal
    
End Sub

Private Sub mnuQueryExecute_Click()

    Me.ActiveForm.DAL.BindGridWithData Me.ActiveForm, Me.ActiveForm.conn

End Sub

Private Sub mnuViewProperties_Click()

    Me.ActiveForm.ViewColumnProperties

End Sub

Private Sub mnuViewResults_Click()

    mnuViewResults.Checked = Not mnuViewResults.Checked
    Me.ActiveForm.Toolbar1.Buttons("HideResults").value = IIf(Me.ActiveForm.Toolbar1.Buttons("HideResults").value = 1, 0, 1)
    Me.ActiveForm.HideResults Not mnuViewResults.Checked

End Sub

Private Sub mnuViewSchema_Click()

    mnuViewSchema.Checked = Not mnuViewSchema.Checked
    Me.ActiveForm.Toolbar1.Buttons("HideDatabase").value = IIf(Me.ActiveForm.Toolbar1.Buttons("HideDatabase").value = 1, 0, 1)
    Me.ActiveForm.HideDatabase Not mnuViewSchema.Checked
    
End Sub

Private Sub mnuWinTileHorizontally_Click()

    On Error Resume Next

    Me.Arrange vbTileHorizontal

End Sub

Private Sub mnuWinTileVertically_Click()

    On Error Resume Next

    Me.Arrange vbTileVertical

End Sub

Private Sub mnuFilePrintSetup_Click()

    On Error Resume Next
    
    CommonDialog1.flags = &H40
    CommonDialog1.ShowPrinter

End Sub

Private Sub mnuHelpAbout_Click()

    On Error Resume Next

    frmAbout.Show vbModal

End Sub

Private Sub mnuODBCManager_Click()
Dim retVal
    
    On Error Resume Next
    
    retVal = Shell("rundll32.exe shell32.dll,Control_RunDLL odbccp32.cpl,,3", 1)
    
End Sub

Private Sub mnuToolsOptions_Click()

    On Error Resume Next

    frmOptions.Show vbModal

End Sub

Private Function SetMenuNoChildren()
'
'Enables/Disables appropriate menu options
'
    
    On Error Resume Next
    
    'File menu
    mnuFileSaveQuery.Enabled = False
    mnuFileOpenQuery.Enabled = False
    mnuFilePrint.Enabled = False
    mnuFilePrintPreview.Enabled = False
    mnuFilePageSetup.Enabled = False
    
    'Edit menu
    mnuEditSelectAll.Enabled = False
    mnuEditFind.Enabled = False
    mnuEditAdvanced.Enabled = False
    
    'View menu
    mnuViewSchema.Enabled = False
    mnuViewResults.Enabled = False
    mnuViewProperties.Enabled = False
    
    'Query menu
    mnuQueryExecute.Enabled = False
    
    'Window menu
    mnuWinCascade.Enabled = False
    mnuWinTileHorizontally.Enabled = False
    mnuWinTileVertically.Enabled = False
    mnuWinArrangeIcons.Enabled = False

End Function

Private Function SetMenuChildren()
'
'Enables/Disables appropriate menu options
'
    
    On Error Resume Next
    
    'File menu
    mnuFileSaveQuery.Enabled = True
    mnuFileOpenQuery.Enabled = True
    mnuFilePrint.Enabled = True
    mnuFilePrintPreview.Enabled = True
    mnuFilePageSetup.Enabled = True
    
    'Edit menu
    mnuEditSelectAll.Enabled = True
    mnuEditFind.Enabled = True
    mnuEditAdvanced.Enabled = True
    
    'View menu
    mnuViewSchema.Enabled = True
    mnuViewResults.Enabled = True
    mnuViewProperties.Enabled = True
    
    'Query menu
    mnuQueryExecute.Enabled = True
    
    'Window menu
    mnuWinCascade.Enabled = True
    mnuWinTileHorizontally.Enabled = True
    mnuWinTileVertically.Enabled = True
    mnuWinArrangeIcons.Enabled = True

End Function































