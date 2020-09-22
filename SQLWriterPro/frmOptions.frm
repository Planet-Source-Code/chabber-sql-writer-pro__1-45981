VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Options"
   ClientHeight    =   4560
   ClientLeft      =   2565
   ClientTop       =   1500
   ClientWidth     =   6150
   Icon            =   "frmOptions.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4560
   ScaleWidth      =   6150
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   3735
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   6588
      _Version        =   393216
      Style           =   1
      Tabs            =   5
      TabsPerRow      =   8
      TabHeight       =   520
      TabCaption(0)   =   "Startup Properties"
      TabPicture(0)   =   "frmOptions.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Editor"
      TabPicture(1)   =   "frmOptions.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame6"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Results"
      TabPicture(2)   =   "frmOptions.frx":0044
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame5"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Connections"
      TabPicture(3)   =   "frmOptions.frx":0060
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame4"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "General"
      TabPicture(4)   =   "frmOptions.frx":007C
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame3"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).ControlCount=   1
      Begin VB.Frame Frame6 
         Caption         =   "Editor:"
         Height          =   3015
         Left            =   -74805
         TabIndex        =   24
         Top             =   480
         Width           =   5535
         Begin VB.CheckBox Check2 
            Caption         =   "Use SQL Writer Pro as the default editor for files with .sql extension"
            Height          =   375
            Left            =   240
            TabIndex        =   32
            Top             =   1560
            Width           =   2895
         End
         Begin VB.CheckBox Check1 
            Caption         =   "&Word Wrap"
            Height          =   255
            Left            =   240
            TabIndex        =   31
            Top             =   1080
            Width           =   1215
         End
         Begin VB.TextBox txtTabSize 
            Height          =   285
            Left            =   2040
            TabIndex        =   30
            Text            =   "5"
            Top             =   600
            Width           =   975
         End
         Begin VB.Label lblTabSize 
            Caption         =   "&Tab size (in spaces):"
            Height          =   255
            Left            =   240
            TabIndex        =   29
            Top             =   600
            Width           =   1695
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Results:"
         Height          =   3015
         Left            =   -74805
         TabIndex        =   23
         Top             =   480
         Width           =   5535
         Begin VB.OptionButton Option2 
            Caption         =   "Select and &Sort Column"
            Height          =   255
            Left            =   840
            TabIndex        =   36
            Top             =   1800
            Width           =   2175
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Se&lect Column"
            Height          =   255
            Left            =   840
            TabIndex        =   35
            Top             =   1440
            Width           =   2295
         End
         Begin VB.CheckBox Check3 
            Caption         =   "Beep when a query completes"
            Height          =   255
            Left            =   360
            TabIndex        =   33
            Top             =   480
            Width           =   4095
         End
         Begin VB.Label Label6 
            Caption         =   "When a column header is clicked:"
            Height          =   255
            Left            =   360
            TabIndex        =   34
            Top             =   1080
            Width           =   3015
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "New Connections:"
         Height          =   3015
         Left            =   -74805
         TabIndex        =   22
         Top             =   480
         Width           =   5535
         Begin VB.CheckBox Check4 
            Caption         =   "Auto-Commit"
            Height          =   255
            Left            =   240
            TabIndex        =   39
            Top             =   1200
            Width           =   1815
         End
         Begin VB.TextBox Text2 
            Height          =   285
            Left            =   240
            TabIndex        =   38
            Text            =   "100"
            Top             =   600
            Width           =   1575
         End
         Begin VB.Label Label7 
            Caption         =   "&Cache Size:"
            Height          =   255
            Left            =   240
            TabIndex        =   37
            Top             =   360
            Width           =   1815
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "General:"
         Height          =   3015
         Left            =   -74805
         TabIndex        =   18
         Top             =   480
         Width           =   5535
         Begin VB.ComboBox cboCommentSel 
            Height          =   315
            ItemData        =   "frmOptions.frx":0098
            Left            =   2040
            List            =   "frmOptions.frx":00A5
            Style           =   2  'Dropdown List
            TabIndex        =   28
            Top             =   960
            Width           =   1455
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Left            =   2025
            TabIndex        =   20
            Text            =   "5"
            Top             =   465
            Width           =   1095
         End
         Begin VB.Label Label5 
            Caption         =   "Comment Selection:"
            Height          =   255
            Left            =   240
            TabIndex        =   27
            Top             =   960
            Width           =   1575
         End
         Begin VB.Label Label2 
            Caption         =   "files*"
            Height          =   255
            Left            =   3240
            TabIndex        =   21
            Top             =   480
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "Recent file list contains:"
            Height          =   255
            Left            =   240
            TabIndex        =   19
            Top             =   480
            Width           =   1815
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Other:"
         Height          =   1095
         Left            =   180
         TabIndex        =   15
         Top             =   2400
         Width           =   5535
         Begin VB.CheckBox chkHideResults 
            Caption         =   "Hide Results"
            Height          =   255
            Left            =   240
            TabIndex        =   17
            Top             =   600
            Width           =   1455
         End
         Begin VB.CheckBox chkHideSchema 
            Caption         =   "Hide Schema"
            Height          =   255
            Left            =   240
            TabIndex        =   16
            Top             =   360
            Width           =   1455
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Connection:"
         Height          =   1815
         Left            =   195
         TabIndex        =   10
         Top             =   480
         Width           =   5535
         Begin VB.CommandButton cmdRemove 
            Caption         =   "&Remove"
            Height          =   375
            Left            =   4500
            TabIndex        =   14
            Top             =   1155
            Width           =   855
         End
         Begin VB.CommandButton cmdBuild 
            Caption         =   "&Build"
            Height          =   375
            Left            =   4500
            TabIndex        =   13
            Top             =   675
            Width           =   855
         End
         Begin VB.TextBox txtDSNStartup 
            Height          =   1005
            Left            =   240
            MultiLine       =   -1  'True
            TabIndex        =   12
            Top             =   600
            Width           =   4095
         End
         Begin VB.Label lblDSNStartup 
            Caption         =   "*&Load connection string on startup:"
            Height          =   255
            Left            =   240
            TabIndex        =   11
            Top             =   360
            Width           =   2655
         End
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   3
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample4 
         Caption         =   "Sample 4"
         Height          =   1785
         Left            =   2100
         TabIndex        =   8
         Top             =   840
         Width           =   2055
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   2
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample3 
         Caption         =   "Sample 3"
         Height          =   1785
         Left            =   1545
         TabIndex        =   7
         Top             =   675
         Width           =   2055
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   1
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample2 
         Caption         =   "Sample 2"
         Height          =   1785
         Left            =   645
         TabIndex        =   6
         Top             =   300
         Width           =   2055
      End
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4920
      TabIndex        =   2
      Top             =   3975
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3720
      TabIndex        =   1
      Top             =   3975
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   2490
      TabIndex        =   0
      Top             =   3975
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "*"
      Height          =   255
      Left            =   120
      TabIndex        =   26
      Top             =   3960
      Width           =   135
   End
   Begin VB.Label Label3 
      Caption         =   "Change takes effect next time the application is started"
      Height          =   495
      Left            =   240
      TabIndex        =   25
      Top             =   3960
      Width           =   2295
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboCommentSel_Click()

    Select Case cboCommentSel.List(cboCommentSel.ListIndex)
        Case "Use '--' Style"
            gstrCommentStyle = "--"
        Case "Use '/**/' Style"
            gstrCommentStyle = "/**/"
        Case "Use ' Style"
            gstrCommentStyle = "'"
    End Select
    
    cmdApply.Enabled = True

End Sub

Private Sub cmdApply_Click()

    On Error Resume Next

    'Save to registry
    SaveSetting App.Title, "Settings", "Startup DSN", txtDSNStartup.Text
    SaveSetting App.Title, "Settings", "Comment Style", cboCommentSel.ListIndex
    SaveSetting App.Title, "Settings", "Hide Schema", chkHideSchema.value
    SaveSetting App.Title, "Settings", "Hide Results", chkHideResults.value
    
    cmdApply.Enabled = False

End Sub

Private Sub cmdBuild_Click()
Dim DAL As New clsSQLProDAL

    On Error Resume Next
    
    txtDSNStartup.Text = DAL.ConnectionWizard
    
End Sub

Private Sub cmdCancel_Click()

    On Error Resume Next
    
    Unload Me

End Sub

Private Sub cmdOK_Click()

    On Error Resume Next
    
    cmdApply_Click
    Unload Me

End Sub

Private Sub cmdRemove_Click()

    On Error Resume Next
    
    txtDSNStartup.Text = ""
    DeleteSetting App.Title, "Settings", "Startup DSN"

End Sub

Private Sub Form_Load()

    On Error Resume Next
    
    txtDSNStartup.Text = GetSetting(App.Title, "Settings", "Startup DSN", "")
    chkHideSchema.value = GetSetting(App.Title, "Settings", "Hide Schema", "0")
    chkHideResults.value = GetSetting(App.Title, "Settings", "Hide Results", "0")

End Sub

Private Sub txtDSNStartup_Change()

    On Error Resume Next
        
    cmdApply.Enabled = True

End Sub
