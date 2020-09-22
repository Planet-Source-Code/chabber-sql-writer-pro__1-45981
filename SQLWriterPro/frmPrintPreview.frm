VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPrintPreview 
   BackColor       =   &H8000000C&
   Caption         =   "Print Preview"
   ClientHeight    =   8430
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7815
   Icon            =   "frmPrintPreview.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8430
   ScaleWidth      =   7815
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      LargeChange     =   1440
      Left            =   0
      Max             =   500
      SmallChange     =   1440
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   7920
      Visible         =   0   'False
      Width           =   7575
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   7575
      LargeChange     =   1440
      Left            =   7560
      Max             =   500
      SmallChange     =   1440
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   360
      Visible         =   0   'False
      Width           =   270
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   8175
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12171
            Text            =   "Page 1"
            TextSave        =   "Page 1"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Object.Width           =   1058
            MinWidth        =   1058
            TextSave        =   "NUM"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3000
      Top             =   9240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrintPreview.frx":27A2
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrintPreview.frx":2CE4
            Key             =   "Setup"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   635
      ButtonWidth     =   1138
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      Begin VB.CommandButton cmdZoomOut 
         Caption         =   "Zoom Out"
         Enabled         =   0   'False
         Height          =   255
         Left            =   5400
         TabIndex        =   7
         Top             =   30
         Width           =   975
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "Close"
         Height          =   255
         Left            =   6720
         TabIndex        =   6
         Top             =   30
         Width           =   975
      End
      Begin VB.CommandButton cmdZoomIn 
         Caption         =   "Zoom In"
         Height          =   255
         Left            =   4080
         TabIndex        =   5
         Top             =   30
         Width           =   975
      End
      Begin VB.CommandButton cmdPrevPage 
         Caption         =   "Prev Page"
         Enabled         =   0   'False
         Height          =   255
         Left            =   1440
         TabIndex        =   4
         Top             =   30
         Width           =   975
      End
      Begin VB.CommandButton cmdNextPage 
         Caption         =   "Next Page"
         Enabled         =   0   'False
         Height          =   255
         Left            =   2760
         TabIndex        =   3
         Top             =   30
         Width           =   975
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "Print..."
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   30
         Width           =   975
      End
   End
   Begin VB.PictureBox picTarget 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   7031
      Left            =   1200
      ScaleHeight     =   7005
      ScaleWidth      =   5025
      TabIndex        =   0
      Top             =   720
      Width           =   5051
   End
End
Attribute VB_Name = "frmPrintPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Enum Orientation
    Portrait = 1
    Landscape = 2
End Enum

Public CurrentPage As Integer

'Public variables
Public gintZoomPercentage As Integer
Public gPageOrientation As Orientation

Private Sub cmdClose_Click()
    
    Unload Me

End Sub

Private Sub cmdNextPage_Click()

    CurrentPage = CurrentPage + 1
    SetNextPrevious
    
    RichTextToPictureBox frmParent.ActiveForm.txtSQL, _
                         picTarget, _
                         frmParent.ActiveForm.picSQLText, _
                         frmParent.ActiveForm.gMarginLeft, _
                         frmParent.ActiveForm.gMarginRight, _
                         frmParent.ActiveForm.gMarginTop, _
                         frmParent.ActiveForm.gMarginBottom, _
                         CurrentPage
                         
    StatusBar1.Panels(1).Text = "Page " & CurrentPage & " of " & gNumPages

End Sub

Private Sub cmdPrevPage_Click()

    CurrentPage = CurrentPage - 1
    SetNextPrevious

    RichTextToPictureBox frmParent.ActiveForm.txtSQL, _
                         picTarget, _
                         frmParent.ActiveForm.picSQLText, _
                         frmParent.ActiveForm.gMarginLeft, _
                         frmParent.ActiveForm.gMarginRight, _
                         frmParent.ActiveForm.gMarginTop, _
                         frmParent.ActiveForm.gMarginBottom, _
                         CurrentPage
                         
    
    
End Sub

Private Sub cmdPrint_Click()

    Unload Me
    frmParent.ActiveForm.PrintSQL

End Sub

Private Sub cmdZoomIn_Click()
    
    picTarget.Width = cnstFullPreviewWidth * 2
    picTarget.Height = cnstFullPreviewHeight * 2
    
    Form_Resize
    
    If picTarget.Width > (Me.Width - VScroll1.Width) Then
        HScroll1.Visible = True
    Else
        HScroll1.Visible = False
    End If
    
    If picTarget.Height > (Me.Height - HScroll1.Height - StatusBar1.Height - Toolbar1.Height - 520) Then
        VScroll1.Visible = True
    Else
        VScroll1.Visible = False
    End If
    
    cmdZoomIn.Enabled = False
    cmdZoomOut.Enabled = True
    
    ResizePreviewToPictureControl picTarget, frmParent.ActiveForm.picSQLText

End Sub

Private Sub cmdZoomOut_Click()
    
    picTarget.Width = cnstFullPreviewWidth
    picTarget.Height = cnstFullPreviewHeight
    
    Form_Resize
    
    HScroll1.Visible = False
    VScroll1.Visible = False
    
    cmdZoomIn.Enabled = True
    cmdZoomOut.Enabled = False
    
    ResizePreviewToPictureControl picTarget, frmParent.ActiveForm.picSQLText

End Sub

Private Sub Form_Activate()

    SetNextPrevious

End Sub

Private Sub SetNextPrevious()

    'Turn on/off next page button
    If gNumPages > 1 And CurrentPage <> gNumPages Then
        cmdNextPage.Enabled = True
    Else
        cmdNextPage.Enabled = False
    End If
    
    'Turn on/off previous page button
    If CurrentPage > 1 Then
        cmdPrevPage.Enabled = True
    Else
        cmdPrevPage.Enabled = False
    End If

End Sub
Private Sub Form_Load()

    HScroll1.Min = -120
    HScroll1.Max = (cnstFullPreviewWidth * 2) / 3
    HScroll1.Value = -120
    
    VScroll1.Min = -480
    VScroll1.Max = (cnstFullPreviewHeight * 2) / 2
    VScroll1.Value = -480
    
    StatusBar1.Panels(1).Text = "Page " & CurrentPage & " of " & gNumPages
    
End Sub

Private Sub Form_Resize()
Dim sngAppSpace As Single
Dim sngTopPos   As Single
Dim sngLeftPos  As Single

    'Capture application space
    sngAppSpace = Me.Height - HScroll1.Height - StatusBar1.Height - Toolbar1.Height - 520

    'Position and size scrollbars
    VScroll1.Left = Me.Width - 375
    HScroll1.Top = Me.Height - 1020
    HScroll1.Width = Me.Width - VScroll1.Width - 100
    VScroll1.Height = Abs(sngAppSpace)
    
    'Position preview page
    sngLeftPos = ((Me.Width - VScroll1.Width) / 2) - (picTarget.Width / 2)
    If sngLeftPos > 0 Then
        picTarget.Left = sngLeftPos
    Else
        picTarget.Left = 120
    End If
    
    sngTopPos = (sngAppSpace / 2) - (picTarget.Height / 2)
    If sngTopPos > 480 Then
        picTarget.Top = sngTopPos
    Else
        picTarget.Top = 480
    End If
    
    'Hide/Show Scrollbars
    If picTarget.Width > Me.Width Then
        HScroll1.Visible = True
    Else
        HScroll1.Visible = False
    End If
    
    If picTarget.Height > (Me.Height - HScroll1.Height - StatusBar1.Height - Toolbar1.Height - 520) Then
        VScroll1.Visible = True
    Else
        VScroll1.Visible = False
    End If

End Sub

Private Sub HScroll1_Change()

    picTarget.Left = -HScroll1.Value

End Sub

Private Sub HScroll1_Scroll()

    HScroll1_Change
    
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.Key
        Case "Setup"
            ViewPageSetup

    End Select
    
End Sub

Private Sub ModifyPageDisplay()
'
'Draws the Page in either landscape or portrait mode and resizes the
'preview according one page or two
'

    Select Case gPageOrientation
        Case Orientation.Landscape
            
        Case Else
            
    End Select

End Sub

Private Sub VScroll1_Change()

    picTarget.Top = -VScroll1.Value

End Sub

Private Sub VScroll1_Scroll()

    VScroll1_Change

End Sub
