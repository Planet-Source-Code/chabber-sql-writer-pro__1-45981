VERSION 5.00
Begin VB.Form frmFind 
   Caption         =   "Find"
   ClientHeight    =   1560
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6240
   Icon            =   "frmFind.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1560
   ScaleWidth      =   6240
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkMatchCase 
      Caption         =   "Match &case"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1110
      Width           =   1815
   End
   Begin VB.CheckBox chkMatchWholeWord 
      Caption         =   "Match &whole word only"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   2055
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4680
      TabIndex        =   3
      Top             =   720
      Width           =   1335
   End
   Begin VB.CommandButton cmdFindNext 
      Caption         =   "&Find Next"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   375
      Left            =   4680
      TabIndex        =   2
      Top             =   240
      Width           =   1335
   End
   Begin VB.TextBox txtFind 
      Height          =   285
      Left            =   1080
      TabIndex        =   1
      Top             =   240
      Width           =   3375
   End
   Begin VB.Label lblFind 
      Caption         =   "Fi&nd what:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   855
   End
End
Attribute VB_Name = "frmFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Private module-level variables
Private lngStart  As Long
Private lngEnd    As Long

Private Sub cmdCancel_Click()
    
    Unload Me
    
End Sub

Private Sub cmdFindNext_Click()
Dim pad As frmChild
Dim lngOptions As Long

    'Set reference to active child
    Set pad = frmParent.ActiveForm

    'Set find options
    If chkMatchWholeWord.Value Then
        If chkMatchCase.Value Then
            lngOptions = rtfWholeWord Or rtfMatchCase
        Else
            lngOptions = rtfWholeWord
        End If
    Else
        If chkMatchCase.Value Then
            lngOptions = rtfMatchCase
        Else
            lngOptions = 0
        End If
    End If
    
    'Determine start and end
    If lngStart <= 0 Then
        lngStart = 1
        lngEnd = Len(pad.txtSQL.Text)
    End If
    
    'Find instance of search text
    lngStart = pad.txtSQL.Find(txtFind.Text, lngStart + 1, lngEnd, lngOptions)
    
    If lngStart = -1 Then
        lngStart = 0
        MsgBox "SQL Writer Pro has finished searching the query text.", vbInformation
    End If
    
End Sub

Private Sub txtFind_KeyUp(KeyCode As Integer, Shift As Integer)

    If txtFind.Text <> "" Then
        cmdFindNext.Enabled = True
    Else
        cmdFindNext.Enabled = False
    End If
    
End Sub
