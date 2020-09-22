VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmProperties 
   ClientHeight    =   3855
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   3630
   ControlBox      =   0   'False
   Icon            =   "frmProperties.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3855
   ScaleWidth      =   3630
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   3615
      Left            =   0
      TabIndex        =   2
      Top             =   240
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   6376
      _Version        =   393216
      FixedCols       =   0
      BackColorBkg    =   16777215
      HighLight       =   2
      ScrollBars      =   2
      AllowUserResizing=   1
      Appearance      =   0
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ï"
      BeginProperty Font 
         Name            =   "Wingdings 2"
         Size            =   9
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   3360
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   "Close"
      Top             =   15
      Width           =   240
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000002&
      Caption         =   "Properties"
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
      Height          =   210
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3615
   End
End
Attribute VB_Name = "frmProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mbMove As Boolean
Private Const ScrollBarWidth = 292

Public Function DisplayProperties(colFields As ADODB.Fields) As Boolean
Dim intCount    As Integer

    With MSFlexGrid1
        .ColWidth(0) = 1440
        .Row = 0
        .Col = 0
        .Text = "Property"
        .Col = 1
        .Text = "Value"
    End With

    For intCount = 0 To colFields.count - 1
        With MSFlexGrid1
            .AddItem colFields(intCount).Name & vbTab & colFields(intCount).value
        End With
    Next

    MSFlexGrid1.RemoveItem 1
    Me.Show vbModeless
    AlwaysOnTop Me.hwnd, True

End Function

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Resize()

    With MSFlexGrid1
        .Left = 0
        .Width = Me.Width - 136
        .Height = Me.Height - Label1.Height - 135
        .ColWidth(1) = .Width - 1440 - ScrollBarWidth
    End With
    
    Label1.Width = Me.Width - 136
    Command1.Left = Label1.Width - 255
    
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mbMove = True
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If mbMove Then
        Me.Move Me.Left + X, Me.Top + Y
    End If
    
End Sub

Private Sub Label1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mbMove = False
End Sub
