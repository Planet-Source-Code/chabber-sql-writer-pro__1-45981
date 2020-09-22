VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmChild 
   Caption         =   "Untitled1"
   ClientHeight    =   10710
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12660
   Icon            =   "frmChild.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10710
   ScaleWidth      =   12660
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList imglstDatabases 
      Left            =   6240
      Top             =   5760
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
            Picture         =   "frmChild.frx":030A
            Key             =   "Database"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChild.frx":06A4
            Key             =   "Not Supported"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imglstTreeView 
      Left            =   5040
      Top             =   5400
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   19
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChild.frx":0A3E
            Key             =   "tables"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChild.frx":0DD8
            Key             =   "Database"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChild.frx":1172
            Key             =   "views"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChild.frx":150C
            Key             =   "system_tables"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChild.frx":18A6
            Key             =   "synonyms"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChild.frx":1C40
            Key             =   "temporary_tables"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChild.frx":1FDA
            Key             =   "functions"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChild.frx":2374
            Key             =   "TABLE"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChild.frx":270E
            Key             =   "VIEW"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChild.frx":2AA8
            Key             =   "PROCEDURE"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChild.frx":2E42
            Key             =   "SYNONYM"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChild.frx":31DC
            Key             =   "SYSTEM TABLE"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChild.frx":3576
            Key             =   "GLOBAL TEMPORARY"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChild.frx":3910
            Key             =   "TABLE KEY"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChild.frx":3CAA
            Key             =   "TABLE COLUMN"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChild.frx":4044
            Key             =   "Relationships"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChild.frx":43DE
            Key             =   "Columns"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChild.frx":4778
            Key             =   "Indexes"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChild.frx":4B12
            Key             =   "Constraints"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imglstToolbar 
      Left            =   6720
      Top             =   3960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   24
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChild.frx":4EAC
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChild.frx":53EE
            Key             =   "Database_Pane"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChild.frx":5788
            Key             =   "Results_Pane"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChild.frx":5B22
            Key             =   "Comment"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChild.frx":5EBC
            Key             =   "New"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChild.frx":63FE
            Key             =   "UnComment"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChild.frx":6798
            Key             =   "Indent"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChild.frx":6B32
            Key             =   "Outdent"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChild.frx":6ECC
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChild.frx":731E
            Key             =   "Options"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChild.frx":7478
            Key             =   "Font"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChild.frx":7652
            Key             =   "Color"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChild.frx":77AC
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChild.frx":7CEE
            Key             =   "Redo"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChild.frx":8230
            Key             =   "Cut"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChild.frx":8772
            Key             =   "Paste"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChild.frx":8CB4
            Key             =   "Execute"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChild.frx":8FCE
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChild.frx":9510
            Key             =   "Copy"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChild.frx":9A52
            Key             =   "Undo"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChild.frx":9F94
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChild.frx":C746
            Key             =   "Properties"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChild.frx":CAE0
            Key             =   "Eraser"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChild.frx":CE7A
            Key             =   "Stop"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picSQLText 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   7440
      ScaleHeight     =   675
      ScaleWidth      =   1035
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   2520
      Visible         =   0   'False
      Width           =   1095
   End
   Begin ComCtl3.CoolBar CoolBar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   12660
      _ExtentX        =   22331
      _ExtentY        =   741
      BandCount       =   2
      _CBWidth        =   12660
      _CBHeight       =   420
      _Version        =   "6.7.8988"
      MinHeight1      =   360
      Width1          =   1395
      NewRow1         =   0   'False
      MinWidth2       =   1395
      MinHeight2      =   360
      Width2          =   2805
      NewRow2         =   0   'False
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   120
         TabIndex        =   7
         Top             =   45
         Width           =   12420
         _ExtentX        =   21908
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         Style           =   1
         ImageList       =   "imglstToolbar"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   27
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "New"
               Object.ToolTipText     =   "Open duplicate connection"
               ImageKey        =   "New"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Open"
               Object.ToolTipText     =   "Open a saved connection"
               ImageKey        =   "Open"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Save"
               Object.ToolTipText     =   "Save a connection"
               ImageKey        =   "Save"
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Print"
               Object.ToolTipText     =   "Print"
               ImageKey        =   "Print"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Preview"
               Object.ToolTipText     =   "Print Preview"
               ImageKey        =   "Preview"
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Key             =   "Cut"
               Object.ToolTipText     =   "Cut"
               ImageKey        =   "Cut"
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Key             =   "Copy"
               Object.ToolTipText     =   "Copy"
               ImageKey        =   "Copy"
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Paste"
               Object.ToolTipText     =   "Paste"
               ImageKey        =   "Paste"
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Find"
               Object.ToolTipText     =   "Find"
               ImageKey        =   "Find"
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Color"
               Object.ToolTipText     =   "Color"
               ImageKey        =   "Color"
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Font"
               Object.ToolTipText     =   "Font"
               ImageKey        =   "Font"
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Key             =   "Properties"
               Object.ToolTipText     =   "Properties"
               ImageKey        =   "Properties"
            EndProperty
            BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Key             =   "Indent"
               Object.ToolTipText     =   "Indent Text"
               ImageKey        =   "Indent"
            EndProperty
            BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Key             =   "Outdent"
               Object.ToolTipText     =   "Outdent Text"
               ImageKey        =   "Outdent"
            EndProperty
            BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Key             =   "Comment"
               Object.ToolTipText     =   "Comment Text"
               ImageKey        =   "Comment"
            EndProperty
            BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Key             =   "UnComment"
               Object.ToolTipText     =   "UnComment Text"
               ImageKey        =   "UnComment"
            EndProperty
            BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "HideDatabase"
               Object.ToolTipText     =   "Hide/Show Schema Pane"
               ImageKey        =   "Database_Pane"
               Style           =   1
            EndProperty
            BeginProperty Button23 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "HideResults"
               Object.ToolTipText     =   "Hide/Show Results Pane"
               ImageKey        =   "Results_Pane"
               Style           =   1
            EndProperty
            BeginProperty Button24 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button25 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Execute"
               Object.ToolTipText     =   "Execute SQL (F5)"
               ImageKey        =   "Execute"
            EndProperty
            BeginProperty Button26 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Eraser"
               Object.ToolTipText     =   "Clear contents of query window"
               ImageKey        =   "Eraser"
            EndProperty
            BeginProperty Button27 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
         EndProperty
         Begin MSComctlLib.ImageCombo cboCatalogs 
            Height          =   330
            Left            =   7800
            TabIndex        =   11
            Top             =   0
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   582
            _Version        =   393216
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            Text            =   "ImageCombo1"
            ImageList       =   "imglstDatabases"
         End
      End
   End
   Begin VB.PictureBox hSplitter 
      Align           =   2  'Align Bottom
      Height          =   30
      Left            =   0
      MousePointer    =   7  'Size N S
      ScaleHeight     =   30
      ScaleWidth      =   12660
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   8190
      Width           =   12660
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   10455
      Width           =   12660
      _ExtentX        =   22331
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7938
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   706
            MinWidth        =   706
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7938
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      Align           =   3  'Align Left
      BackColor       =   &H00FFFFFF&
      Height          =   7770
      Left            =   0
      ScaleHeight     =   7710
      ScaleWidth      =   3030
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   420
      Width           =   3090
      Begin MSComctlLib.TreeView tvDB 
         Height          =   7695
         Left            =   0
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   0
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   13573
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   176
         Style           =   7
         FullRowSelect   =   -1  'True
         ImageList       =   "imglstTreeView"
         Appearance      =   0
      End
   End
   Begin VB.PictureBox vSplitter 
      Align           =   3  'Align Left
      Height          =   7770
      Left            =   3090
      MousePointer    =   9  'Size W E
      ScaleHeight     =   7770
      ScaleWidth      =   45
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   420
      Width           =   45
   End
   Begin VB.PictureBox Picture2 
      Align           =   2  'Align Bottom
      BackColor       =   &H00FFFFFF&
      Height          =   2235
      Left            =   0
      ScaleHeight     =   2175
      ScaleWidth      =   12600
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   8220
      Width           =   12660
      Begin MSDataGridLib.DataGrid grdResults 
         Height          =   2175
         Left            =   0
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   0
         Width           =   12615
         _ExtentX        =   22251
         _ExtentY        =   3836
         _Version        =   393216
         AllowUpdate     =   -1  'True
         BackColor       =   16777215
         BorderStyle     =   0
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   2
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   9
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   4080
         Top             =   720
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
   End
   Begin RichTextLib.RichTextBox txtSQL 
      Height          =   7770
      Left            =   3120
      TabIndex        =   0
      Top             =   405
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   13705
      _Version        =   393217
      BorderStyle     =   0
      HideSelection   =   0   'False
      ScrollBars      =   3
      AutoVerbMenu    =   -1  'True
      OLEDragMode     =   0
      OLEDropMode     =   1
      TextRTF         =   $"frmChild.frx":D414
   End
End
Attribute VB_Name = "frmChild"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Module-Level Variables
Private i                   As Integer
Private sngOldVLocation     As Single
Private sngOldHLocation     As Single
Public mstrConnection       As String
Private mblnGroupExpanded   As Boolean
Private db                  As New ADOX.Catalog

Public conn                 As ADODB.Connection
Public gMarginTop           As Long
Public gMarginLeft          As Long
Public gMarginRight         As Long
Public gMarginBottom        As Long

Public DAL As New clsSQLProDAL

'Constants
Private Const intCaptionBarHeight = 620

Private Sub SetPanes()

        If GetSetting(App.Title, "Settings", "Hide Schema", "0") = 1 Then
            frmParent.mnuViewSchema.Checked = False
            Toolbar1.Buttons("HideDatabase").value = 0
            HideDatabase True
        Else
            frmParent.mnuViewSchema.Checked = True
            Toolbar1.Buttons("HideDatabase").value = 1
            HideDatabase False
        End If
    
        If GetSetting(App.Title, "Settings", "Hide Results", "0") = 1 Then
            frmParent.mnuViewResults.Checked = False
            Toolbar1.Buttons("HideResults").value = 0
            HideResults True
        Else
            frmParent.mnuViewResults.Checked = True
            Toolbar1.Buttons("HideResults").value = 1
            HideResults False
        End If
        
       ResizeControls
    
End Sub

Private Sub cboCatalogs_Click()
    
    Screen.MousePointer = vbHourglass
    If cboCatalogs.Text <> "" Then
        conn.Properties("Current Catalog") = cboCatalogs.Text
        tvDB.Nodes.Clear
        AddTableGroupsToTV
    End If
    Screen.MousePointer = vbNormal

End Sub

Private Sub Form_Activate()
    
    frmParent.mnuViewResults.Checked = CBool(Me.Toolbar1.Buttons("HideResults").value)
    frmParent.mnuViewSchema.Checked = CBool(Me.Toolbar1.Buttons("HideDatabase").value)
    
End Sub

Private Sub Form_Load()
Dim strDataSource As String

    On Error GoTo Form_Load_Error

    'Open connection to database
    Set conn = New ADODB.Connection
    conn.ConnectionString = mstrConnection
    conn.Properties("Persist Security Info") = True
    conn.Properties("Prompt") = adPromptComplete
    mstrConnection = Replace(mstrConnection, "Persist Security Info=False", "Persist Security Info=True")
    conn.Open , , , adConnectUnspecified
    mstrConnection = conn.ConnectionString
    db.ActiveConnection = mstrConnection
    
    'Set recently opened
    strDataSource = conn.Properties("Data Source") & " (" & conn.Properties("DBMS Name") & ")"
    modMRUF.Add strDataSource, conn.ConnectionString
    modMRUF.Save App.Title
    modMRUF.Update frmParent
    
    'Populate database tree
    AddTableGroupsToTV
    AddCatalogs
    
    'Set status
    Me.StatusBar1.Panels(1).Text = "Ready"
    
    'Set default margins
    gMarginTop = 1
    gMarginLeft = 1
    gMarginBottom = 1
    gMarginRight = 1

    sngOldVLocation = txtSQL.Height
    sngOldHLocation = txtSQL.Left
    
    Call SetPanes: DoEvents

Form_Load_EXIT:
    Exit Sub

Form_Load_Error:
    MsgBox Err.Description, vbCritical
    Unload Me
    Resume Form_Load_EXIT

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    frmParent.NumberOfChildren = frmParent.NumberOfChildren - 1

End Sub

Private Sub Form_Resize()

    ResizeControls

End Sub

Public Function ResizeControls()

    On Error Resume Next
    
    If Picture1.Visible Then
        txtSQL.Width = Abs(Me.Width - Picture1.Width - 100)
    Else
        txtSQL.Width = Abs(Me.Width - 100)
    End If
    
    If Picture2.Visible Then
        txtSQL.Height = Abs(Me.Height - Picture2.Height - Toolbar1.Height - StatusBar1.Height - intCaptionBarHeight)
    Else
        txtSQL.Height = Abs(Me.Height - Toolbar1.Height - StatusBar1.Height - intCaptionBarHeight)
    End If
    
    tvDB.Height = Abs(txtSQL.Height - 75)
    grdResults.Width = Me.Width - 175

End Function

Private Sub hSplitter_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    i = 1

End Sub

Private Sub hSplitter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next

    If i = 1 Then
        Picture2.Height = Picture2.Height - Y
        txtSQL.Height = Abs(Me.Height - Picture2.Height - Toolbar1.Height - StatusBar1.Height - intCaptionBarHeight)
        grdResults.Height = Picture2.Height
        tvDB.Height = txtSQL.Height
    End If
    
End Sub

Private Sub hSplitter_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    i = 0

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

    On Error Resume Next

    Select Case Button.Key
        Case "New"
            frmParent.OpenNewChild Me.mstrConnection
        Case "HideResults"
            If Button.value = tbrUnpressed Then
                HideResults True
                frmParent.mnuViewResults.Checked = False
            Else
                HideResults False
                frmParent.mnuViewResults.Checked = True
            End If
        Case "HideDatabase"
            If Button.value = tbrUnpressed Then
                HideDatabase True
                frmParent.mnuViewSchema.Checked = False
            Else
                HideDatabase False
                frmParent.mnuViewSchema.Checked = True
            End If
        Case "Find"
            FindText
        Case "Color"
            AdjustFontColor
        Case "Font"
            ChangeFont
        Case "Execute"
            DAL.BindGridWithData Me, conn
        Case "Preview"
            PreviewSQL
        Case "Print"
            PrintSQL
        Case "Cut"
            CutSQLText
        Case "Copy"
            CopySQLText
        Case "Paste"
            PasteSQLText
        Case "Indent"
            IncreaseIndent
        Case "Outdent"
            DecreaseIndent
        Case "Comment"
            CommentText
        Case "UnComment"
            UnCommentText
        Case "Eraser"
            txtSQL.Text = ""
    End Select

End Sub

Public Sub HideResults(value As Boolean)
'
'Hide/Show the results pane
'
    
    If value = True Then
        Picture2.Visible = False
        sngOldVLocation = txtSQL.Height
        txtSQL.Height = Abs(Me.Height - Toolbar1.Height - StatusBar1.Height - intCaptionBarHeight)
        hSplitter.Enabled = False
    Else
        Picture2.Visible = True
        txtSQL.Height = sngOldVLocation
        hSplitter.Enabled = True
    End If
    
    tvDB.Height = txtSQL.Height
    
End Sub

Public Sub HideDatabase(value As Boolean)
'
'Hide/Show the database pane
'

    If value = True Then
        Picture1.Visible = False
        sngOldHLocation = txtSQL.Left
        txtSQL.Left = 0
        txtSQL.Width = Abs(Me.Width - 100)
    Else
        Picture1.Visible = True
        vSplitter.Enabled = True
        txtSQL.Left = sngOldHLocation
        txtSQL.Width = Abs(Me.Width - Picture1.Width - 100)
    End If
    
End Sub

Public Sub PrintSQL()
'
'Prints the current SQL Query text
'
    Screen.MousePointer = vbHourglass
        SendToPrinter txtSQL
    Screen.MousePointer = vbNormal

End Sub

Public Sub PreviewSQL()
'
'Preview the current SQL Query text
'

    Screen.MousePointer = vbHourglass
        RichTextToPictureBox txtSQL, frmPrintPreview.picTarget, picSQLText, gMarginLeft, gMarginRight, gMarginTop, gMarginBottom
        frmPrintPreview.Show
        frmPrintPreview.Visible = True
   Screen.MousePointer = vbNormal

End Sub

Public Sub CutSQLText()
'
'Copies and deletes selected text from the query text into the system clipboard
'

    'First copy to clipboard
    Clipboard.SetText txtSQL.SelText
    
    'Delete selected text from query
    txtSQL.SelText = ""

End Sub

Public Sub CopySQLText()
'
'Copies selected text from the query text into the system clipboard
'

    Clipboard.SetText txtSQL.SelText

End Sub

Public Sub PasteSQLText()
'
'Pasts the contents of the system clipboard into the query text
'
Dim intSelStart As Integer

    intSelStart = txtSQL.SelStart
    txtSQL.Text = txtSQL.Text & Clipboard.GetText
    txtSQL.SelStart = intSelStart + Len(Clipboard.GetText)
    
End Sub

Private Sub FindText()

    frmFind.Show vbModal

End Sub

Private Sub AdjustFontColor()
'
'Changes the font color for the selected text
'
    On Error Resume Next

    With CommonDialog1
        .ShowColor
        txtSQL.SelColor = .Color
    End With

End Sub

Private Sub ChangeFont()
'
'Changes the font type for the selected text
'

    On Error Resume Next
    
    With CommonDialog1
        .flags = cdlCFBoth
        .ShowFont
        txtSQL.SelFontName = .FontName
        txtSQL.SelBold = .FontBold
        txtSQL.SelItalic = .FontItalic
        txtSQL.SelFontSize = .FontSize
        txtSQL.SelUnderline = .FontUnderline
    End With
    
End Sub

Private Sub tvDB_DblClick()
Dim strSQL          As String
Dim strTableName    As String
Dim strFieldName    As String
Dim intCount        As Integer
Dim intSelStart     As Integer
Dim SelNode         As Node

    On Error Resume Next
    
    Screen.MousePointer = vbHourglass
    Set SelNode = tvDB.SelectedItem

    Select Case Left(SelNode.Key, 6)
        Case "TABLE_"
            strTableName = ParseTableName(SelNode.Text)
            strSQL = "Select "
            
            'Enumerate columns
            For intCount = 0 To db.Tables(strTableName).Columns.count - 1
                strFieldName = db.Tables(strTableName).Columns(intCount).Name
                strSQL = strSQL & """" & strFieldName & """, "
            Next
            
            'Remove extra comma
            Mid(strSQL, InStrRev(strSQL, ","), 1) = " "
            
            strSQL = Trim(strSQL) & " From " & strTableName

        Case "FIELD_"
            strTableName = ParseTableName(SelNode.Parent.Text)
            strSQL = "Select """ & db.Tables(strTableName).Columns(SelNode.Text).Name & """ From " & strTableName
        Case Else
            Screen.MousePointer = vbNormal
            Exit Sub
    End Select
    
    'Update text box
    If txtSQL.Text <> "" Then
        txtSQL.Text = txtSQL.Text & vbCrLf & vbCrLf
    End If
    
    intSelStart = Len(txtSQL.Text)
    txtSQL.Text = txtSQL.Text & strSQL
    
    'Select SQL
    txtSQL.SelStart = intSelStart
    txtSQL.SelLength = Len(strSQL)
    
    'Execute SQL
    DAL.BindGridWithData Me, conn
    
End Sub

Public Function CommentText()
Dim intSTART    As Integer
Dim intEND      As Integer
Dim intLENGTH   As Integer
Dim txtNEW      As String
Dim SPL_TXT
Dim txtORIGINAL As String

On Error Resume Next

    Select Case gstrCommentStyle
        Case "--", "'"
            intSTART = txtSQL.SelStart     '// REMEMBER THE START POS OF CARRET
            txtORIGINAL = txtSQL.SelText
            txtNEW = ""
            SPL_TXT = Split(txtSQL.SelText, vbCrLf)
                For i = 0 To UBound(SPL_TXT)    '// ADD SPACE BEFORE EACH LINE
                    If Not i = UBound(SPL_TXT) Then
                        txtNEW = txtNEW & gstrCommentStyle & SPL_TXT(i) & vbCrLf
                    Else
                        txtNEW = txtNEW & gstrCommentStyle & SPL_TXT(i) '// DON'T ADD LINE BREAK ON LAST LINE
                    End If
                Next i
            txtSQL.SelText = txtNEW
            txtSQL.SelStart = intSTART
            txtSQL.SelLength = Len(txtNEW)
        Case "/**/"
            intSTART = InStrRev(txtSQL.Text, vbCrLf, IIf(txtSQL.SelStart = 0, 1, txtSQL.SelStart))
            intEND = InStr(txtSQL.SelStart + txtSQL.SelLength, txtSQL.Text, vbCrLf) + 1

            If intSTART <> 0 Then
                intSTART = intSTART + 1
            End If

            txtSQL.SelStart = intSTART
            txtSQL.SelLength = 0
            txtSQL.SelText = "/*"

            If intEND = 1 Then
                intEND = Len(txtSQL.Text)
            End If

            txtSQL.SelStart = intEND
            txtSQL.SelLength = 0
            txtSQL.SelText = "*/"
            
            txtSQL.SelStart = intSTART
            txtSQL.SelLength = (intEND + 2) - intSTART
    End Select

End Function

Public Function UnCommentText()
Dim lngSelEnd   As Integer
Dim lngSelSTart As Integer

On Error Resume Next

    Select Case gstrCommentStyle
        Case "--", "'"
            lngSelEnd = txtSQL.SelStart + txtSQL.SelLength
            
            Do Until lngSelSTart > lngSelEnd
                lngSelSTart = InStrRev(txtSQL.Text, vbCrLf, IIf(txtSQL.SelStart = 0, 1, txtSQL.SelStart))

                If lngSelSTart <> 0 Then
                    lngSelSTart = lngSelSTart + 1
                End If
            
                txtSQL.SelStart = lngSelSTart
                txtSQL.SelLength = Len(gstrCommentStyle)
                txtSQL.SelText = Replace(txtSQL.SelText, gstrCommentStyle, "")
                
                'move to new line
                lngSelSTart = InStr(lngSelSTart, txtSQL.Text, vbCrLf) + 2
                txtSQL.SelStart = lngSelSTart
            Loop
        Case "/**/"
            lngSelSTart = InStrRev(txtSQL.Text, vbCrLf, IIf(txtSQL.SelStart = 0, 1, txtSQL.SelStart))
            lngSelEnd = InStr(txtSQL.SelStart + txtSQL.SelLength - 2, txtSQL.Text, "*/") - 3

            If lngSelSTart <> 0 Then
                lngSelSTart = lngSelSTart + 1
            End If
            
            txtSQL.SelStart = lngSelSTart
            txtSQL.SelLength = 2
            txtSQL.SelText = ""
            
            If lngSelEnd = 1 Then
                lngSelEnd = Len(txtSQL.Text)
            End If

            txtSQL.SelStart = lngSelEnd
            txtSQL.SelLength = 2
            txtSQL.SelText = ""
    End Select

End Function

Public Function IncreaseIndent()
Dim intSTART    As Integer
Dim intLENGTH   As Integer
Dim txtNEW      As String
Dim SPL_TXT
Dim txtORIGINAL As String

    intSTART = txtSQL.SelStart     '// REMEMBER THE START POS OF CARRET
    txtORIGINAL = txtSQL.SelText
    txtNEW = ""
    SPL_TXT = Split(txtSQL.SelText, vbCrLf)
        For i = 0 To UBound(SPL_TXT)    '// ADD SPACE BEFORE EACH LINE
            If Not i = UBound(SPL_TXT) Then
                txtNEW = txtNEW & Space(5) & SPL_TXT(i) & vbCrLf
            Else
                txtNEW = txtNEW & Space(5) & SPL_TXT(i) '// DON'T ADD LINE BREAK ON LAST LINE
            End If
        Next i
    txtSQL.SelText = txtNEW
    txtSQL.SelStart = intSTART
    txtSQL.SelLength = Len(txtNEW)
    
End Function

Public Function DecreaseIndent()
Dim strORIGINAL     As String
Dim i               As Integer
Dim strNEW          As String
Dim intSTART        As Integer
Dim arrLine
Dim LN

    intSTART = txtSQL.SelStart
    strORIGINAL = txtSQL.SelText
    arrLine = Split(strORIGINAL, vbCrLf)
    '// TAKE EACH LINE AND CHECK IF WE HAVE ROOM TO MOVE
    For i = 0 To UBound(arrLine)
        If Mid(arrLine(i), 1, 5) = Space(5) Then
            LN = Mid(arrLine(i), 5, Len(arrLine(i)))
        Else
            LN = TrimLeftSpaces(arrLine(i), 5)
        End If
        '// CREATE NEW BLOCK
        If i = UBound(arrLine) Then
          strNEW = strNEW & LN
        Else
          strNEW = strNEW & LN & vbCrLf
        End If
    Next i
    '// SET NEW BLOCK OF TEXT
    txtSQL.SelText = strNEW
    txtSQL.SelStart = intSTART
    txtSQL.SelLength = Len(strNEW)

End Function

Private Function TrimLeftSpaces(strString, intLENGTH As Long) As String
Dim intCount    As Integer
Dim strNEW      As String

    strNEW = strString

    For intCount = 1 To intLENGTH
        If Mid(strString, intCount, 1) = " " Then
            strNEW = Right(strNEW, Len(strNEW) - 1)
        Else
            Exit For
        End If
    Next
    
    TrimLeftSpaces = strNEW

End Function

Public Function MakeUppercase()
Dim strSelStart As String
Dim strSelLength As String

    strSelStart = txtSQL.SelStart
    strSelLength = txtSQL.SelLength
    txtSQL.SelText = UCase(txtSQL.SelText)
    txtSQL.SelStart = strSelStart
    txtSQL.SelLength = strSelLength

End Function

Public Function MakeLowercase()
Dim strSelStart As String
Dim strSelLength As String

    strSelStart = txtSQL.SelStart
    strSelLength = txtSQL.SelLength
    txtSQL.SelText = LCase(txtSQL.SelText)
    txtSQL.SelStart = strSelStart
    txtSQL.SelLength = strSelLength
    
End Function

Private Function ParseTableName(strNodeText As String, Optional strReturnSchema As String = "") As String

    On Error Resume Next

    If InStr(1, strNodeText, "(") > 0 Then
        ParseTableName = Left(strNodeText, InStr(1, strNodeText, " (") - 1)
        strReturnSchema = Mid(strNodeText, _
                              InStr(1, strNodeText, "(") + 1, _
                              (InStr(1, strNodeText, ")") - 1) - InStr(1, strNodeText, "("))
    Else
        ParseTableName = strNodeText
    End If

End Function

Private Sub tvDB_Expand(ByVal Node As MSComctlLib.Node)
Dim strTableName    As String
Dim strFieldName    As String
Dim intCount        As Integer
Dim strDelNodeKey   As String
Dim intIndexes      As Integer
Dim strImageKey     As String
Dim intColumns      As Integer
Dim strSchema       As String

    On Error Resume Next
    
    If Left(Node.Key, 8) = "COLUMNS_" And Node.Child.Text = "wait" And Node.Children = 1 Then
        StatusBar1.Panels(1).Text = "Please wait...": DoEvents
        Screen.MousePointer = vbHourglass
        
        strTableName = ParseTableName(Node.Parent.Text)
        
        For intCount = 0 To db.Tables(strTableName).Columns.count - 1
            strFieldName = db.Tables(strTableName).Columns(intCount).Name
            
            strImageKey = "TABLE COLUMN"

            
            For intIndexes = 0 To db.Tables(strTableName).Indexes.count - 1
                For intColumns = 0 To db.Tables(strTableName).Indexes(intIndexes).Columns.count - 1
                    If db.Tables(strTableName).Indexes(intIndexes).Columns(intColumns).Name = strFieldName And _
                       db.Tables(strTableName).Columns(intCount).Attributes = adColFixed Then
                        If Err.Number = 0 Then
                            strImageKey = "TABLE KEY"
                        End If
                        
                        Exit For
                    End If
                Next
            Next

            
            tvDB.Nodes.Add Node.Key, tvwChild, strTableName & "_FIELD_" & Node.Text & strFieldName, strFieldName, strImageKey
        Next
        
        strDelNodeKey = Node.Child.Key
        tvDB.Nodes(Node.Key).Sorted = False
        
        Screen.MousePointer = vbNormal
        StatusBar1.Panels(1).Text = "Ready"
    ElseIf Left(Node.Key, 14) = "RELATIONSHIPS_" And Node.Child.Text = "wait" And Node.Children = 1 Then
        
        strDelNodeKey = Node.Child.Key
        strTableName = ParseTableName(Node.Parent.Text, strSchema)
        AddRelationshipsToTV strTableName, strSchema
        
    ElseIf Left(Node.Key, 8) = "INDEXES_" And Node.Child.Text = "wait" And Node.Children = 1 Then
        
        strDelNodeKey = Node.Child.Key
        strTableName = ParseTableName(Node.Parent.Text, strSchema)
        AddIndexesToTV strTableName, strSchema
        
    ElseIf Left(Node.Key, 12) = "CONSTRAINTS_" And Node.Child.Text = "wait" And Node.Children = 1 Then
    
        strDelNodeKey = Node.Child.Key
        strTableName = ParseTableName(Node.Parent.Text, strSchema)
        AddConstraintsToTV strTableName, strSchema
    
    Else
        If Node.Child.Text = "wait" And Node.Children = 1 Then
            strDelNodeKey = Node.Child.Key
            AddTablesToTV Node.Key
        End If
    End If
    
    'Remove wait
    If strDelNodeKey <> "" Then
        tvDB.Nodes.Remove strDelNodeKey
    End If


End Sub

Private Sub txtSQL_GotFocus()
'Dim Control As Control
'
'   ' Ignore errors for controls without the TabStop property.
'   On Error Resume Next
'   ' Switch off the change of focus when pressing TAB.
'   For Each Control In Controls
'      Control.TabStop = False
'   Next Control
End Sub

Private Sub txtSQL_KeyDown(KeyCode As Integer, Shift As Integer)
  
  If txtSQL.SelLength > 0 Then            '// IF NOTHING IS SELECTED THEN DO NOTHING
    If KeyCode = 9 And Shift = 1 Then  '// DON'T TUCH OTHER KEYS
       Call DecreaseIndent                '// UNINDENT THE BLOCK OF TEXT
       KeyCode = 27                    '// CANCEL THE TAB KEY
    End If
  End If
  
End Sub

Private Sub txtSQL_KeyPress(KeyAscii As Integer)

    If KeyAscii = 9 Then             '// DON'T TUCH OTHER KEYS
       If Not txtSQL.SelLength = 0 Then '// IF NOTHING IS SELECTED THEN ALLOW FOR NORMAL TAB KEY
            Call IncreaseIndent      '// MOVE TEXT IF TAB IS PRESSED
            KeyAscii = 27            '// CANCEL THE TAB KEY
       End If
    End If

End Sub

Private Sub txtSQL_SelChange()

    If txtSQL.SelLength > 0 Then
        Toolbar1.Buttons("Copy").Enabled = True
        Toolbar1.Buttons("Cut").Enabled = True
        Toolbar1.Buttons("Indent").Enabled = True
        Toolbar1.Buttons("Outdent").Enabled = True
        Toolbar1.Buttons("Comment").Enabled = True
        Toolbar1.Buttons("UnComment").Enabled = True
        frmParent.mnuEditAdvancedIncreaseIndent.Enabled = True
        frmParent.mnuEditAdvancedDecreaseIndent.Enabled = True
        frmParent.mnuEditAdvancedSelLowerCase.Enabled = True
        frmParent.mnuEditAdvancedSelUpperCase.Enabled = True
        frmParent.mnuEditAdvancedComment.Enabled = True
        frmParent.mnuEditAdvancedUncomment.Enabled = True
    Else
        Toolbar1.Buttons("Copy").Enabled = False
        Toolbar1.Buttons("Cut").Enabled = False
        Toolbar1.Buttons("Indent").Enabled = False
        Toolbar1.Buttons("Outdent").Enabled = False
        Toolbar1.Buttons("Comment").Enabled = False
        Toolbar1.Buttons("UnComment").Enabled = False
        frmParent.mnuEditAdvancedIncreaseIndent.Enabled = False
        frmParent.mnuEditAdvancedDecreaseIndent.Enabled = False
        frmParent.mnuEditAdvancedSelLowerCase.Enabled = False
        frmParent.mnuEditAdvancedSelUpperCase.Enabled = False
        frmParent.mnuEditAdvancedComment.Enabled = False
        frmParent.mnuEditAdvancedUncomment.Enabled = False
    End If

End Sub

Private Sub vSplitter_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
     
     i = 1
     
End Sub

Private Sub vSplitter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    On Error Resume Next
    
    If i = 1 Then
        Picture1.Width = Abs(Picture1.Width + X)
        txtSQL.Width = Me.Width - Picture1.Width - 100
        txtSQL.Left = Picture1.Left + Picture1.Width + vSplitter.Width
        tvDB.Width = Picture1.Width
    End If

End Sub

Private Sub vSplitter_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    i = 0

End Sub

Private Function AddTableGroupsToTV()
Dim strDataSource As String

    On Error Resume Next
    
    strDataSource = conn.Properties("Data Source") & " (" & conn.Properties("DBMS Name") & ")"
    
    tvDB.Nodes.Add , tvwFirst, "DataSource", strDataSource, "Database"
    tvDB.Nodes.Add "DataSource", tvwChild, "TABLE", "Tables", "tables"
        tvDB.Nodes.Add "TABLE", tvwChild, "TABLEwait", "wait"
    tvDB.Nodes.Add "DataSource", tvwChild, "VIEW", "Views", "views"
        tvDB.Nodes.Add "VIEW", tvwChild, "VIEWwait", "wait"
    tvDB.Nodes.Add "DataSource", tvwChild, "SYSTEM TABLE", "System Tables", "system_tables"
        tvDB.Nodes.Add "SYSTEM TABLE", tvwChild, "SYSTEM TABLEwait", "wait"
    tvDB.Nodes.Add "DataSource", tvwChild, "SYNONYM", "Aliases/Synonyms", "synonyms"
        tvDB.Nodes.Add "SYNONYM", tvwChild, "SYNONYMwait", "wait"
    tvDB.Nodes.Add "DataSource", tvwChild, "GLOBAL TEMPORARY", "Temporary Tables", "temporary_tables"
        tvDB.Nodes.Add "GLOBAL TEMPORARY", tvwChild, "GLOBAL TEMPORARYwait", "wait"
    tvDB.Nodes.Add "DataSource", tvwChild, "PROCEDURE", "Procedures/Functions", "functions"
        tvDB.Nodes.Add "PROCEDURE", tvwChild, "PROCEDUREwait", "wait"
    
    tvDB.Nodes("DataSource").Expanded = True

End Function

Private Function AddCatalogs()
Dim rs          As New ADODB.Recordset
Dim strDBMS     As String
Dim intIndexCount As Integer

    On Error GoTo AddCatalogs_Error

    strDBMS = conn.Properties("Current Catalog")
    intIndexCount = 1

    '
    'ADO
    '
    Set rs = conn.OpenSchema(adSchemaCatalogs)
    
    Do Until rs.EOF
        cboCatalogs.ComboItems.Add intIndexCount, , rs("CATALOG_NAME"), "Database"
        
        If rs("CATALOG_NAME") = strDBMS Then
            cboCatalogs.ComboItems(intIndexCount).Selected = True
        End If
        
        intIndexCount = intIndexCount + 1
        rs.MoveNext
    Loop
    
    'Close recordsets
    rs.Close
    Set rs = Nothing

AddCatalogs_EXIT:
    Exit Function

AddCatalogs_Error:
    If Err.Number = 3251 Then
        cboCatalogs.ComboItems.Add 1, , "<Not Supported>", "Not Supported"
        cboCatalogs.ComboItems.ITem(1).Selected = True
        cboCatalogs.Enabled = False
    End If
    
    Resume AddCatalogs_EXIT
    
End Function

Private Function AddTablesToTV(strSection As String)
Dim intCount    As Integer
Dim rs          As New ADODB.Recordset
Dim strRelative As String
Dim strTableKey As String
Dim strTableName As String
Dim dbProc      As ADOX.Procedure

    On Error Resume Next
    
    Screen.MousePointer = vbHourglass
    Me.StatusBar1.Panels(1).Text = "Please wait...": DoEvents
    tvDB.Visible = False
   
    If strSection <> "PROCEDURE" Then
        '
        'ADO
        '
        Set rs = conn.OpenSchema(adSchemaTables)

        Do Until rs.EOF
            strRelative = rs("TABLE_TYPE")
            
            If strRelative = strSection Then
                If Trim(rs("TABLE_SCHEMA") & "") <> "" Then
                    strTableKey = "TABLE_" & rs("TABLE_NAME") & " (" & rs("TABLE_SCHEMA") & ")"
                    strTableName = rs("TABLE_NAME") & " (" & rs("TABLE_SCHEMA") & ")"
                Else
                    strTableKey = "TABLE_" & rs("TABLE_NAME")
                    strTableName = rs("TABLE_NAME")
                End If
            
                'Table
                tvDB.Nodes.Add strSection, tvwChild, strTableKey, strTableName, strRelative
                
                'Columns
                tvDB.Nodes.Add strTableKey, tvwChild, "COLUMNS_" & strTableKey, "Columns", "Columns"
                tvDB.Nodes.Add "COLUMNS_" & strTableKey, tvwChild, strTableKey & "columnwait", "wait"
                
                'Constraints
                tvDB.Nodes.Add strTableKey, tvwChild, "CONSTRAINTS_" & strTableKey, "Constraints", "Constraints"
                tvDB.Nodes.Add "CONSTRAINTS_" & strTableKey, tvwChild, strTableKey & "constraintswait", "wait"
        
                'Indexes
                tvDB.Nodes.Add strTableKey, tvwChild, "INDEXES_" & strTableKey, "Indexes", "Indexes"
                tvDB.Nodes.Add "INDEXES_" & strTableKey, tvwChild, strTableKey & "indexeswait", "wait"
                
                'Relationships
                tvDB.Nodes.Add strTableKey, tvwChild, "RELATIONSHIPS_" & strTableKey, "Relationships", "Relationships"
                tvDB.Nodes.Add "RELATIONSHIPS_" & strTableKey, tvwChild, strTableKey & "relationshipwait", "wait"
                
            End If
          
            rs.MoveNext
        Loop
        
        'Close recordsets
        rs.Close
        Set rs = Nothing
    Else
        '
        'ADOX
        '
        For Each dbProc In db.Procedures
            tvDB.Nodes.Add "PROCEDURE", tvwChild, , dbProc.Name, "PROCEDURE"
        Next
    End If
    
    'Sort table names
    tvDB.Nodes(strSection).Sorted = True
    
    tvDB.Visible = True
    Me.StatusBar1.Panels(1).Text = "Ready"
    Screen.MousePointer = vbNormal

End Function

Private Function AddRelationshipsToTV(strTable As String, strSchema As String)
Dim rs          As New ADODB.Recordset
Dim strTableKey As String


    On Error Resume Next
    
    Screen.MousePointer = vbHourglass
    Me.StatusBar1.Panels(1).Text = "Please wait...": DoEvents
    tvDB.Visible = False
    strTableKey = "TABLE_" & strTable & " (" & strSchema & ")"

    '
    'ADO
    '
    Set rs = conn.OpenSchema(adSchemaForeignKeys)
    
    'Primary Keys
    rs.Filter = "PK_TABLE_NAME = '" & strTable & "' AND FK_TABLE_SCHEMA = '" & strSchema & "'"
    
    Do Until rs.EOF
        tvDB.Nodes.Add "RELATIONSHIPS_" & strTableKey, tvwChild, strTable & "_RELATIONSHIPS_" & rs("FK_NAME"), rs("FK_NAME"), "TABLE KEY"
        tvDB.Nodes.Add strTable & "_RELATIONSHIPS_" & rs("FK_NAME"), _
                        tvwChild, _
                        strTable & rs("FK_NAME") & rs("FK_TABLE_NAME"), _
                        rs("FK_TABLE_NAME"), _
                        "TABLE"
        tvDB.Nodes.Add strTable & rs("FK_NAME") & rs("FK_TABLE_NAME"), _
                       tvwChild, _
                       strTable & rs("FK_NAME") & rs("FK_TABLE_NAME") & rs("FK_COLUMN_NAME"), _
                       rs("FK_COLUMN_NAME"), _
                       "TABLE COLUMN"
            
        rs.MoveNext
    Loop
    
    rs.Filter = ""
    
    'Foriegn Keys
    rs.Filter = "FK_TABLE_NAME = '" & strTable & "' AND FK_TABLE_SCHEMA = '" & strSchema & "'"
    
    Do Until rs.EOF
        tvDB.Nodes.Add "RELATIONSHIPS_" & strTableKey, tvwChild, strTable & "_RELATIONSHIPS_" & rs("FK_NAME"), rs("FK_NAME"), "Relationships"
        tvDB.Nodes.Add strTable & "_RELATIONSHIPS_" & rs("FK_NAME"), _
                        tvwChild, _
                        strTable & "_RELATIONSHIPS_" & rs("FK_NAME") & rs("PK_TABLE_NAME"), _
                        rs("PK_TABLE_NAME"), _
                        "TABLE"
        tvDB.Nodes.Add strTable & "_RELATIONSHIPS_" & rs("FK_NAME") & rs("PK_TABLE_NAME"), _
                        tvwChild, _
                        strTable & "_RELATIONSHIPS_" & rs("FK_NAME") & rs("PK_TABLE_NAME") & rs("PK_COLUMN_NAME"), _
                        rs("PK_COLUMN_NAME"), _
                        "TABLE COLUMN"
        rs.MoveNext
    Loop
    
    'Close recordsets
    rs.Close
    Set rs = Nothing
    
    tvDB.Visible = True
    Me.StatusBar1.Panels(1).Text = "Ready"
    Screen.MousePointer = vbNormal
    
End Function

Public Function AddIndexesToTV(strTable As String, strSchema As String)
Dim rs          As New ADODB.Recordset
Dim rsCol       As New ADODB.Recordset
Dim strTableKey As String
Dim strIndexKey As String


    On Error Resume Next
    
    Screen.MousePointer = vbHourglass
    Me.StatusBar1.Panels(1).Text = "Please wait...": DoEvents
    tvDB.Visible = False
    strTableKey = "TABLE_" & strTable & " (" & strSchema & ")"

    '
    'ADO
    '
    Set rs = conn.OpenSchema(adSchemaIndexes)
    Set rsCol = conn.OpenSchema(adSchemaIndexes)
    
    rs.Filter = "TABLE_NAME = '" & strTable & "' AND INDEX_SCHEMA = '" & strSchema & "'"
        
    Do Until rs.EOF
    
        strIndexKey = strTable & "_INDEXES_" & rs("INDEX_NAME")
    
        tvDB.Nodes.Add "INDEXES_" & strTableKey, tvwChild, strIndexKey, rs("INDEX_NAME"), "Indexes"
    
        rsCol.Filter = "TABLE_NAME = '" & strTable & "' AND INDEX_SCHEMA = '" & strSchema & "' AND INDEX_NAME = '" & rs("INDEX_NAME") & "'"
    
        Do Until rsCol.EOF
            tvDB.Nodes.Add strIndexKey, tvwChild, strIndexKey & "_COLUMN_" & rsCol("COLUMN_NAME"), rsCol("COLUMN_NAME"), "TABLE COLUMN"
        
            rsCol.MoveNext
            rs.MoveNext
        Loop
        
        rsCol.Filter = ""
    
        'rs.MoveNext
    Loop

    'Close recordsets
    rs.Close
    rsCol.Close
    Set rs = Nothing
    Set rsCol = Nothing
    
    tvDB.Visible = True
    Me.StatusBar1.Panels(1).Text = "Ready"
    Screen.MousePointer = vbNormal

End Function

Public Function AddConstraintsToTV(strTable As String, strSchema As String)
Dim strTableKey As String
Dim strIndexKey As String
Dim rs          As New ADODB.Recordset
Dim rsCol       As New ADODB.Recordset

    On Error Resume Next
    
    Screen.MousePointer = vbHourglass
    Me.StatusBar1.Panels(1).Text = "Please wait...": DoEvents
    tvDB.Visible = False
    strTableKey = "TABLE_" & strTable & " (" & strSchema & ")"

    '
    'ADO
    '
    Set rs = conn.OpenSchema(adSchemaCatalogs)
    
    Set rsCol = conn.OpenSchema(adSchemaColumns)
    
    Set Me.grdResults.DataSource = rsCol.DataSource
    
    'Close recordsets
    'rs.Close
    'Set rs = Nothing
    
    tvDB.Visible = True
    Me.StatusBar1.Panels(1).Text = "Ready"
    Screen.MousePointer = vbNormal


End Function

Public Function ViewColumnProperties()
Dim rs              As New ADODB.Recordset
Dim strTable        As String
Dim strSchema       As String
Dim strColumnName   As String
Dim strSQL          As String

    Set rs = conn.OpenSchema(adSchemaColumns)
    
    strTable = ParseTableName(tvDB.SelectedItem.Parent.Parent.Text, strSchema)
    strColumnName = tvDB.SelectedItem.Text
    strSQL = "TABLE_NAME = '" & strTable & "'"
    
    If strSchema <> "" Then
        strSQL = strSQL & " AND TABLE_SCHEMA = '" & strSchema & "'"
    End If
    
    strSQL = strSQL & " AND COLUMN_NAME = '" & strColumnName & "'"
    
    rs.Filter = strSQL

    frmProperties.DisplayProperties rs.Fields

End Function





























