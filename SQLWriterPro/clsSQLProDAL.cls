VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSQLProDAL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents rs   As ADODB.Recordset
Attribute rs.VB_VarHelpID = -1
Private WithEvents mCnn As ADODB.Connection
Attribute mCnn.VB_VarHelpID = -1
Private mFrm            As Form
Private mbActionQuery   As Boolean
Private mbCommit        As Boolean

Public Function ConnectionWizard() As String
Dim connObj As MSDASC.DataLinks
Dim strConnection As String
    
    On Error Resume Next
    
    'Establish OLE DB connection
    Set connObj = New MSDASC.DataLinks
    strConnection = connObj.PromptNew
    Set connObj = Nothing

    'Return a fully qualified connection string
    ConnectionWizard = strConnection
        
End Function

Public Function CreateRecordset(strSQL As String, strConnection As String, conn As ADODB.Connection) As Boolean
    
    On Error GoTo CreateRecordset_Error
    
    If rs.State = adStateOpen Then
        rs.Close
    End If
    
    'Open the recordset and fetch records asynchronously
    rs.Open strSQL, _
             conn, _
             ADODB.adOpenStatic, _
             ADODB.adLockBatchOptimistic, _
             adAsyncFetch
             
    'Return first batch of results immediately
    Set mFrm.grdResults.DataSource = rs
             
    'Disconnect the recordset.
    Set rs.ActiveConnection = Nothing
        
CreateRecordset_Exit:
    Exit Function
    
CreateRecordset_Error:
    MsgBox Err.Description, vbCritical, "SQL Writer"
    Resume CreateRecordset_Exit

End Function

Public Function BindGridWithData(frm As frmChild, conn As ADODB.Connection) As Boolean
Dim strSQL      As String

    On Error GoTo BindGridWithData_Error
    
    If frm.Toolbar1.Buttons("Execute").ToolTipText = "Stop Fetching" Then
        rs.Cancel: DoEvents
        
        frm.Toolbar1.Buttons("Execute").ToolTipText = "Execute SQL (F5)"
        frm.Toolbar1.Buttons("Execute").Image = "Execute"
    Else
        frm.StatusBar1.Panels(1).Text = "Executing..."
        frm.StatusBar1.Panels(3).Text = "Fetching..."
        frm.Toolbar1.Buttons("Execute").ToolTipText = "Stop Fetching"
        frm.Toolbar1.Buttons("Execute").Image = "Stop"
                
        If frm.txtSQL <> "" Then
            Set mCnn = conn
            Set mFrm = frm
            
            If frm.txtSQL.SelLength = 0 Then
                strSQL = frm.txtSQL.Text
            Else
                strSQL = frm.txtSQL.SelText
            End If
            
            Screen.MousePointer = vbHourglass
            
            If LCase(Left(strSQL, 6)) = "delete" Then
                mbActionQuery = True
                mCnn.BeginTrans
                mCnn.Execute strSQL, , adCmdText
                
                If mbCommit Then
                    mCnn.CommitTrans
                Else
                    mCnn.RollbackTrans
                End If
            Else
                CreateRecordset strSQL, frm.mstrConnection, mCnn
            End If
        End If
    End If
    
    BindGridWithData = True
    
BindGridWithData_Exit:
    frm.StatusBar1.Panels(1).Text = "Ready"
    Screen.MousePointer = vbNormal
    Exit Function
    
BindGridWithData_Error:
    'Report error
    BindGridWithData = False
    MsgBox Err.Description, vbCritical, "SQL Writer"
    
    mbActionQuery = False
    mCnn.RollbackTrans
    
    'Exit immediately
    Resume BindGridWithData_Exit

End Function

Private Sub Class_Initialize()

    Set rs = New ADODB.Recordset
    
    rs.CursorLocation = adUseClient
    rs.CacheSize = 1
    rs.Properties("Initial Fetch Size") = 1
    rs.Properties("Background Fetch Size") = 1
    rs.Properties("Asynchronous Rowset Processing") = 2
    
End Sub

Private Sub mCnn_ExecuteComplete(ByVal RecordsAffected As Long, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pCommand As ADODB.Command, ByVal pRecordset As ADODB.Recordset, ByVal pConnection As ADODB.Connection)
Dim answer As VbMsgBoxResult

    If adStatus = adStatusOK And mbActionQuery Then
        answer = MsgBox("You are about to delete " & RecordsAffected & " row(s) from the specified table." & _
                        vbCrLf & vbCrLf & _
                        "Once you click Yes, you can't Undo this command to reverse the changes." & vbCrLf & _
                        "Are you sure you want to delete the selected records?", vbExclamation + vbYesNo + vbDefaultButton1, "SQL Writer Pro")
        If answer = vbYes Then
            mbCommit = True
        Else
            mbCommit = False
        End If
        
        mbActionQuery = False
    End If

End Sub


Private Sub rs_FetchComplete(ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
'
'If OLEDB provider supports Asynchronous Fetching then this event will fire
'
    If adStatus = adStatusOK Then
        'Set grid with complete datasource
        Set mFrm.grdResults.DataSource = pRecordset
        mFrm.grdResults.Refresh
        mFrm.StatusBar1.Panels(3).Text = pRecordset.RecordCount & " Records"
    ElseIf adStatus = adStatusCancel Then
        mFrm.StatusBar1.Panels(3).Text = "Cancelled"
    Else
        'Return Error
        mFrm.StatusBar1.Panels(3).Text = "Error"
    End If
    
    'Reset icon
    mFrm.Toolbar1.Buttons("Execute").ToolTipText = "Execute SQL (F5)"
    mFrm.Toolbar1.Buttons("Execute").Image = "Execute"
    
End Sub

