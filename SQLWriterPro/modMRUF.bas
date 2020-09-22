Attribute VB_Name = "modMRUF"
Option Explicit

Private colMRUFiles As New Dictionary
Private Num As Integer

Public Sub Add(ConnectionLabel As String, ConnectionString As String)

    If colMRUFiles.count = 0 Then
        'Add, because we know it doesn't already exist
        colMRUFiles.Add ConnectionLabel, ConnectionString
    Else
        'Remove if it exists, then add it
        On Error Resume Next
        colMRUFiles.Remove ConnectionLabel
        colMRUFiles.Add ConnectionLabel, ConnectionString

        'Drop last one and keep only the top five(NUM)
        If colMRUFiles.count > Num Then
            colMRUFiles.Remove colMRUFiles.count
        End If
    End If
    
End Sub

Public Sub Clear()
    
    'Clears all files from the list
    Do While colMRUFiles.count > 0
        colMRUFiles.Remove 1
    Loop
    
End Sub

Public Property Get count() As Long
    
    'Returns the number of files in the list.
    count = colMRUFiles.count
    
End Property

Public Property Get ITem(strKey As String) As String
    
    'Returns the nth item from the list
    On Error GoTo ItemError
    ITem = colMRUFiles.ITem(strKey)
    
    Exit Property
    
ItemError:
    ITem = ""
    
End Property

Public Sub Load(Optional AppName As Variant)
Dim v                   As Variant
Dim i                   As Integer
Dim J                   As Integer
Dim AppN                As String
Dim ConnectionLabel     As String
Dim ConnectionString    As String

    'Ensure we have an application title
    If IsMissing(AppName) Then
        AppN = App.Title
    Else
        AppN = CStr(AppName)
    End If
    
    'Returns a Variant as a two-dimensional array
    v = GetAllSettings(AppN, "colMRUFiles")

    'Add to dictionary collection
    If Not IsEmpty(v) Then
        i = UBound(v, 1)    'Number of connections
        ConnectionLabel = v(i, 0)
        ConnectionString = v(i, 1)
        
        'Add the last one
        colMRUFiles.Add ConnectionLabel, ConnectionString

        'Add all the rest
        For J = i - 1 To LBound(v, 1) Step -1
            colMRUFiles.Add v(J, 0), v(J, 1)
        Next J
    End If
    
End Sub

Public Property Get Number() As Integer

    'Gets the maximum size of the list
    Number = Num
    
End Property


Public Property Let Number(i As Integer)
    
    'Sets the maximum size of the list
    Num = i
    
End Property

Public Sub Remove(ConnectionLabel As String)

    On Error Resume Next
    colMRUFiles.Remove ConnectionLabel
    
End Sub

Public Sub Save(Optional AppName As Variant)
Dim i As Integer
Dim AppN As String
    
    On Error Resume Next

    'Ensure we have an App title
    If IsMissing(AppName) Then
        AppN = App.Title
    Else
        AppN = CStr(AppName)
    End If
    
    'First delete
    DeleteSetting AppN, "colMRUFiles"

    'Then add dictionary object
    For i = 0 To colMRUFiles.count - 1
        SaveSetting AppN, "colMRUFiles", colMRUFiles.Keys(i), colMRUFiles.Items(i)
    Next i
    
End Sub

Public Sub Update(F As Form)
' *** Note: The form must contain a menu
'     control array
' ***named mnuMRUFiles that is at least
'     as big
' ***as Number.
Dim i As Long
    
    On Error GoTo NextStep

    'First hide all menus
    For i = 1 To Num
        F.mnuMRUFiles(i).Visible = False
    Next i

NextStep:
    On Error GoTo MenuEnd
    
    If colMRUFiles.count > 0 Then
        For i = 1 To Num 'colMRUFiles.count
            F.mnuMRUFiles(i).Caption = i & " " & colMRUFiles.Keys(i - 1)
            F.mnuMRUFiles(i).Visible = True
        Next i
        
        F.mnuMRUFiles(Num + 1).Visible = True
    Else
        F.mnuMRUFiles(Num + 1).Visible = False
    End If

MenuEnd:
End Sub

Private Sub Class_Initialize()
    
    'Number of connections to track
    Num = 5
    
End Sub

