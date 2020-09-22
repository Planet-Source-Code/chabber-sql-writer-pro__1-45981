Attribute VB_Name = "modPrintPreview"
Option Explicit

Public Type POINTAPI
        x As Long
        Y As Long
End Type

Private Type Rect
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public gNumPages As Long

Public Type PAGESETUPDLG
        lStructSize As Long
        hwndOwner As Long
        hDevMode As Long
        hDevNames As Long
        flags As Long
        ptPaperSize As POINTAPI
        rtMinMargin As Rect
        rtMargin As Rect
        hInstance As Long
        lCustData As Long
        lpfnPageSetupHook As Long
        lpfnPagePaintHook As Long
        lpPageSetupTemplateName As String
        hPageSetupTemplate As Long
End Type

Public Declare Function PAGESETUPDLG Lib "comdlg32.dll" Alias "PageSetupDlgA" (pPagesetupdlg As PAGESETUPDLG) As Long

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
     (ByVal hwnd As Long, ByVal msg As Long, ByVal wp As Long, Ip As Any) As Long
     
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, _
    ByVal Y As Long, ByVal mDestWidth As Long, ByVal mDestHeight As Long, _
    ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal mSrcWidth As Long, _
    ByVal mSrcHeight As Long, ByVal dwRop As Long) As Long

Private Const WM_PASTE = &H302
Private Const WM_USER As Long = &H400
Private Const EM_FORMATRANGE As Long = WM_USER + 57

Private Type CharRange
    firstChar As Long         ' First character of range (0 for start of doc)
    lastChar As Long          ' Last character of range (-1 for end of doc)
End Type

Private Type FormatRange
    hdc As Long               ' Actual DC to draw on
    hdcTarget As Long         ' Target DC for determining text formatting
    rectRegion As Rect        ' Region of the DC to draw to (in twips)
    rectPage As Rect          ' Page size of the entire DC (in twips)
    mCharRange As CharRange   ' Range of text to draw (see above user type)
End Type

Private mFormatRange As FormatRange
Private rectDrawTo As Rect
Private rectPage As Rect

Public Const cnstFullPreviewWidth = 5851
Public Const cnstFullPreviewHeight = 7831
Public Const PSD_MARGINS = &H2
Public Const PSD_DISABLEORIENTATION = &H100
Public Const PSD_DISABLEPAPER = &H200

Public Function RichTextToPictureBox(ByRef ctlSourceRichTextBox As RichTextBox, ByRef ctlDestPictureBox As PictureBox, ByVal ctlOrigPictureBox As PictureBox, _
                                     Optional gLeftMargin As Long = 1, _
                                     Optional gRightMargin As Long = 1, _
                                     Optional gTopMargin As Long = 1, _
                                     Optional gBottomMargin As Long = 1, _
                                     Optional gPageNum As Integer = 1) As Boolean
Dim retValue As Long
Dim CurrentPage As Integer
Dim dumpaway As Long

    frmPrintPreview.CurrentPage = gPageNum
    Printer.ScaleMode = vbTwips

    'This just gets the original size into a picture box
    With ctlOrigPictureBox
        .Left = ctlSourceRichTextBox.Left
        .Top = ctlSourceRichTextBox.Top
        .Width = cnstFullPreviewWidth * 2
        .Height = cnstFullPreviewHeight * 2
    End With
    DoEvents

    'Ensure number of pages has been set
    CalculateNumberOfPages ctlSourceRichTextBox, ctlOrigPictureBox, gLeftMargin, gRightMargin, gTopMargin, gBottomMargin

    'Set printable area rect.
    rectPage.Left = 0
    rectPage.Top = 0
    rectPage.Right = ctlSourceRichTextBox.Width
    rectPage.Bottom = ctlSourceRichTextBox.Height
    
    'Set rect in which to print (relative to printable area)
    rectDrawTo.Left = gLeftMargin * 1440
    rectDrawTo.Top = gTopMargin * 1440
    rectDrawTo.Right = (ctlOrigPictureBox.Width) - gRightMargin * 1440
    rectDrawTo.Bottom = (ctlOrigPictureBox.Height) - gBottomMargin * 1440
    
    'Set up the print instructions
    mFormatRange.hdc = ctlOrigPictureBox.hdc            ' Use the same DC for measuring and rendering
    mFormatRange.hdcTarget = ctlOrigPictureBox.hdc      ' Point at hDC
    mFormatRange.rectRegion = rectDrawTo                ' Area on page to draw to
    mFormatRange.rectPage = rectPage                    ' Entire size of page
    mFormatRange.mCharRange.firstChar = 0               ' Start of text
    mFormatRange.mCharRange.lastChar = -1               ' End of the text
    
    'Get current page
    CurrentPage = 1
    Do
        retValue = SendMessage(ctlSourceRichTextBox.hwnd, EM_FORMATRANGE, True, mFormatRange)
        If retValue >= Len(ctlSourceRichTextBox.Text) Then
            Exit Do
        End If
        If gPageNum = CurrentPage Then
            Exit Do
        End If
        ctlOrigPictureBox.Picture = LoadPicture()
        mFormatRange.mCharRange.firstChar = retValue
        mFormatRange.hdc = ctlOrigPictureBox.hdc
        mFormatRange.hdcTarget = ctlOrigPictureBox.hdc
        CurrentPage = CurrentPage + 1
        DoEvents
    Loop
    dumpaway = SendMessage(ctlOrigPictureBox.hdc, EM_FORMATRANGE, False, ByVal CLng(0))

    'Size to destination picture box from the original-size picture box
    ResizePreviewToPictureControl ctlDestPictureBox, ctlOrigPictureBox

End Function

Public Function CalculateNumberOfPages(ByRef ctlSourceRichTextBox As RichTextBox, ByVal PrintToDevice As Object, _
                                       Optional gLeftMargin As Long = 1, _
                                       Optional gRightMargin As Long = 1, _
                                       Optional gTopMargin As Long = 1, _
                                       Optional gBottomMargin As Long = 1) As Boolean
'
'Calculates the number of pages need to print document
'
Dim retValue As Long
Dim CurrentPage As Integer
Dim dumpaway As Long

    'Set printable area rect.
    rectPage.Left = 0
    rectPage.Top = 0
    rectPage.Right = ctlSourceRichTextBox.Width
    rectPage.Bottom = ctlSourceRichTextBox.Height
 
    'Set rect in which to print (relative to printable area)
    rectDrawTo.Left = gLeftMargin * 1440
    rectDrawTo.Top = gTopMargin * 1440
    rectDrawTo.Right = PrintToDevice.Width - gRightMargin * 1440
    rectDrawTo.Bottom = PrintToDevice.Height - gBottomMargin * 1440

    'Set up the print instructions
    mFormatRange.hdc = PrintToDevice.hdc                ' Use the same DC for measuring and rendering
    mFormatRange.hdcTarget = PrintToDevice.hdc          ' Point at hDC
    mFormatRange.rectRegion = rectDrawTo                ' Area on page to draw to
    mFormatRange.rectPage = rectPage                    ' Entire size of page
    mFormatRange.mCharRange.firstChar = 0               ' Start of text
    mFormatRange.mCharRange.lastChar = -1               ' End of the text

    'Set total number of pages
    gNumPages = 1
    Do
        retValue = SendMessage(ctlSourceRichTextBox.hwnd, EM_FORMATRANGE, True, mFormatRange)
        If retValue >= Len(ctlSourceRichTextBox.Text) Then
            Exit Do
        End If
        mFormatRange.mCharRange.firstChar = retValue
        mFormatRange.hdc = PrintToDevice.hdc
        mFormatRange.hdcTarget = PrintToDevice.hdc
        gNumPages = gNumPages + 1
        DoEvents
    Loop
    dumpaway = SendMessage(PrintToDevice.hdc, EM_FORMATRANGE, False, ByVal CLng(0))

End Function

Public Function SendToPrinter(ByRef ctlSourceRichTextBox As RichTextBox, _
                              Optional gLeftMargin As Long = 1, _
                              Optional gRightMargin As Long = 1, _
                              Optional gTopMargin As Long = 1, _
                              Optional gBottomMargin As Long = 1) As Boolean
Dim CurrentPage As Integer
Dim dumpaway    As Long
Dim retValue    As Long
Dim mFromPage   As Integer
Dim mToPage     As Integer

    On Error Resume Next
    
    'Ensure number of pages has been set
    CalculateNumberOfPages ctlSourceRichTextBox, Printer, gLeftMargin, gRightMargin, gTopMargin, gBottomMargin

    With frmParent.CommonDialog1
        .DialogTitle = "Print SQL"
        .CancelError = True
        .flags = 0
        .flags = cdlPDReturnDC + cdlPDNoSelection
        .Min = 1
        .Max = gNumPages
        .FromPage = 1
        .ToPage = gNumPages
        .ShowPrinter
        mFromPage = .FromPage
        mToPage = .ToPage
    End With
    
    DoEvents
    
    If Err = MSComDlg.cdlCancel Then
        'do nothing
    Else
        '***********************************************************************************************
        'Print to From Page
        '***********************************************************************************************
        
        'Reset printer
        Printer.Print ""
        Printer.ScaleMode = vbTwips
    
        'Set printable area rect.
        rectPage.Left = 0
        rectPage.Top = 0
        rectPage.Right = Printer.ScaleWidth
        rectPage.Bottom = Printer.ScaleHeight
     
        'Set rect in which to print (relative to printable area)
        rectDrawTo.Left = gLeftMargin * 1440
        rectDrawTo.Top = gTopMargin * 1440
        rectDrawTo.Right = Printer.ScaleWidth - gRightMargin * 1440
        rectDrawTo.Bottom = Printer.ScaleHeight - gBottomMargin * 1440
    
        'Set up the print instructions
        mFormatRange.hdc = Printer.hdc                      ' Use the same DC for measuring and rendering
        mFormatRange.hdcTarget = Printer.hdc                ' Point at hDC
        mFormatRange.rectRegion = rectDrawTo                ' Area on page to draw to
        mFormatRange.rectPage = rectPage                    ' Entire size of page
        mFormatRange.mCharRange.firstChar = 0               ' Start of text
        mFormatRange.mCharRange.lastChar = -1               ' End of the text
    
        CurrentPage = 1
        Do
            If mFromPage = CurrentPage Then
                Exit Do
            End If
            
            retValue = SendMessage(ctlSourceRichTextBox.hwnd, EM_FORMATRANGE, True, mFormatRange)
            
            mFormatRange.mCharRange.firstChar = retValue
            Printer.NewPage
            Printer.Print ""
            mFormatRange.hdc = Printer.hdc
            mFormatRange.hdcTarget = Printer.hdc
            CurrentPage = CurrentPage + 1
            DoEvents
        Loop
        dumpaway = SendMessage(Printer.hdc, EM_FORMATRANGE, False, ByVal CLng(0))
        
        If retValue >= Len(ctlSourceRichTextBox.Text) Then
            Exit Function
        End If
        
        '***********************************************************************************************
        'Print to To Page
        '***********************************************************************************************
        
        'Reset printer
        Printer.Print ""
        Printer.ScaleMode = vbTwips
    
        'Set printable area rect.
        rectPage.Left = 0
        rectPage.Top = 0
        rectPage.Right = Printer.ScaleWidth
        rectPage.Bottom = Printer.ScaleHeight
     
        'Set rect in which to print (relative to printable area)
        rectDrawTo.Left = gLeftMargin * 1440
        rectDrawTo.Top = gTopMargin * 1440
        rectDrawTo.Right = Printer.ScaleWidth - gRightMargin * 1440
        rectDrawTo.Bottom = Printer.ScaleHeight - gBottomMargin * 1440
    
        'Set up the print instructions
        mFormatRange.hdc = Printer.hdc                      ' Use the same DC for measuring and rendering
        mFormatRange.hdcTarget = Printer.hdc                ' Point at hDC
        mFormatRange.rectRegion = rectDrawTo                ' Area on page to draw to
        mFormatRange.rectPage = rectPage                    ' Entire size of page
        mFormatRange.mCharRange.firstChar = retValue        ' Start of text
        mFormatRange.mCharRange.lastChar = -1               ' End of the text
    
        Do
            retValue = SendMessage(ctlSourceRichTextBox.hwnd, EM_FORMATRANGE, True, mFormatRange)
            If retValue >= Len(ctlSourceRichTextBox.Text) Then
                Exit Do
            End If
            If mToPage = CurrentPage Then
                Exit Do
            End If
            
            mFormatRange.mCharRange.firstChar = retValue
            Printer.NewPage
            Printer.Print ""
            mFormatRange.hdc = Printer.hdc
            mFormatRange.hdcTarget = Printer.hdc
            CurrentPage = CurrentPage + 1
            DoEvents
        Loop
        dumpaway = SendMessage(Printer.hdc, EM_FORMATRANGE, False, ByVal CLng(0))
        
        Printer.EndDoc
        
    End If

End Function

Public Function ResizePreviewToPictureControl(ctlDestPictureBox As PictureBox, ctlOrigPictureBox As PictureBox) As Boolean
    Dim SrcX As Long, SrcY As Long
    Dim DestX As Long, DestY As Long
    Dim SrcWidth As Long, SrcHeight As Long
    Dim DestWidth As Long, DestHeight As Long
    Dim SrcHDC As Long, DestHDC As Long
    Dim mresult
      
    'Initialize
    SrcX = 0: SrcY = 0: DestX = 0: DestY = 0
    
    'Set Source width and height
    SrcWidth = ctlOrigPictureBox.ScaleWidth
    SrcHeight = ctlOrigPictureBox.ScaleHeight
    SrcHDC = ctlOrigPictureBox.hdc

    'Set destination width and height
    DestWidth = ctlDestPictureBox.ScaleWidth
    DestHeight = ctlDestPictureBox.ScaleHeight
    DestHDC = ctlDestPictureBox.hdc

    'Copy and stretch
    mresult = StretchBlt(DestHDC, DestX, DestY, DestWidth, DestHeight, SrcHDC, _
      SrcX, SrcY, SrcWidth, SrcHeight, vbSrcCopy)
      
    'Check for errors
    If mresult = 0 Then
        MsgBox "Error occurred in sizing images. Cannot continue"
        ResizePreviewToPictureControl = False
    Else
        ResizePreviewToPictureControl = True
    End If
    
    ctlDestPictureBox.Refresh
    
End Function

Public Function ViewPageSetup() As Long
'
'This procedure uses a Windows API call to display the Page Setup Dialog
'
Dim x As PAGESETUPDLG
Dim lngResult As Long

    'Set flags and initial margins
    x.flags = PSD_MARGINS + PSD_DISABLEORIENTATION + PSD_DISABLEPAPER
    x.rtMargin.Top = frmParent.ActiveForm.gMarginTop * 1000
    x.rtMargin.Left = frmParent.ActiveForm.gMarginLeft * 1000
    x.rtMargin.Right = frmParent.ActiveForm.gMarginRight * 1000
    x.rtMargin.Bottom = frmParent.ActiveForm.gMarginBottom * 1000
    
    'Initialize Structure
    x.lStructSize = Len(x)
    x.hwndOwner = frmParent.hwnd
    x.hInstance = App.hInstance
    
    'Windows API Call
    lngResult = PAGESETUPDLG(x)
        
    'If user didn't cancel, apply settings
    If lngResult <> 0 Then
    
        'Set Margins
        frmParent.ActiveForm.gMarginTop = x.rtMargin.Top / 1000
        frmParent.ActiveForm.gMarginLeft = x.rtMargin.Left / 1000
        frmParent.ActiveForm.gMarginRight = x.rtMargin.Right / 1000
        frmParent.ActiveForm.gMarginBottom = x.rtMargin.Bottom / 1000
        
    End If
    
    'Return
    ViewPageSetup = lngResult
        
End Function
