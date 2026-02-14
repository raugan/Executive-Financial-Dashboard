Sub ExportDashboardToPDF()
    '
    ' Export Dashboard as Board-Ready PDF
    ' This macro exports the Dashboard sheet as a high-quality PDF file
    '
    
    Dim ws As Worksheet
    Dim SavePath As String
    Dim FileName As String
    
    ' Set the Dashboard worksheet
    Set ws = ThisWorkbook.Sheets("Dashboard")
    
    ' Generate filename with current date
    FileName = "Financial_Dashboard_" & Format(Date, "YYYY-MM-DD") & ".pdf"
    
    ' Get save path (prompts user to select location)
    SavePath = Application.GetSaveAsFilename( _
        InitialFileName:=FileName, _
        FileFilter:="PDF Files (*.pdf), *.pdf", _
        Title:="Save Dashboard as PDF")
    
    ' Exit if user cancels
    If SavePath = "False" Then
        MsgBox "PDF export cancelled.", vbInformation
        Exit Sub
    End If
    
    ' Configure print settings for optimal PDF output
    With ws.PageSetup
        .Orientation = xlLandscape
        .FitToPagesWide = 1
        .FitToPagesTall = 1
        .PrintArea = "A1:P65"
        .CenterHorizontally = True
        .CenterVertically = True
        .PrintGridlines = False
        .PrintHeadings = False
        .LeftMargin = Application.InchesToPoints(0.5)
        .RightMargin = Application.InchesToPoints(0.5)
        .TopMargin = Application.InchesToPoints(0.5)
        .BottomMargin = Application.InchesToPoints(0.5)
    End With
    
    ' Export to PDF
    ws.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        FileName:=SavePath, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=True
    
    ' Success message
    MsgBox "Dashboard successfully exported to:" & vbCrLf & SavePath, vbInformation, "Export Complete"
    
End Sub

Sub QuickExportDashboardToPDF()
    '
    ' Quick Export to Desktop without prompting
    ' Saves PDF directly to Desktop folder
    '
    
    Dim ws As Worksheet
    Dim SavePath As String
    Dim FileName As String
    
    Set ws = ThisWorkbook.Sheets("Dashboard")
    
    ' Save to Desktop
    FileName = "Financial_Dashboard_" & Format(Date, "YYYY-MM-DD") & ".pdf"
    SavePath = CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\" & FileName
    
    ' Configure print settings
    With ws.PageSetup
        .Orientation = xlLandscape
        .FitToPagesWide = 1
        .FitToPagesTall = 1
        .PrintArea = "A1:P65"
        .CenterHorizontally = True
        .CenterVertically = True
        .PrintGridlines = False
        .PrintHeadings = False
    End With
    
    ' Export to PDF
    ws.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        FileName:=SavePath, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=True
    
    MsgBox "Dashboard exported to Desktop", vbInformation
    
End Sub

Sub FormatDashboard()
    '
    ' Apply final polish to Dashboard
    ' Removes gridlines and protects formula cells
    '
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Dashboard")
    
    ' Turn off gridlines
    ActiveWindow.DisplayGridlines = False
    
    ' Protect formulas (optional - uncomment if needed)
    ' ws.Protect Password:="", DrawingObjects:=False, Contents:=True, Scenarios:=False
    ' ws.Protect UserInterfaceOnly:=True
    
    MsgBox "Dashboard formatting applied", vbInformation
    
End Sub
