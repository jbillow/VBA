Option Explicit

Public oKill As String

Sub OpenFile(MyPath As String, name As String)

    Dim MyFile As String
    Dim LatestFile As String
    Dim LatestDate As Date
    Dim FDT As Date
    
    'Add backslash to end if not there
    If Right(MyPath, 1) <> "\" Then MyPath = MyPath & "\"

    If name = "" Then
        MsgBox "The file at " & MyPath & " not found"
        End
    End If
    
    'Get full file path and name
    MyFile = Dir(MyPath & name, vbNormal)
    
    ' for debugging, remove this if statement in deployed version
'    If Len(MyFile) = 0 Then
'
'        MsgBox "No files were found…"
'
'        Exit Sub
'
'    End If

    'Get most recently modified file
    Do While Len(MyFile) > 0
        FDT = FileDateTime(MyPath & MyFile)
        If FDT > LatestDate Then
            LatestFile = MyFile
            LatestDate = FDT
        End If

        MyFile = Dir

    Loop

    Workbooks.Open MyPath & LatestFile

End Sub
 
Sub Unzip(zipFilePath As Variant, destFolderPath As Variant)

    'Both must be Variant with late binding of Shell object
    
    Dim Sh As Object
    
    If Right(destFolderPath, 1) <> "\" Then destFolderPath = destFolderPath & "\"
    
    'Unzip all files in the .zip file
    Set Sh = CreateObject("Shell.Application")
    
    'Unzip all files inside the .zip file
    Sh.Namespace("" & destFolderPath).CopyHere Sh.Namespace("" & zipFilePath).Items
    
End Sub

Function GetLatestFile(MyPath As String, name As String) As String

    Dim MyFile As String
    Dim LatestFile As String
    Dim LatestDate As Date
    Dim FDT As Date
    
    'Add backslash to end if not there
    If Right(MyPath, 1) <> "\" Then MyPath = MyPath & "\"

    'Get full file path and name
    MyFile = Dir(MyPath & name, vbNormal)
    
    ' for debugging, remove this if statement in deployed version
'    If Len(MyFile) = 0 Then
'
'        MsgBox "No files were found…"
'
'        Exit Sub
'
'    End If

    'Get most recently modified file
    Do While Len(MyFile) > 0
        FDT = FileDateTime(MyPath & MyFile)
        If FDT > LatestDate Then
            LatestFile = MyFile
            LatestDate = FDT
        End If

        MyFile = Dir

    Loop

    GetLatestFile = MyPath & LatestFile
    
End Function

Function LastRow(Sh As Worksheet, Col As String) As Long

    'Function to find the last occupied row in a given column on a given worksheet
    
    LastRow = Sh.Cells(Sh.Rows.Count, Col).End(xlUp).Row

End Function

Sub InsertLookups(W As Workbook, W2 As Workbook, S As String, TargetCol As String, TargetRng As String, StRow As Long, StCol As Long, EndCol As Long)

    Dim k As Long
    'Copy the lookup table from tuner settings into the frequency tuner lookup sheet
    'stRow, stCol, k, and EndCol are the starting row, starting column, ending row, and ending column of the desired range, passed in as parameters
    W.Worksheets(S).Activate
    k = LastRow(ActiveSheet, TargetCol)
    ActiveSheet.Range(Cells(StRow, StCol), Cells(k, EndCol)).Select
    Selection.Copy
    W2.Worksheets("Lookups").Activate
    ActiveSheet.Range(TargetRng).Select
    ActiveSheet.Paste

End Sub

Sub SetUp(T As Workbook, B As Workbook, S As Workbook, A As Workbook)

    Dim k As Long

    'Delete old booked data
    T.Worksheets("Booked").Activate
    With ActiveSheet
        .Cells.Select
        Selection.ClearContents
    End With
'    'Filter for "JAM | JI_"
    B.Activate
'    With ActiveSheet
'        .Range("B1").AutoFilter Field:=2, Criteria1:="*JAM | JI_*"
'        .Cells.SpecialCells(xlCellTypeVisible).Select
'        Selection.Copy
'    End With
    ActiveSheet.Cells.Select
    Selection.Copy
    'Paste in tuner
    T.Worksheets("Booked").Activate
    With ActiveSheet
        .Range("A1").Select
        .Paste
    End With
    B.Close
    'Clear old lookups
    T.Worksheets("Lookups").Activate
    With ActiveSheet
        k = LastRow(ActiveSheet, "B")
        .Range(Cells(3, 1), Cells(k, 2)).Select
        Selection.ClearContents
        k = LastRow(ActiveSheet, "G")
        .Range(Cells(3, 7), Cells(k, 8)).Select
        Selection.ClearContents
        k = LastRow(ActiveSheet, "J")
        .Range(Cells(3, 10), Cells(k, 11)).Select
        Selection.ClearContents
    End With
    'Insert lookups
    Call InsertLookups(S, T, "Advertiser_Settings", "A", "A3", 7, 1, 2)
    Call InsertLookups(S, T, "Order_Overrides", "B", "J3", 8, 2, 3)
    Call InsertLookups(S, T, "Line_Item_Overrides", "B", "G3", 8, 2, 3)
    'Clear old DFP data and old calculations from report
    T.Worksheets("Report").Activate
    With ActiveSheet
        .Range("A:N").Select
        Selection.ClearContents
        k = LastRow(ActiveSheet, "O")
        .Range(Cells(12, 15), Cells(k, 43)).Select
        Selection.ClearContents
    End With
    'Copy data from Audience_Tuner_DFP report
    A.Activate
    With ActiveSheet
        .Range("A:N").Select
        Selection.Copy
    End With
    'Paste into report
    T.Worksheets("Report").Activate
    With ActiveSheet
        .Range("A1").Select
        .Paste
    End With

End Sub

Sub Execute(T As Workbook, S As Workbook, SavLoc As String)

    Dim k As Long
    Dim SourceRange As Range, FillRange As Range
    Dim err As Integer: err = 0
    
    'Put formulas in every row
    T.Worksheets("Report").Activate
    With ActiveSheet
        .Range("O10:AQ10").Select
        Selection.Copy
        .Range("O12").Select
        .Paste
        Set SourceRange = .Range("O12:AQ12")
        k = LastRow(ActiveSheet, "N")
        Set FillRange = .Range(Cells(12, 15), Cells(k, 43))
        SourceRange.AutoFill Destination:=FillRange
    End With

End Sub

Sub SendEmail(SendTo As String, Subject As String, Body As String, Optional T As Workbook)

    Dim OlApp As Object
    Dim NewMail As Object
    Dim Signature As String
    
    Set OlApp = CreateObject("Outlook.Application")
    Set NewMail = OlApp.CreateItem(0)
    
    'Get signature
    NewMail.Display
    Signature = NewMail.HTMLBody
    
    'Create email
    On Error Resume Next
    With NewMail
        .To = SendTo
        .Subject = Subject
        .HTMLBody = Body & Signature
        .Attachments.Add T.FullName
        .Send
    End With
    
    On Error GoTo 0
    Set NewMail = Nothing
    Set OlApp = Nothing

End Sub

Function ErrorCheck(T As Workbook, S As Workbook, Settings As String, Archive As String) As Integer

    Dim k As Long
    Dim i As Integer: i = 0
    Dim OrigSettingsName As String
    Dim temp As Workbook
    Dim SourceRange As Range, FillRange As Range
    Dim ShortName As String
    Dim pic As Range
    Dim img As String
    Dim Subject As String
    Dim SendTo As String
    Dim Body As String
    
    Call Shell(oKill)
    
    If Right(Archive, 1) <> "\" Then Archive = Archive & "\"
    
    OrigSettingsName = GetLatestFile(Settings, "*.xlsx*")
    
    T.Worksheets("Summary").Activate
    
    If ActiveSheet.Range("C20").Value <> 0 Then
        i = i + 1
        T.Worksheets("Report").Activate
        With ActiveSheet
            k = LastRow(ActiveSheet, "N")
            .Range("O11").AutoFilter Field:=15, Criteria1:="*Missing*"
            k = LastRow(ActiveSheet, "A")
            .Range(Cells(12, 1), Cells(k, 1)).Select
            Selection.SpecialCells(xlCellTypeVisible).Select
            Selection.Copy
        End With
        Set temp = Workbooks.Add(xlWBATWorksheet)
        temp.Activate
        With ActiveSheet
            .Range("A1").Select
            .Paste
            .Range("A:A").RemoveDuplicates Columns:=Array(1)
            If IsEmpty(.Range("A2").Value) = True Then
                .Range("B1").Value = "No"
                .Range("C1").Value = "No"
                .Range("D1").Value = "25%"
                .Range("E1").Value = "No"
            ElseIf IsEmpty(.Range("A2").Value) = False Then
                .Range("B1").Value = "No"
                .Range("C1").Value = "No"
                .Range("D1").Value = "25%"
                .Range("E1").Value = "No"
                Set SourceRange = .Range("B1:E1")
                k = LastRow(ActiveSheet, "A")
                Set FillRange = .Range(Cells(1, 2), Cells(k, 5))
                SourceRange.AutoFill Destination:=FillRange
            End If
            k = LastRow(ActiveSheet, "A")
            .Range(Cells(1, 1), Cells(k, 5)).Select
            Selection.Copy
        End With
        
        S.Worksheets("Advertiser_Settings").Activate
        With ActiveSheet
            k = LastRow(ActiveSheet, "A")
            .Cells(k + 1, 1).Select
            .Paste
        End With
        
        temp.Activate
        With ActiveSheet
            .Columns("A:A").AutoFit
            k = LastRow(ActiveSheet, "A")
            Set pic = .Range(Cells(1, 1), Cells(k, 1))
            pic.CopyPicture xlScreen
        End With
        
       'Create and export image
        With ActiveSheet.ChartObjects.Add(pic.Left, pic.Top, pic.Width, pic.Height)
            .Activate
            .Chart.Paste
            .Chart.Export Filename:=Environ("Userprofile") & "\Desktop\" & "image.jpg", FilterName:="JPG"
        End With
       
        ActiveSheet.ChartObjects(ActiveSheet.ChartObjects.Count).Delete
    
        Set pic = Nothing
    
        img = Environ("Userprofile") & "\Desktop\" & "image.jpg"
        
        SendTo = ThisWorkbook.Worksheets(1).Range("C18").Value
        
        Body = "<BODY Style=font-size:11pt;font-family:Calibri> Hello team, <br><br> New advertisers have been detected in today's Audience Tuner:" _
               & "<br><br>" _
               & "<img src=" & img & ">" _
               & "<br><br>" _
               & "<BODY Style=font-size:11pt;font-family:Calibri> They have been EXCLUDED in today's tuner, and will continue to be excluded from all tuners until further action is taken. If these are your advertisers, please update Tuner Settings here: <br><br>https://jumpstart.app.box.com/folder/66111489049<br><br>Please make sure to update the filename with todays date and your initials, and move the old file to the archive folder.<br><br>Thanks,"
        
        Subject = "ATTN: New Advertisers Detected"
        
        Call SendEmail(SendTo, Subject, Body)
        
        Kill img
        
        Call SaveFile(S, "Tuner Settings", "_NEEDS UPDATE", Settings)
        S.Close
        temp.Close
        
        'Archive old settings
        ShortName = Right(OrigSettingsName, Len(OrigSettingsName) - InStrRev(OrigSettingsName, "\"))
        Name OrigSettingsName As Archive & ShortName
    End If

    ErrorCheck = i

End Function

Sub SaveFile(W As Workbook, name1 As String, name2 As String, SaveLoc As String)
    
    Application.DisplayAlerts = False
    Dim d As Variant
    Dim fName As String
    
    'Get the system date
    d = Format(Now(), "yyyy-mm-dd")
    
    'Create the file name
    fName = name1 & "_" & d & name2
    'MsgBox fName
    
    'Add backslash to end of save location if not there
    If Right(SaveLoc, 1) <> "\" Then SaveLoc = SaveLoc & "\"

    'Save file
    W.SaveAs Filename:=SaveLoc & fName
    
End Sub

Sub CreateUpload(T As Workbook, U As Workbook, SavLoc As String)

    Dim k As Long
    Dim SortRange As Range
    
    'Delete old values from upload sheet
    U.Activate
    With ActiveSheet
        k = LastRow(ActiveSheet, "A")
        .Range(Cells(2, 1), Cells(k, 3)).Select
        Selection.EntireRow.Delete
    End With
    
    'Sort for valid upload
    T.Worksheets("Report").Activate
    With ActiveSheet
        .Range("AB11").AutoFilter Field:=28, Criteria1:="*Yes*"
        'Copy upload data
        k = LastRow(ActiveSheet, "AB")
        .Range(Cells(12, 41), Cells(k, 42)).Select
        Selection.Copy
    End With
    
    'Paste upload data
    U.Activate
    With ActiveSheet
        .Range("A2").Select
        .Paste
    End With
    
    'Save Upload
    Call SaveFile(U, "DFP Audience_Tuner_Upload", "", SavLoc)

End Sub

Sub Prepare(T As Workbook, SavLoc As String)

    'Save Template
    Call SaveFile(T, "DFP Audience_Tuner_Template", "", SavLoc)
    
    T.Worksheets("Report").Activate
    With ActiveSheet
        .Cells.Select
        Selection.Copy
        .Range("A1").PasteSpecial Paste:=xlPasteValues
    End With
    
    With T
        Worksheets("Instructions").Delete
        Worksheets("Booked").Delete
    End With
    
    'Save Report File
    Call SaveFile(T, "DFP Audience_Tuner_Report", "", SavLoc)
    
End Sub

Sub PrepareEmail(T As Workbook)

    Dim pic As Range
    Dim img As String
    Dim Subject As String
    Dim SendTo As String
    Dim Body As String

    T.Worksheets("Summary").Activate
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    With ActiveSheet
        Set pic = .Range("A1:D23")
        pic.CopyPicture xlScreen, xlPicture
    End With
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    With ActiveSheet.ChartObjects.Add(pic.Left, pic.Top, pic.Width, pic.Height)
        .Activate
        .Chart.Paste
        .Chart.Export Filename:=Environ("Userprofile") & "\Desktop\" & "image.jpg", FilterName:="JPG"
    End With
    
    ActiveSheet.ChartObjects(ActiveSheet.ChartObjects.Count).Delete
    
    Set pic = Nothing
    
    img = Environ("Userprofile") & "\Desktop\" & "image.jpg"
    
    SendTo = ThisWorkbook.Worksheets(1).Range("C18").Value
    
    Body = "<BODY Style=font-size:14pt;font-family:Calibri> Attached is today's DFP Audience Tuner:<br>" _
            & "<br>" _
            & "<img src=" & img & ">" _
            & "<br>" _
    
    Subject = ActiveWorkbook.name
    
    Call SendEmail(SendTo, Subject, Body, ActiveWorkbook)
    
    Kill img
    
End Sub

Sub Archive()
    
    'Moves any file that does not contain todays date to the archive folder
    Dim d As String: d = Format(Now(), "yyyy-mm-dd")
    Dim strFileName As String
    Dim strFolder As String: strFolder = ThisWorkbook.Worksheets(1).Range("C4").Value
    Dim strArchive As String: strArchive = ThisWorkbook.Worksheets(1).Range("C22").Value
    If Right(strFolder, 1) <> "\" Then strFolder = strFolder & "\"
    If Right(strArchive, 1) <> "\" Then strArchive = strArchive & "\"
    Dim strFileSpec As String: strFileSpec = strFolder & "*.*"
    Dim cont As Boolean
    
    strFileName = Dir(strFileSpec)
    
    Do While Len(strFileName) > 0
        cont = Contains(strFileName, d)
        If cont Then
            'Do nothing
        Else
            Name strFolder & strFileName As strArchive & strFileName 'do nothing
        End If
        
        strFileName = Dir
    Loop
    
End Sub

Function Contains(ByVal string_source As String, ByVal find_text As String, Optional ByVal caseSensitive As Boolean = False) As Boolean

    'Checks if a string contains specified text
    Dim compareMethod As VbCompareMethod

    If caseSensitive Then
        compareMethod = vbBinaryCompare
    Else
        compareMethod = vbTextCompare
    End If

    Contains = (InStr(1, string_source, find_text, compareMethod) <> 0)

End Function

Sub Main()
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    With ActiveSheet
        'Get all needed file paths
        Dim Path1 As String: Path1 = .Range("C4").Value 'tuner files
        Dim Path2 As String: Path2 = .Range("C5").Value 'Audience Tuner DFP
        Dim Path3 As String: Path3 = .Range("C6").Value 'Tuner Settings
        Dim Path4 As String: Path4 = .Range("C7").Value 'Booked GAM
        Dim Path5 As String: Path5 = .Range("C8").Value 'Unzip location
        Dim Path6 As String: Path6 = .Range("C22").Value 'Tuner archive
        Dim Path7 As String: Path7 = .Range("C24").Value 'Tuner settings archive
        
        'Get all needed file name wildcards
        Dim name1 As String: name1 = .Range("C10").Value 'Template
        Dim name2 As String: name2 = .Range("C11").Value 'Upload
        Dim Name3 As String: Name3 = .Range("C12").Value 'Settings
        Dim Name4 As String: Name4 = .Range("C14").Value 'Audience Tuner DFP
        Dim Name5 As String: Name5 = .Range("C13").Value 'Booked GAM
        
        'Get the save location
        Dim SavLoc As String: SavLoc = .Range("C16").Value
        
        'Set outlook kill
        oKill = .Range("C26").Value
    End With
    
    'Variable used for error checking
    Dim err As Integer
    
    'Variables to store each workbook
    Dim Tuner As Workbook
    Dim Upload As Workbook
    Dim Audience As Workbook
    Dim Settings As Workbook
    Dim Booked As Workbook
    
    'Variable for unzipping GAM line items report
    Dim MyFile As String
    
    'Close outlook
    Call Shell(oKill)

    'Open Template
    Call OpenFile(Path1, name1)
    Set Tuner = ActiveWorkbook

    'Open Upload
    Call OpenFile(Path1, name2)
    Set Upload = ActiveWorkbook

    'Open Audience_Tuner_DFP report
    Call OpenFile(Path2, Name4)
    Set Audience = ActiveWorkbook

    'Open Tuner Settings
    Call OpenFile(Path3, Name3)
    Set Settings = ActiveWorkbook
    
   'Unzip booked GAM report
    MyFile = GetLatestFile(Path4, Name5)
    Call Unzip(MyFile, Path5)
    'Open unzipped file
    Call OpenFile(Path5, Name5)
    Set Booked = ActiveWorkbook
    
    'Set up
    Call SetUp(Tuner, Booked, Settings, Audience)
    
    'Execute
    Call Execute(Tuner, Settings, SavLoc)
    
    'Error check
    err = ErrorCheck(Tuner, Settings, Path3, Path7)
    
    If err <> 0 Then
        Tuner.Close
        Call OpenFile(Path3, Name3)
        Set Settings = ActiveWorkbook
        Call OpenFile(Path1, name1)
        Set Tuner = ActiveWorkbook
        Call OpenFile(Path5, Name5)
        Set Booked = ActiveWorkbook
        Call SetUp(Tuner, Booked, Settings, Audience)
        Call Execute(Tuner, Settings, SavLoc)
    End If
    
    'Prepare for upload creation
    Call Prepare(Tuner, SavLoc)
    
    'Create Upload
    Call CreateUpload(Tuner, Upload, SavLoc)
    
    'Prepeare and send email
    Call PrepareEmail(Tuner)
    
    'Archive Files
    Call Archive
    
    Tuner.Close
    Upload.Close
    Audience.Close
    Settings.Close
    Booked.Close
    
End Sub
