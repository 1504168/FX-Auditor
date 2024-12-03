Attribute VB_Name = "SubCollection"
'@GithubRawURL: https://raw.githubusercontent.com/1504168/All-Personal-VBA-Code/master/Reusable%20Code/SubCollection.bas

'@Folder("Reusable.Sub")
Option Explicit

#If Not Mac Then
    #If VBA7 Then                                ' Excel 2010 or later
        Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal Milliseconds As LongPtr)
    #Else                                        ' Excel 2007 or earlier
        Public Declare Sub Sleep Lib "kernel32" (ByVal Milliseconds As Long)
    #End If
#End If

Public Sub SortArray(ByRef InputArr As Variant _
                     , Optional ByVal LowIndex As Long = -1 _
                      , Optional ByVal HighIndex As Long = -1 _
                       , Optional ByVal Bycolumn As Long = -1)

    Dim Pivot As Variant                         'This variable for taking the pivot value.
    Dim i As Long, j As Long
    Dim TemporaryColumn As Long
    If IsValidArray(InputArr) = False Then Exit Sub

    'Setup index value.
    If LowIndex = -1 Then LowIndex = LBound(InputArr, 1)
    If HighIndex = -1 Then HighIndex = UBound(InputArr, 1)
    If Bycolumn = -1 Then Bycolumn = LBound(InputArr, 2)
    If LowIndex >= HighIndex Then Exit Sub

    Pivot = InputArr((LowIndex + HighIndex) / 2, Bycolumn)
    i = LowIndex
    j = HighIndex
    
    'Check variable type of the pivot..0 for empty,1 for null,9 for object,10 for error. >17 some are available for 64 bit and user defined data type
    Select Case VarType(Pivot)
        Case 0, 1, 9, 10, 13, Is > 17
            
            ' Change the index to higher so that recursive will not run.
            i = HighIndex
            ' Change the index to lower.
            j = LowIndex
            
    End Select

    While i <= j
        
        ' Check for least portion of array..till it less than then increase the value of i
        While InputArr(i, Bycolumn) < Pivot And i < HighIndex
            i = i + 1
        Wend
        
        ' Check for greater portion of array..till it greater than then decrease the value of j
        While Pivot < InputArr(j, Bycolumn) And j > LowIndex
            j = j - 1
        Wend

        If i <= j Then

            For TemporaryColumn = LBound(InputArr, 2) To UBound(InputArr, 2)
                SwapItem i, j, TemporaryColumn, InputArr
            Next TemporaryColumn
            ' Increase index
            i = i + 1
            ' Decrease index
            j = j - 1

        End If

    Wend
    
    ' Run for less section of the pivot.
    If LowIndex < j Then SortArray InputArr, LowIndex, j, Bycolumn
    ' Run for greater section of the pivot.
    If i < HighIndex Then SortArray InputArr, i, HighIndex, Bycolumn

End Sub

Private Sub SwapItem(i As Long, j As Long, TemporaryColumn As Long, ByRef InputArray As Variant)
    
    Dim TemporaryElement As Variant
    ' Hold the value to a temporary place.
    AssignProperly InputArray(i, TemporaryColumn), TemporaryElement
    AssignProperly InputArray(j, TemporaryColumn), InputArray(i, TemporaryColumn)
    AssignProperly TemporaryElement, InputArray(j, TemporaryColumn)

End Sub

Private Sub AssignProperly(ByRef FromItem As Variant, ByRef ToItem As Variant)

    If IsObject(FromItem) Then
        Set ToItem = FromItem
    Else
        ToItem = FromItem
    End If

End Sub

Private Function IsValidArray(InputArray As Variant) As Boolean

    Const ArrayIdentifier As String = "()"
    If IsEmpty(InputArray) Then
        IsValidArray = False
    ElseIf InStr(TypeName(InputArray), ArrayIdentifier) < 1 Then
        'if array is integer type then typeName returns Integer() so ArrayIdentifier should be present atleast greater then 1th _
        place otherwise broken array.
        IsValidArray = False
    Else
        IsValidArray = True
    End If

End Function

Public Sub PrintArray(ByVal GivenArray As Variant)

    Dim FirstColumnIndex  As Long
    FirstColumnIndex = LBound(GivenArray, 2)
    Dim FirstRowIndex As Long
    FirstRowIndex = LBound(GivenArray, 1)
    Dim CurrentRowIndex As Long
    Dim OutputText As String
    For CurrentRowIndex = LBound(GivenArray, 1) To UBound(GivenArray, 1)
        Dim CurrentColumnIndex As Long
        For CurrentColumnIndex = LBound(GivenArray, 2) To UBound(GivenArray, 2)
            OutputText = OutputText & vbTab & GivenArray(CurrentRowIndex, CurrentColumnIndex)
        Next CurrentColumnIndex
        Debug.Print Right(OutputText, Len(OutputText) - Len(vbTab))
        OutputText = vbNullString
    Next CurrentRowIndex

End Sub

Public Sub LinkSheets(Optional GivenWorkbook As Workbook)

    If GivenWorkbook Is Nothing Then
        GivenWorkbook = ThisWorkbook
    End If

    Const IndexSheetName As String = "Index Sheet"
    Dim isIndexSheetExist As Boolean
    isIndexSheetExist = IsSheetExist(IndexSheetName, GivenWorkbook)

    Dim LinkHolderSheet As Worksheet
    If isIndexSheetExist Then
        Set LinkHolderSheet = GivenWorkbook.Worksheets(IndexSheetName)
        LinkHolderSheet.Cells.Clear
    Else
        AddNewSheet IndexSheetName, GivenWorkbook, 1
        Set LinkHolderSheet = GivenWorkbook.Worksheets(IndexSheetName)
    End If

    LinkHolderSheet.Range("A1").Value = "Sheet Name"

    Dim currentLinkHolderRow As Long
    currentLinkHolderRow = 2

    Dim CurrentWorksheet As Worksheet
    For Each CurrentWorksheet In GivenWorkbook.Worksheets
        If CurrentWorksheet.Name <> LinkHolderSheet.Name Then
            'Anchor = where you want the link.
            'Screen tip = when you Hoover your mouse to that link you will see this.
            'Text To Display = This one is appeared in the cell with underline.
            'Sub address = adreess of the link sheet and cell.
            LinkHolderSheet.Hyperlinks.Add Anchor:=LinkHolderSheet.Cells(currentLinkHolderRow, "A"), Address:=vbNullString, _
                                           SubAddress:="'" & CurrentWorksheet.Name & "'!A1", _
                                           ScreenTip:="Click to go to " & CurrentWorksheet.Name & " Sheet ", _
                                           TextToDisplay:=CurrentWorksheet.Name

            currentLinkHolderRow = currentLinkHolderRow + 1 'For next cell address

        End If
    Next CurrentWorksheet

    AutoFitRangeCols LinkHolderSheet.Columns("A:A"), 50

End Sub

Private Sub AddNewSheet(ByVal SheetName As String, Optional GivenWorkbook As Workbook, _
                        Optional IndexNumber As Long)

    Dim TemporarySheet As Worksheet
    If GivenWorkbook Is Nothing Then
        Set TemporarySheet = ThisWorkbook.Worksheets.Add(ThisWorkbook.Worksheets(IndexNumber))
    Else
        Set TemporarySheet = GivenWorkbook.Worksheets.Add(GivenWorkbook.Worksheets(IndexNumber))
    End If
    TemporarySheet.Name = SheetName
    Set TemporarySheet = Nothing

End Sub

Sub PrintTableNameAsConstDeclaration()
    Dim CurrentTable As ListObject
    Dim CurrentSheet As Worksheet
    For Each CurrentSheet In ThisWorkbook.Worksheets
        For Each CurrentTable In CurrentSheet.ListObjects
            Debug.Print "Public Const " & CurrentTable.Name & " As String =""" & CurrentTable.Name & """"
        Next CurrentTable
    Next CurrentSheet
End Sub

'@Example call EventOption IsEnable:=True for enable and False for disable

Public Sub EventOption(ByVal IsEnable As Boolean)
    Application.ScreenUpdating = IsEnable
    Application.EnableEvents = IsEnable
    If IsEnable Then
        Application.Calculation = xlCalculationAutomatic
    Else
        Application.Calculation = xlCalculationManual
    End If
End Sub

'@Description("This will toggle the events. Like set to manual and when done reset back to previous status. Need to call in pair to work properly")
'@Dependency("No Dependency")
'@ExampleCall :
'@Date : 26 April 2022 09:46:40 AM
'@PossibleError :

Public Static Sub ToggleEvents()

    Dim DisplayAlertStatus As Boolean
    Dim ScreenUpdatingStatus As Boolean
    Dim EnableEventStatus As Boolean
    Dim CalculationStatus As XlCalculation
    Dim Counter As Long
    Counter = Counter + 1
    If Counter Mod 2 = 1 Then
        DisplayAlertStatus = Application.DisplayAlerts
        ScreenUpdatingStatus = Application.ScreenUpdating
        EnableEventStatus = Application.EnableEvents
        CalculationStatus = Application.Calculation
        Application.DisplayAlerts = False
        Application.ScreenUpdating = False
        Application.EnableEvents = False
        Application.Calculation = xlCalculationManual
    Else
        Application.DisplayAlerts = DisplayAlertStatus
        Application.ScreenUpdating = ScreenUpdatingStatus
        Application.EnableEvents = EnableEventStatus
        Application.Calculation = CalculationStatus
    End If

End Sub

Public Sub CopyDataToClipBoard(ByVal GivenText As String)
    
    Dim TotalWait As Long
    On Error GoTo HandleError
    CreateObject("htmlfile").parentWindow.clipboardData.SetData "text", GivenText
    Exit Sub
HandleError:
    If Err.Number = -2147352319 And Err.Description = "Automation error" Then
        Debug.Print "Wait for one sec."
        TotalWait = TotalWait + 1
        Sleep 1000
        If TotalWait > 5 Then Exit Sub
        DoEvents
        Resume
    End If
    
End Sub

Public Sub DeleteOrInsertRow(HowManyRow As Long, StartAtRow As Long, InSheet As Worksheet)

    Dim Counter As Long
    With InSheet

        If HowManyRow > 0 Then
            For Counter = 1 To HowManyRow
                .Rows(StartAtRow).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
            Next Counter
        ElseIf HowManyRow < 0 Then
            For Counter = 0 To HowManyRow Step -1
                .Rows(StartAtRow + Counter).Delete Shift:=xlUp
            Next Counter
        End If

    End With

End Sub

'@Description : This is for moving one sheet to another workbook.
'@It will always add at the end.
Private Sub MoveSheet(GivenSheet As Worksheet, ToWorkbook As Workbook)
    If Not GivenSheet Is Nothing Then
        Dim LastSheet As Worksheet
        Set LastSheet = ToWorkbook.Worksheets(ToWorkbook.Worksheets.Count)
        GivenSheet.Move After:=LastSheet
    End If
End Sub

'@Description: This will write collection item to a text file .
'@IsAppend=True : Means add data in to the text file instead of overwriting that.
'@IsAppend=False : Overwrite
'@If file doesn't exit then it will create a new text file by that name.
'@ExampleCall: SubCollection.WriteCollectionDataIntoTextFile "C:\Users\Ismail\Desktop\Test\Read Text File Project\Test.txt", V, False
'@ExampleCall: SubCollection.WriteCollectionDataIntoTextFile "C:\Users\Ismail\Desktop\Test\Read Text File Project\Test.txt", V, True

Public Sub WriteCollectionDataIntoTextFile(FullFileName As String, GivenCollection As Collection, Optional IsAppend As Boolean = False)

    Dim FileNo As Long
    FileNo = FreeFile()
    If IsAppend Then
        Open FullFileName For Append As #FileNo
    Else
        Open FullFileName For Output As #FileNo
    End If
    Dim CurrentRowText As Variant
    For Each CurrentRowText In GivenCollection
        Print #FileNo, CurrentRowText
    Next CurrentRowText
    Close #FileNo

End Sub

'@Description: This will write array data into a text file. Make sure array has only one column otherwise it will make each cell into a single line.
'@ExampleCall: ExportArrayToTextFile AllData, "C:\Users\Ismail\Desktop\Test\Output.txt"
Public Sub ExportArrayToTextFile(ByVal AllData As Variant, ByVal FullFileName As String, Optional Delimiter As String = vbTab)

    Dim FileNo As Long
    FileNo = FreeFile()
    Open FullFileName For Output As #FileNo
    Dim LastCol As Long
    LastCol = UBound(AllData, 2)
    Dim CurrentRowText As String
    Dim CurrentRowIndex As Long
    For CurrentRowIndex = LBound(AllData, 1) To UBound(AllData, 1)
        Dim CurrentColumnIndex As Long
        For CurrentColumnIndex = LBound(AllData, 2) To UBound(AllData, 2)
            CurrentRowText = CurrentRowText & AllData(CurrentRowIndex, CurrentColumnIndex)
            If CurrentColumnIndex <> LastCol Then CurrentRowText = CurrentRowText & Delimiter
        Next CurrentColumnIndex
        Print #FileNo, CurrentRowText
        CurrentRowText = vbNullString
    Next CurrentRowIndex
    Close #FileNo

End Sub

Public Sub WriteStringToTextFile(Content As String, ToFilePath As String)
    
    Dim FileNo As Long
    FileNo = FreeFile()
    Open ToFilePath For Output As #FileNo
    Print #FileNo, Content
    Close #FileNo
        
End Sub

'@Description: This sub will print all the item of a collection in debug window.
'@Only valid for normal data type(except object data type)
Public Sub PrintCollection(GivenCollection As Collection)

    Dim CurrentElement As Variant
    Dim Counter As Long
    For Each CurrentElement In GivenCollection
        Counter = Counter + 1
        Debug.Print "Element Index : " & Counter & vbTab & "Item : " & CurrentElement
    Next CurrentElement

End Sub

''@Description : It will create folder recursively.
''@Dependency : Microsoft Scripting Runtime
''@Early Binding
''@ExampleCall : CreateFolderRecursively "C:\Users\Ismail\Desktop\Test\New folder\SubFolder\AnotherSubFolder"
'Public Sub CreateFolderRecursively(ByVal GivenFolderPathString As String)
'
'    Dim FolderManager As Scripting.FileSystemObject
'    Set FolderManager = New Scripting.FileSystemObject
'
'    Dim SubFolderList As Variant
'    SubFolderList = Split(GivenFolderPathString, Application.PathSeparator)
'
'    Dim CurrentSubFolderPath As String
'    CurrentSubFolderPath = FolderManager.GetDriveName(GivenFolderPathString)
'    Dim CurrentRowIndex As Long
'    For CurrentRowIndex = LBound(SubFolderList, 1) + 1 To UBound(SubFolderList, 1)
'        CurrentSubFolderPath = CurrentSubFolderPath & Application.PathSeparator & SubFolderList(CurrentRowIndex)
'        If Not FolderManager.FolderExists(CurrentSubFolderPath) Then
'            FolderManager.CreateFolder (CurrentSubFolderPath)
'        End If
'    Next CurrentRowIndex
'    MsgBox "Folder Created Successfully"
'
'End Sub


''@Description : It will create folder recursively.
''@Dependency : Microsoft Scripting Runtime
''@Late Binding
''@ExampleCall : CreateFolderRecursively "C:\Users\Ismail\Desktop\Test\New folder\SubFolder\AnotherSubFolder"
'Public Sub CreateFolderRecursively(ByVal GivenFolderPathString As String)
'
'    Dim FolderManager As Object
'    Set FolderManager = CreateObject("Scripting.FileSystemObject")
'
'    Dim SubFolderList As Variant
'    SubFolderList = Split(GivenFolderPathString, Application.PathSeparator)
'
'    Dim CurrentSubFolderPath As String
'    CurrentSubFolderPath = FolderManager.GetDriveName(GivenFolderPathString)
'    Dim CurrentRowIndex As Long
'    For CurrentRowIndex = LBound(SubFolderList, 1) + 1 To UBound(SubFolderList, 1)
'        CurrentSubFolderPath = CurrentSubFolderPath & Application.PathSeparator & SubFolderList(CurrentRowIndex)
'        If Not FolderManager.FolderExists(CurrentSubFolderPath) Then
'            FolderManager.CreateFolder (CurrentSubFolderPath)
'        End If
'    Next CurrentRowIndex
'    MsgBox "Folder Created Successfully"
'
'End Sub


'@Description : It will create folder recursively.
'@Dependency : No Dependency
'@ExampleCall : CreateFolderRecursively "C:\Users\Ismail\Desktop\Test\New folder\SubFolder\AnotherSubFolder"
Public Sub CreateFolderRecursively(ByVal GivenFolderPathString As String)

    Dim SubFolderList As Variant
    SubFolderList = Split(GivenFolderPathString, Application.PathSeparator)

    Dim CurrentSubFolderPath As String
    CurrentSubFolderPath = SubFolderList(LBound(SubFolderList, 1))
    Dim CurrentRowIndex As Long
    For CurrentRowIndex = LBound(SubFolderList, 1) + 1 To UBound(SubFolderList, 1)
        CurrentSubFolderPath = CurrentSubFolderPath & Application.PathSeparator & SubFolderList(CurrentRowIndex)
        If Dir(CurrentSubFolderPath, vbDirectory) = vbNullString Then
            MkDir CurrentSubFolderPath
        End If
    Next CurrentRowIndex
    'MsgBox "Folder Created Successfully"

End Sub

'@Description("It will Create folder if that folder is not exist")
'@Dependency("Microsoft Scripting Runtime")
'@ExampleCall : CreateFolderIfNotExist "C:\Users\USER\Desktop\VBA Code Automation\Reusable Code"
'@Date : 13 October 2022 09:43:42 PM
Private Sub CreateFolderIfNotExist(FolderPath As String)

    Dim FolderManager As Object
    Set FolderManager = CreateObject("Scripting.FileSystemObject")
    If Not FolderManager.FolderExists(FolderPath) Then
        FolderManager.Createfolder FolderPath
    End If
    Set FolderManager = Nothing

End Sub

'@Description("It will format the table in a defined way")
'@Dependency("No Dependency")
'@ExampleCall : FormatTable Sheet1.ListObjects("TableName")
'@Date : 18 October 2021 01:31:23 PM
Public Sub FormatTable(GivenTable As ListObject)

    With GivenTable.Range

        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlInsideVertical).LineStyle = xlContinuous
        .Borders(xlInsideHorizontal).LineStyle = xlContinuous
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Name = "Times New Roman"
        .Font.Size = 14
        AutoFitRangeCols GivenTable.Range, 50
    End With
    GivenTable.ShowAutoFilterDropDown = False

End Sub

'@Description("This is for logging message into debug window and in status bar")
'@Dependency("No Dependency")
'@ExampleCall : LogMessage "This is just a sample test"
'@Date : 23 November 2021 08:48:22 PM
Public Sub LogMessage(Message As String)
    Application.StatusBar = Message
    Debug.Print Message
End Sub

'@Description("Replace a picture with new picture")
'@Dependency("No Dependency")
'@ExampleCall : UpdatePicture "TestPicture", LogsSheet, "C:\Users\Ismail\Desktop\BCS\DSC_1736.jpg"
'@Date : 03 December 2021 11:08:04 AM
Public Sub UpdatePicture(PreviousPictureName As String, PictureInSheet As Worksheet, NewPicturePath As String)
    
    With PictureInSheet
        Dim OldPicture As Shape
        Set OldPicture = .Shapes(PreviousPictureName)
        Dim NewPicture As Shape
        Set NewPicture = .Shapes.AddPicture(NewPicturePath, False, True, OldPicture.Left, OldPicture.Top, OldPicture.Width, OldPicture.Height)
        OldPicture.Delete
        NewPicture.Name = PreviousPictureName
    End With
    
End Sub

'@Description("This will convert to array if not an array. It will be 2D Array. It will change the input Data.")
'@Dependency("No Dependency")
'@ExampleCall : ConvertToArrayIfNotArray GivenArray
'@Date : 24 December 2021 07:42:36 PM
Public Sub ConvertToArrayIfNotArray(ByRef GivenData As Variant, Optional ColumnIndex As Long = 1)
    
    If IsArray(GivenData) Then
        Exit Sub
    Else
        Dim Result As Variant
        ReDim Result(1 To 1, ColumnIndex To ColumnIndex)
        Result(1, ColumnIndex) = GivenData
        GivenData = Result
    End If
    
End Sub

'@Description("This will print Vector or 1D array.")
'@Dependency("FunctionCollection.ConcatenateVector")
'@ExampleCall : Print1DArray GivenArray
'@Date : 26 December 2021 01:36:09 AM
Public Sub Print1DArray(ByVal GivenArray As Variant)
    Debug.Print ConcatenateVector(GivenArray)
End Sub

'@Description("This will delete all the query links from given workbook")
'@Dependency("")
'@ExampleCall : DeleteQueriesFromWorkbook DeleteFromWorkbook:=GivenWorkbook
'@Date : 12 January 2022 08:32:04 PM
'@PossibleError :
Public Sub DeleteQueriesFromWorkbook(DeleteFromWorkbook As Workbook)

    Dim CurrentQuery As WorkbookQuery

    For Each CurrentQuery In DeleteFromWorkbook.Queries
        CurrentQuery.Delete
    Next CurrentQuery

End Sub

'@Description("This will delete all the query table links from given sheet")
'@Dependency("No Dependency")
'@ExampleCall : DeleteQueryTablesFromWorksheet DeleteFromSheet:=CurrentSheet
'@Date : 12 January 2022 08:31:29 PM
'@PossibleError :
Public Sub DeleteQueryTablesFromWorksheet(DeleteFromSheet As Worksheet)

    Dim CurrentQuery As QueryTable
    For Each CurrentQuery In DeleteFromSheet.QueryTables
        CurrentQuery.Delete
    Next CurrentQuery

End Sub

'@Description("This will delete all the query links from given workbook")
'@Dependency("SubCollection.DeleteQueryTablesFromWorksheet")
'@ExampleCall : DeleteQueryTablesFromWorkbook DeleteFromWorkbook:=GivenWorkbook
'@Date : 12 January 2022 08:32:04 PM
'@PossibleError :
Public Sub DeleteQueryTablesFromWorkbook(DeleteFromWorkbook As Workbook)

    Dim CurrentSheet As Worksheet
    For Each CurrentSheet In DeleteFromWorkbook.Worksheets
        SubCollection.DeleteQueryTablesFromWorksheet DeleteFromSheet:=CurrentSheet
    Next CurrentSheet

End Sub

'@Description("This will remove all the item from the collection.")
'@Dependency("")
'@ExampleCall :SubCollection.RemoveAllFromCollection GivenCollection
'@Date : 12 January 2022 08:32:04 PM
'@PossibleError :
Public Sub RemoveAllFromCollection(GivenCollection As Collection)
    
    Do While GivenCollection.Count > 0
        GivenCollection.Remove 1
    Loop
    
End Sub

'@Description("This will save the given file as csv in the given full path(Folder+FileName)")
'@Dependency("No Dependency")
'@ExampleCall : SaveSheetAsCSV ActiveSheet,"D:\Downloads\Test.csv"
'@Date : 07 March 2022 09:02:23 PM
'@PossibleError :

Public Sub SaveSheetAsCSV(CurrentSheet As Worksheet, FullFilePath As String)
    
    Dim NewWorkbook As Workbook
    Set NewWorkbook = Application.Workbooks.Add
    CopySheet CurrentSheet, NewWorkbook
    Application.DisplayAlerts = False
    NewWorkbook.Worksheets(1).Delete
    NewWorkbook.SaveAs FullFilePath, xlCSV
    NewWorkbook.Close
    Application.DisplayAlerts = True
    
End Sub

'@Description("This will copy down item from one collection to another.")
'@Dependency("No Dependency")
'@ExampleCall : CopyFromOneCollectionToAnother InvalidDataRowsForCurrentClient, InvalidDataRows
'@Date : 18 March 2022 01:02:07 AM
'@PossibleError :
Public Sub CopyFromOneCollectionToAnother(FromCollection As Collection, ToCollection As Collection)

    If FromCollection Is Nothing Or ToCollection Is Nothing Then Exit Sub
    Dim CurrentItem As Variant
    For Each CurrentItem In FromCollection
        ToCollection.Add CurrentItem
    Next CurrentItem

End Sub

'@Description("This will copy down item from array to collection.")
'@Dependency("No Dependency")
'@ExampleCall : CopyFromArrayToCollection InvalidDataRowsForCurrentClient, InvalidDataRows
'@Date : 18 March 2022 01:02:07 AM
'@PossibleError :
Public Sub CopyFromArrayToCollection(FromArray As Variant, ToCollection As Collection)

    If Not IsArray(FromArray) Or ToCollection Is Nothing Then Exit Sub
    Dim CurrentItem As Variant
    For Each CurrentItem In FromArray
        ToCollection.Add CurrentItem
    Next CurrentItem

End Sub

'@Description("This will add a Named range. If Dynamic spill formula is present then it will try to guess the name from Parent cell top or left cell value. This depends on user selction")
'@Dependency("FunctionCollection.HasDynamicFormula,FunctionCollection.MakeValidDefinedName")
'@ExampleCall : AddNameRange True, Sheet3
'@Date : 23 March 2022 07:47:31 PM
'@PossibleError :
Public Sub AddNameRange(Optional IsLocal As Boolean, Optional ScopeSheet As Worksheet)

    Dim SelectionRange As Range
    Set SelectionRange = Selection
    Dim DefaultName As String
    Dim IsDynamicFormula As Boolean
    If SelectionRange.Cells.Count = 1 And HasDynamicFormula(SelectionRange) Then
        If SelectionRange.SpillParent.Offset(-1, 0).Value <> vbNullString Then
            DefaultName = SelectionRange.SpillParent.Offset(-1, 0).Value
        Else
            DefaultName = SelectionRange.SpillParent.Offset(0, -1).Value
        End If
        IsDynamicFormula = True
    Else
        IsDynamicFormula = False
    End If
    Const STRING_TYPE As Byte = 2
    DefaultName = Application.InputBox("Give Name of this Named Range", , DefaultName, Type:=STRING_TYPE)
    DefaultName = MakeValidDefinedName(DefaultName)
    Dim Reference As String
    Reference = ConvertToReference(SelectionRange, IsDynamicFormula)
    If IsLocal And Not ScopeSheet Is Nothing Then
        ScopeSheet.Names.Add Name:=DefaultName, RefersTo:=Reference
    ElseIf IsLocal Then
        SelectionRange.Parent.Names.Add Name:=DefaultName, RefersToR1C1:=Reference
    Else
        ActiveWorkbook.Names.Add Name:=DefaultName, RefersTo:=Reference
    End If

End Sub

'@Description("This will remove all the name range from the given object (Sheet or Workbook)")
'@Dependency("No Dependency")
'@ExampleCall : RemoveNameRangeFromObject CurrentSheet, RemoveNameRangeFromObject GivenWorkbook
'@Date : 23 March 2022 07:50:23 PM
'@PossibleError :
Public Sub RemoveNameRangeFromObject(GivenObject As Object)
    
    Dim CurrentName As Name
    For Each CurrentName In GivenObject.Names
        Debug.Print CurrentName.Name
        On Error Resume Next
        CurrentName.Delete
    Next CurrentName
    
End Sub

Public Sub SelectFileInExplorer(FileFullPath As String)

    If Dir(FileFullPath, vbNormal) <> vbNullString Then
        Const EXPLORER_PATH As String = "C:\Windows\explorer.exe"
        Dim CommandText As String
        CommandText = EXPLORER_PATH & Space(1) & """""" & "/select," & FileFullPath & """"""
        Shell CommandText, vbMaximizedFocus
    End If

End Sub

Public Sub OpenFolderInExplorerOrFinder(PathToOpen As String)
    
    If Dir(PathToOpen, vbDirectory) <> vbNullString Then

        #If Mac Then
            Dim ScriptPath As String
            Const LIBRARY_EXTRA_PART As String = "Library/Application Scripts/com.microsoft.Excel/"
            
            ScriptPath = GetSpecialFolderPathInMac("home folder") _
                         & Replace(LIBRARY_EXTRA_PART, "/", Application.PathSeparator) _
                         & "AppleScriptsForVBA.scpt"
                         
            If Not IsFileExist(ScriptPath) Then
                MsgBox ScriptPath & " not found. Make sure you have installed all necessary file."
                Exit Sub
            End If
            
            AppleScriptTask "AppleScriptsForVBA.scpt", "OpenFileOrFolder", PathToOpen
            
        #Else
            Const EXPLORER_PATH As String = "C:\Windows\explorer.exe"
            Shell EXPLORER_PATH & Space(1) & PathToOpen, vbMaximizedFocus
        #End If
        
    End If

End Sub

Private Function IsFileExist(ByVal FilePath As String) As Boolean


    '@Description("This will check if file is found or not. Sometimes Dir function gives false result specially for temp/appdata folder")
    '@Dependency("No Dependency")
    '@ExampleCall :
    '@Date : 05 February 2023 12:00:18 PM
    '@PossibleError:

    #If Mac Then
        IsFileExist = IsFileOrFolderExistsOnMac(FilePath, True)
    #Else
        Dim FileManager As Object
        Set FileManager = CreateObject("Scripting.FileSystemObject")
        IsFileExist = FileManager.FileExists(FilePath)
        Set FileManager = Nothing
    #End If

End Function

Private Function IsFileOrFolderExistsOnMac(ByVal FileOrFolderPath As String _
                                           , ByVal IsFile As Boolean) As Boolean

    On Error GoTo HandleError
    If IsFile Then
        IsFileOrFolderExistsOnMac = (Dir(FileOrFolderPath & "*") <> vbNullString)
    Else
        IsFileOrFolderExistsOnMac = (Dir(FileOrFolderPath & "*", vbDirectory) <> vbNullString)
    End If
    Exit Function

HandleError:
    IsFileOrFolderExistsOnMac = False

End Function

Private Function GetSpecialFolderPathInMac(ByVal NameFolder As String) As String

    '***Possible value for NameFolder param***
    'desktop folder
    'documents folder
    'downloads folder
    'favorites folder
    'home folder
    'startup disk
    'system folder
    'users folder
    'utilities folder

    Dim SpecialFolder As String
    ' Excel 2016 or higher
    If Int(Val(Application.Version)) > 14 Then
        SpecialFolder = MacScript("return POSIX path of (path to " & NameFolder & ") as string")
        'Replace line needed for the special folders Home and documents
        Const ADDED_PART_FOR_HOME_AND_DOCUMENTS As String = "/Library/Containers/com.microsoft.Excel/Data"
        SpecialFolder = Replace(SpecialFolder, ADDED_PART_FOR_HOME_AND_DOCUMENTS, vbNullString)
    Else
        'Excel 2011
        SpecialFolder = MacScript("return (path to " & NameFolder & ") as string")
    End If
    GetSpecialFolderPathInMac = SpecialFolder

End Function

Public Sub SwapTwoRowsInPlace(ByRef InputArray As Variant, FirstRowIndex As Long, SecondRowIndex As Long)
    
    Dim Temp As Variant
    Dim ColumnIndex As Long
    For ColumnIndex = LBound(InputArray, 2) To UBound(InputArray, 2)
        Temp = InputArray(FirstRowIndex, ColumnIndex)
        InputArray(FirstRowIndex, ColumnIndex) = InputArray(SecondRowIndex, ColumnIndex)
        InputArray(SecondRowIndex, ColumnIndex) = Temp
    Next ColumnIndex
    
End Sub

'@Description("This will delete all the files and nested folder recursively.")
'@Dependency("No Dependency")
'@ExampleCall :
'@Date : 07 July 2023 12:03:31 AM
'@PossibleError :
Public Sub DeleteFolder(FolderPath As String, Optional FolderManager As Object)
    
    If FolderManager Is Nothing Then Set FolderManager = CreateObject("Scripting.FileSystemObject")
    Dim CurrentFolder As Object
    Set CurrentFolder = FolderManager.GetFolder(FolderPath)
    Dim CurrentFile As Object
    
    Debug.Print "Processing Folder: " & FolderPath
    On Error GoTo LogAndTryNext
    For Each CurrentFile In CurrentFolder.Files
        Debug.Print "Deleting : " & CurrentFile.Path
        CurrentFile.Delete
    Next CurrentFile
    
    Dim SubFolder As Object
    For Each SubFolder In CurrentFolder.SubFolders
        DeleteFolder SubFolder.Path, FolderManager
    Next SubFolder
    Debug.Print "Deleting Folder: " & CurrentFolder.Path
    CurrentFolder.Delete
    Exit Sub
    
LogAndTryNext:
    Debug.Print Err.Number & "-" & Err.Description
    Resume Next
    
End Sub

Public Sub DeleteLogFilesFromFolder(FolderPath As String, Optional FolderManager As Object)
    
    If FolderManager Is Nothing Then Set FolderManager = CreateObject("Scripting.FileSystemObject")
    Dim CurrentFolder As Object
    Set CurrentFolder = FolderManager.GetFolder(FolderPath)
    Dim CurrentFile As Object
    
    Debug.Print "Processing Folder: " & FolderPath
    On Error GoTo LogAndTryNext
    For Each CurrentFile In CurrentFolder.Files
        If LCase(Right(CurrentFile.Name, 4)) = ".log" Then
            Debug.Print "Deleting : " & CurrentFile.Path
            CurrentFile.Delete
        End If
    Next CurrentFile
    
    Dim SubFolder As Object
    For Each SubFolder In CurrentFolder.SubFolders
        DeleteLogFilesFromFolder SubFolder.Path, FolderManager
    Next SubFolder
    Exit Sub
    
LogAndTryNext:
    Debug.Print Err.Number & "-" & Err.Description
    Resume Next

End Sub

Public Sub CreateHyperlink(ByVal AnchorCell As Range _
                           , Optional ByVal URL As String = vbNullString _
                            , Optional ByVal DisplayText As String = vbNullString _
                             , Optional ByVal ScreenTipText As String = vbNullString)
    
    If URL = vbNullString Then URL = AnchorCell.Value
    If DisplayText = vbNullString Then
        DisplayText = Mid(AnchorCell.Value, InStrRev(AnchorCell.Value, "/") + 1)
    End If
    
    If ScreenTipText = vbNullString Then
        ScreenTipText = "Click on this cell to open: " & URL
    End If
    
    With AnchorCell.Worksheet
        .Hyperlinks.Add Anchor:=AnchorCell, Address:=URL, ScreenTip:=ScreenTipText, TextToDisplay:=DisplayText
    End With
    
End Sub

Private Function ConcatenateVector(ByVal GivenVector As Variant _
                                   , Optional ByVal Delimiter As String = ",") As String
    
    '@Description("This will concatenate a vector or 1D array item to a string")
    '@Dependency("No Dependency")
    '@ExampleCall : ConcatenateVector(a) where a=array(1,2,3,4)
    '@Date : 26 December 2021 01:31:45 AM

    Dim CurrentIndex As Long
    Dim OutputText As String
    For CurrentIndex = LBound(GivenVector) To UBound(GivenVector)
        OutputText = OutputText & Delimiter & GivenVector(CurrentIndex)
    Next CurrentIndex
    ConcatenateVector = Right$(OutputText, Len(OutputText) - Len(Delimiter))

End Function

Public Sub AutoFitRangeCols(ByVal DataRange As Range, Optional ByVal MaxColWidth As Long = 255)
    
    DataRange.Columns.AutoFit
    Dim ColIndex As Long
    For ColIndex = 1 To DataRange.Columns.Count
        If DataRange.Columns(ColIndex).ColumnWidth > MaxColWidth Then
            DataRange.Columns(ColIndex).ColumnWidth = MaxColWidth
        End If
    Next ColIndex

End Sub

Public Sub ProtectAllSheet(ByVal Password As String, Optional ByVal Book As Workbook)
    
    If Book Is Nothing Then Book = ActiveWorkbook
    
    Dim OldActiveSheet As Worksheet
    Set OldActiveSheet = Book.ActiveSheet
    
    Dim CurrentSheet As Worksheet
    For Each CurrentSheet In Book.Worksheets
        CurrentSheet.Protect Password
    Next CurrentSheet
    
    Book.Protect Password
    
    OldActiveSheet.Activate
    
End Sub

Public Sub UnprotectAllSheet(ByVal Password As String, Optional ByVal Book As Workbook)
    
    If Book Is Nothing Then Book = ActiveWorkbook
    
    Dim OldActiveSheet As Worksheet
    Set OldActiveSheet = Book.ActiveSheet
    
    Dim CurrentSheet As Worksheet
    For Each CurrentSheet In Book.Worksheets
        CurrentSheet.Unprotect Password
    Next CurrentSheet
    
    Book.Unprotect Password
    
    OldActiveSheet.Activate
    
End Sub

Private Function IsSheetExist(ByVal SheetTabName As String _
                              , Optional ByVal GivenWorkbook As Workbook) As Boolean

    '@Description("This function will determine if a sheet is exist or not by using tab name")
    '@Dependency("No Dependency")
    '@ExampleCall : IsSheetExist("SheetTabName")
    '@Date : 14 October 2021 07:03:05 PM

    If GivenWorkbook Is Nothing Then Set GivenWorkbook = ThisWorkbook

    Dim TemporarySheet As Worksheet
    On Error Resume Next
    Set TemporarySheet = GivenWorkbook.Worksheets(SheetTabName)

    IsSheetExist = (Not TemporarySheet Is Nothing)
    On Error GoTo 0

End Function

Private Function CopySheet(ByVal SourceSheet As Worksheet _
                           , Optional ByRef DestinationWorkbook As Workbook) As Worksheet

    '@Description("It will copy the SourceSheet and paste at the end of the sheets to the DestinationWorkbook and return the reference")
    '@Dependency("No Dependency")
    '@ExampleCall : CopySheet(Sheet1,Workbooks("WorkBookName"))
    '@Date : 14 October 2021 07:00:47 PM
    'More Info : https://stackoverflow.com/questions/7692274/copy-sheet-and-get-resulting-sheet-object#comment105982030_37704412

    Dim NewSheet As Worksheet
    Dim LastSheet As Worksheet
    Dim LastSheetVisibility As XlSheetVisibility

    If DestinationWorkbook Is Nothing Then
        Set DestinationWorkbook = SourceSheet.Parent
    End If

    With DestinationWorkbook
        Set LastSheet = .Worksheets(.Worksheets.Count)
    End With

    ' store visibility of last sheet
    LastSheetVisibility = LastSheet.Visible
    ' make the last sheet visible
    LastSheet.Visible = xlSheetVisible

    SourceSheet.Copy After:=LastSheet
    Set NewSheet = LastSheet.Next

    ' restore visibility of last sheet
    LastSheet.Visible = LastSheetVisibility

    Set CopySheet = NewSheet

End Function

Private Function HasDynamicFormula(ByVal SelectionRange As Range) As Boolean

    '@Description("This will check if a range has a Dynamic formula or not. If your selection cross dynamic formula section then it will return false. It has to be either a single cell or part of the dynamic formula output range")
    '@Dependency("No Dependency")
    '@ExampleCall : HasDynamicFormula(Sheet1.Range("P5:P10"))=True, HasDynamicFormula(Sheet1.Range("P5:P11"))=False as we have formula upto P10
    '@Date : 23 March 2022 07:41:19 PM
    '@PossibleError:

    On Error Resume Next
    HasDynamicFormula = SelectionRange.HasSpill
    On Error GoTo 0

End Function

Private Function MakeValidDefinedName(ByVal GivenDefinedName As String) As String


    '@Description("This will create a valid name for name range")
    '@Dependency("No Dependency")
    '@ExampleCall :MakeValidDefinedName("1ABC1") = "_1ABC1",MakeValidDefinedName("AB C1") = "_ABC1",MakeValidDefinedName(vbNullString) = "_DefaultName"
    '@Date : 23 March 2022 07:36:46 PM
    '@PossibleError:

    If Trim$(GivenDefinedName) = vbNullString Then
        MakeValidDefinedName = "_DefaultName"
        Exit Function
    End If
    If Not (Left$(GivenDefinedName, 1) = "_" Or Left$(GivenDefinedName, 1) Like "[A-Za-z]") Then
        GivenDefinedName = "_" & GivenDefinedName
    End If
    MakeValidDefinedName = Replace(GivenDefinedName, Space(1), vbNullString)
    'Handle Name like(AB25) cell reference.It will add one more underscore if given name is name range
    On Error GoTo Done:
    If Not Range(MakeValidDefinedName) Is Nothing Then
        MakeValidDefinedName = "_" & MakeValidDefinedName
    End If
Done:

End Function

Private Function ConvertToReference(ByVal DataSource As Range _
                                    , Optional ByVal IsDynamicFormula As Boolean) As String


    '@Description("This will convert range reference to a string which can be used as ReferTo for defining named range")
    '@Dependency("No Dependency")
    '@ExampleCall : ConvertToReference(SelectionRange, IsDynamicFormula)
    '@Date : 23 March 2022 07:39:32 PM
    '@PossibleError:

    Dim SheetNamePrefix As String
    SheetNamePrefix = "'" & DataSource.Parent.Name & "'!"
    If IsDynamicFormula Then
        ConvertToReference = "=" & SheetNamePrefix & DataSource.SpillParent.Address & "#"
    Else
        ConvertToReference = "=" & SheetNamePrefix & Replace(DataSource.Address, ",", "," & SheetNamePrefix)
    End If

End Function

Public Sub CopyFolder(ByVal SourceFolder As String, ByVal DestinationFolder As String)
    
    Dim FolderManager As Object
    Set FolderManager = CreateObject("Scripting.FileSystemObject")
    
    Dim SourceFolderObject As Object
    Set SourceFolderObject = FolderManager.GetFolder(SourceFolder)
    
    Dim DestinationFolderObject As Object
    Set DestinationFolderObject = FolderManager.Createfolder(DestinationFolder & Application.PathSeparator & SourceFolderObject.Name)
    
    Dim CurrentFile As Object
    For Each CurrentFile In SourceFolderObject.Files
        CurrentFile.Copy DestinationFolderObject.Path & Application.PathSeparator & CurrentFile.Name
    Next CurrentFile
    
    Dim SubFolder As Object
    For Each SubFolder In SourceFolderObject.SubFolders
        CopyFolder SubFolder.Path, DestinationFolderObject.Path & Application.PathSeparator & SubFolder.Name
    Next SubFolder

End Sub

Public Sub ExportSheetAsCSV(ByVal CurrentSheet As Worksheet _
                            , ByVal FolderPath As String _
                            , ByVal FileName As String _
                            , Optional ByVal IsReplace As Boolean = False)
    
    Dim FullFilePath As String

    If Right(FolderPath, Len(Application.PathSeparator)) = Application.PathSeparator Then
        FullFilePath = FolderPath
    Else
        FullFilePath = FolderPath & Application.PathSeparator
    End If

    If Right(FileName, 4) = ".csv" Then
        FullFilePath = FullFilePath & FileName
    Else
        FullFilePath = FullFilePath & FileName & ".csv"
    End If

    If IsReplace And IsFileExist(FullFilePath) Then
        Kill FullFilePath
    ElseIf Not IsReplace And IsFileExist(FullFilePath) Then
        Exit Sub
    End If
    
    On Error Resume Next
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Dim NewWorkbook As Workbook
    Set NewWorkbook = Application.Workbooks.Add
    CopySheet CurrentSheet, NewWorkbook
    NewWorkbook.Worksheets(1).Delete
    NewWorkbook.SaveAs FullFilePath, xlCSV
    NewWorkbook.Close
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    On Error GoTo 0
    
End Sub

Public Sub RefreshAndCreateHyperlinkInAColumn(ByVal Table As ListObject _
                                              , ByVal URLColName As String _
                                               , ByVal CreateLinkOnColName As String _
                                                , Optional ByVal RemoveURLCol As Boolean = True)

    Table.QueryTable.Refresh False

    Dim URLColIndex As Long
    URLColIndex = Table.ListColumns(URLColName).Index
    Dim CreateLinkColIndex As Long
    CreateLinkColIndex = Table.ListColumns(CreateLinkOnColName).Index

    Dim RowIndex As Long
    For RowIndex = 1 To Table.ListRows.Count
        With Table
            CreateHyperlink .DataBodyRange(RowIndex, CreateLinkColIndex) _
                            , .DataBodyRange(RowIndex, URLColIndex).Value _
                             , .DataBodyRange(RowIndex, CreateLinkColIndex).Value
        End With
    Next RowIndex

    If RemoveURLCol Then
        Table.ListColumns(URLColName).Delete
    End If

End Sub

Public Sub MoveFromOneTableToAnother(ByVal FromTable As ListObject _
                                     , ByVal ToTable As ListObject _
                                      , ByVal StatusColHeader As String _
                                       , ByVal MoveOnStatus As String _
                                        , ByRef ColMapping() As String)

    Dim StatusColIndex As Long
    StatusColIndex = FromTable.ListColumns(StatusColHeader).Index

    Dim ColMappingIndex() As Long
    ReDim ColMappingIndex(LBound(ColMapping, 1) To UBound(ColMapping, 1), 1 To 2)

    Dim FirstColumnIndex As Long
    FirstColumnIndex = LBound(ColMapping, 2)

    Dim RowIndex As Long
    For RowIndex = LBound(ColMapping, 1) To UBound(ColMapping, 1)
        ColMappingIndex(RowIndex, 1) = FromTable.ListColumns(ColMapping(RowIndex, FirstColumnIndex)).Index
        ColMappingIndex(RowIndex, 2) = ToTable.ListColumns(ColMapping(RowIndex, FirstColumnIndex + 1)).Index
    Next RowIndex

    Dim FromTableData As Variant
    FromTableData = FromTable.DataBodyRange.Value

    Dim ValidRowIndexes As Collection
    Set ValidRowIndexes = New Collection

    For RowIndex = LBound(FromTableData, 1) To UBound(FromTableData, 1)
        If FromTableData(RowIndex, StatusColIndex) = MoveOnStatus Then
            ValidRowIndexes.Add RowIndex
        End If
    Next RowIndex
    
    If ValidRowIndexes.Count = 0 Then Exit Sub

    Dim LeftBottomCell As Range
    If ToTable.ListRows.Count = 0 Then
        Set LeftBottomCell = ToTable.Range(1, 1)
    Else
    
        Set LeftBottomCell = ToTable.DataBodyRange(ToTable.ListRows.Count, 1)
        If IsBlankRange(ToTable.ListRows(ToTable.ListRows.Count).Range) Then
            Set LeftBottomCell = LeftBottomCell.Offset(-1)
        End If
        
    End If
    
    ' Add to ToTable
    Dim Counter As Long
    Dim CurrentColIndex As Long
    For CurrentColIndex = LBound(ColMapping, 1) To UBound(ColMapping, 1)
        
        Dim Temp As Variant
        ReDim Temp(1 To ValidRowIndexes.Count, 1 To 1)
        Counter = 0
        Dim ValidRowIndex As Variant
        For Each ValidRowIndex In ValidRowIndexes
            Counter = Counter + 1
            Temp(Counter, 1) = FromTableData(ValidRowIndex, ColMappingIndex(CurrentColIndex, 1))
        Next ValidRowIndex
        
        LeftBottomCell.Offset(1, ColMappingIndex(CurrentColIndex, 2) - 1).Resize(ValidRowIndexes.Count, 1).Value = Temp
        
    Next CurrentColIndex
    
    ' Delete from FromTable
    Dim CurrentItemIndex As Long
    For CurrentItemIndex = ValidRowIndexes.Count To 1 Step -1
        FromTable.ListRows(ValidRowIndexes.Item(CurrentItemIndex)).Delete
    Next CurrentItemIndex

End Sub

Private Function IsBlankRange(ByVal CheckRange As Range) As Boolean

    On Error Resume Next
    Dim FormulaCells As Range
    IsBlankRange = True
    If CheckRange.Cells.Count = 1 Then
        If CheckRange.HasFormula Then
            IsBlankRange = False
        ElseIf CheckRange.Value <> vbNullString Then
            IsBlankRange = False
        End If
    Else

        Set FormulaCells = CheckRange.SpecialCells(xlCellTypeFormulas)
        If FormulaCells Is Nothing Then
            Dim Values As Variant
            Values = CheckRange.Value
            Dim Element As Variant
            For Each Element In Values
                If Element <> vbNullString Then
                    IsBlankRange = False
                    Exit For
                End If
            Next Element

        Else
            IsBlankRange = False
        End If

    End If

    On Error GoTo 0

End Function

Public Sub PrintChartNameFromSheet(ByVal SourceSheet As Worksheet)
    
    Dim CurrentChartObj As ChartObject
    For Each CurrentChartObj In SourceSheet.ChartObjects
        Debug.Print CurrentChartObj.Name
    Next CurrentChartObj
    
End Sub

Public Sub DeleteOldRecordsFromTable(ByVal Table As ListObject)
    
    ' This will delete old records from table and keep the first row formulas.
    ' But it will remove const cells content from the first row as well.
    If Table Is Nothing Then Exit Sub
    If Table.ListRows.Count = 0 Then Exit Sub
    
    Application.DisplayAlerts = False
    Dim RowIndex As Long
    For RowIndex = Table.ListRows.Count To 2 Step -1
        Table.ListRows(RowIndex).Delete
    Next RowIndex
    
    Application.DisplayAlerts = True
    
    ClearNonFormulaCellsOfRow Table, 1
    
    
End Sub

Private Sub ClearNonFormulaCellsOfRow(ByVal Table As ListObject, ByVal RowIndex As Long)
    
    Dim ColIndex As Long
    For ColIndex = 1 To Table.ListColumns.Count
        
        Dim FirstRowCell As Range
        Set FirstRowCell = Table.DataBodyRange(1, ColIndex)
        
        If Not FirstRowCell.HasFormula Then
            FirstRowCell.ClearContents
        End If
        
    Next ColIndex
    
End Sub
