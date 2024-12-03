Attribute VB_Name = "FXCollection"
    Dim InfoColl As Collection
'Developer: Md.Ismail Hosen
'Please contact for any project or VBA Automation.
'Email : 1997ismail.hosen@gmail.com
'Whatsapp: +8801515649307
'LinkedIn : https://www.linkedin.com/in/md-ismail-hosen-b77500135/
'Facebook : https://www.facebook.com/mdismail.hosen.7
'Youtube : https://www.youtube.com/channel/UCL-q7_WvISkw0Ox9FRBBzmw

'External Dependency: LateBound is being used here.
'1. Microsoft Scripting Runtime
'2.Microsoft ActiveX Data Objects 6.1 Library Or other version(6.0,2.8,2.6 etc)
'3.

'@Folder("Reusable.Function")
'This module contains all the public function which can be used in multiple project

Option Explicit

Public Const QUOTATION_SIGN As String = """"
Private Const SHELL_FOLDER_KEY As String = "SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders"

Private Const ERR_UN_ALLOCATED_ARRAY As Long = vbObjectError + 1
Private Const MSG_UN_ALLOCATED_ARRAY As String = "Un allocated array."

'This is for the Array dimension changing function.
Public Enum ProcessDirection
    LeftToRightThenBottom = 1
    TopToBottomThenRight = 3
End Enum

'This is for the Array dimension changing function
Public Enum ChangedTo
    FixedRow = 1
    FixedColumn = 2
End Enum

' This is for reading from registry
Public Enum BASE_KEY
    HKEY_CURRENT_USER = &H80000001
    HKEY_LOCAL_MACHINE = &H80000002
    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_CONFIG = &H80000005
    HKEY_DYN_DATA = &H80000006
    HKEY_PERFORMANCE_DATA = &H80000004
    HKEY_USERS = &H80000003
End Enum

Public Enum FolderItemFilter
    VISIBLE_ONLY
    HIDDEN_ONLY
    VISIBLE_AND_HIDDEN
End Enum

Public Enum InputBoxEnum
    IBFormula = 0
    IBNumber = 1
    IBString = 2
    IBBoolean = 4
    IBRange = 8
    IBError = 16
    IBArray = 64
End Enum


#If VBA7 Then

    Private Declare PtrSafe Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" ( _
                                     ByVal hKey As LongPtr, _
                                     ByVal lpSubKey As String, _
                                     ByVal ulOptions As Long, _
                                     ByVal samDesired As Long, _
                                     phkResult As Long) As Long

    Private Declare PtrSafe Function RegQueryValueExStr Lib "advapi32" Alias "RegQueryValueExA" ( _
                                     ByVal hKey As LongPtr, _
                                     ByVal lpValueName As String, _
                                     ByVal lpReserved As Long, _
                                     ByRef lpType As Long, _
                                     ByVal szData As String, _
                                     ByRef lpcbData As Long) As Long

    Private Declare PtrSafe Function RegCloseKey Lib "advapi32.dll" ( _
                                     ByVal hKey As LongPtr) As Long

#Else

    Private Declare  Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" ( _
    ByVal HKey As Long, _
    ByVal lpSubKey As String, _
    ByVal ulOptions As Long, _
    ByVal samDesired As Long, _
    phkResult As Long) As Long

    Private Declare  Function RegQueryValueExStr Lib "advapi32" Alias "RegQueryValueExA" ( _
    ByVal HKey As Long, _
    ByVal lpValueName As String, _
    ByVal lpReserved As Long, _
    ByRef lpType As Long, _
    ByVal szData As String, _
    ByRef lpcbData As Long) As Long

    Private Declare  Function RegCloseKey Lib "advapi32.dll" ( _
    ByVal HKey As Long) As Long

#End If

Public Function LastUsedRowNumber(Optional ByVal GivenWorksheet As Worksheet = Nothing _
                                  , Optional ByVal GivenColumn As Variant = 1) As Long

    'This will find the lastrow based on given worksheet and given column..Both are optional
    'If column is not givent then it will use first column
    'If worksheet is not given then it will use activesheet
    'It will use xlup to find lastrow..

    'if worksheet is not given then use activesheet.
    If GivenWorksheet Is Nothing Then
        Set GivenWorksheet = ActiveSheet
    End If
    
    With GivenWorksheet
        'Find lastrow
        LastUsedRowNumber = .Cells(.Rows.Count, GivenColumn).End(xlUp).Row
    End With

End Function

Public Function LastUsedColumnNumber(Optional ByVal GivenWorksheet As Worksheet = Nothing _
                                     , Optional ByVal GivenRow As Variant = 1) As Long

    'This will find the lastColumn based on given worksheet and given row..Both are optional
    'If row is not givent then it will use first row
    'If worksheet is not given then it will use activesheet
    'It will use xlToLeft to find last row..

    'if worksheet is not given then use activesheet.
    If GivenWorksheet Is Nothing Then
        Set GivenWorksheet = ActiveSheet
    End If
    
    With GivenWorksheet
        'Find lastColumn
        LastUsedColumnNumber = .Cells(GivenRow, .Columns.Count).End(xlToLeft).Column
    End With

End Function

Public Function GetSelectedFilePath(ByVal GivenTitle As String _
                                    , Optional ByVal GivenFilter As String = "*.*") As String

    'This will give the selected file path as string.
    'Example call : GetSelectedFilePath("Select Correct CSV","*.csv")

    #If Mac Then
        GetSelectedFilePath = GetSelectedFileOrFilesOnMac(GivenTitle, True, GivenFilter)
    #Else

        Dim SelectedFilePath As FileDialogSelectedItems
        Set SelectedFilePath = GetSelectedFilesPathOnWindows(GivenTitle, GivenFilter, False)
        If SelectedFilePath.Count = 0 Then
            GetSelectedFilePath = vbNullString
        Else
            GetSelectedFilePath = SelectedFilePath.Item(1)
        End If

    #End If

End Function

Public Function GetSelectedFilesPath(ByVal GivenTitle As String _
                                     , Optional ByVal GivenFilter As String = "*.*") As Variant

    #If Mac Then
        GetSelectedFilesPath = GetSelectedFileOrFilesOnMac(GivenTitle, False, GivenFilter)
        Exit Function
    #End If

    Dim SelectedFilePath As FileDialogSelectedItems
    Set SelectedFilePath = GetSelectedFilesPathOnWindows(GivenTitle, GivenFilter, True)
    If SelectedFilePath.Count > 0 Then
        Dim Result As Variant
        ReDim Result(0 To SelectedFilePath.Count - 1)
        Dim Counter As Long
        For Counter = 1 To SelectedFilePath.Count
            Result(Counter - 1) = SelectedFilePath.Item(Counter)
        Next Counter
        GetSelectedFilesPath = Result
    End If

End Function

Private Function GetSelectedFileOrFilesOnMac(ByVal Prompt As String _
                                             , ByVal IsOneFile As Boolean _
                                              , Optional ByVal GivenFilter As String = "*.*") As String

    Dim MyScript As String
    Dim MyFiles As String
    
    Dim FileFilterType As String
    FileFilterType = GetFileFilterForMac(GivenFilter)
    
    Dim FileFilterPartScript As String
    If FileFilterType = vbNullString Then
        FileFilterPartScript = vbNullString
    Else
        FileFilterPartScript = "of type " & FileFilterType & " "
    End If
    
    On Error GoTo HandleError
    If IsOneFile Then
        MyScript = "set theFile to (choose file " & FileFilterPartScript & "with prompt """ & Prompt & """" _
                   & " without multiple selections allowed) as string" & vbNewLine & _
                   "return posix path of theFile"
    Else
        MyScript = _
                 "set theFiles to (choose file " & FileFilterType & "with prompt """ & Prompt & """" _
                 & " with multiple selections allowed)" & vbNewLine & _
                 "set thePOSIXFiles to {}" & vbNewLine & _
                 "repeat with aFile in theFiles" & vbNewLine & _
                 "set end of thePOSIXFiles to POSIX path of aFile" & vbNewLine & _
                 "end repeat" & vbNewLine & _
                 "set {TID, text item delimiters} to {text item delimiters, ASCII character 10}" & vbNewLine & _
                 "set thePOSIXFiles to thePOSIXFiles as text" & vbNewLine & _
                 "set text item delimiters to TID" & vbNewLine & _
                 "return thePOSIXFiles"
    End If

    MyFiles = MacScript(MyScript)
    If IsOneFile Then
        GetSelectedFileOrFilesOnMac = MyFiles
    Else
        GetSelectedFileOrFilesOnMac = Split(MyFiles, Chr$(10))
    End If
    Exit Function

HandleError:
    GetSelectedFileOrFilesOnMac = vbNullString

End Function

Private Function GetFileFilterForMac(Optional ByVal GivenFilter As String = "*.*") As String
    
    Dim ExtensionVsFilterMap As Collection
    Set ExtensionVsFilterMap = New Collection
    
    With ExtensionVsFilterMap
        .Add "com.microsoft.Excel.xls", "*.xls"
        .Add "org.openxmlformats.spreadsheetml.sheet", "*.xlsx"
        .Add "org.openxmlformats.spreadsheetml.sheet.macroenabled", "*.xlsm"
        .Add "com.microsoft.Excel.sheet.binary.macroenabled", "*.xlsb"
        .Add "public.comma-separated-values-text", "*.csv"
        .Add "com.microsoft.word.doc", "*.doc"
        .Add "org.openxmlformats.wordprocessingml.document", "*.docx"
        .Add "org.openxmlformats.wordprocessingml.document.macroenabled", "*.docm"
        .Add "com.microsoft.powerpoint.ppt", "*.ppt"
        .Add "org.openxmlformats.presentationml.presentation", "*.pptx"
        .Add "org.openxmlformats.presentationml.presentation.macroenabled", "*.pptm"
        .Add "public.plain-text", "*.txt"
        .Add "com.adobe.pdf", "*.pdf"
        .Add "public.jpeg", "*.jpg"
        .Add "public.png", "*.png"
        .Add "com.apple.traditional-mac-plain-text", "*.QIF"
        .Add "public.html", "*.htm"
    End With
    
    Dim FileExtensions As Variant
    If IsTextPresent(GivenFilter, ",") Then
        FileExtensions = VBA.Split(GivenFilter, ",")
    ElseIf IsTextPresent(GivenFilter, ";") Then
        FileExtensions = VBA.Split(GivenFilter, ";")
    Else
        FileExtensions = Array(GivenFilter)
    End If
    
    Dim IsAnyInvalidExtension As Boolean
    
    Dim Result As String
    Result = "{"
    
    Dim CurrentExtension As Variant
    For Each CurrentExtension In FileExtensions
        
        If Not IsExistInCollection(ExtensionVsFilterMap, CStr(CurrentExtension)) Then
            IsAnyInvalidExtension = True
            Exit For
        Else
            Result = Result & """" & _
                     ExtensionVsFilterMap.Item(CurrentExtension) _
                     & """" & ","
        End If
        
    Next CurrentExtension
    
    If IsAnyInvalidExtension Then
        Result = vbNullString
    Else
        Result = Left(Result, Len(Result) - 1) & "}"
    End If
    
    GetFileFilterForMac = Result
    
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

Private Function GetSelectedFilesPathOnWindows(ByVal GivenTitle As String _
                                               , ByVal GivenFilter As String _
                                                , Optional ByVal IsMultiSelected As Boolean = False) As FileDialogSelectedItems
    'This will give the selected file path as string.
    'Example call : GetSelectedFilePath("Select Correct CSV","*.csv",True)

    Dim FilePicker As FileDialog
    Set FilePicker = Application.FileDialog(msoFileDialogFilePicker)

    With FilePicker
        .AllowMultiSelect = IsMultiSelected
        .Title = GivenTitle
        .InitialFileName = CurDir$()
        .Filters.Clear
        .Filters.Add "Filter : ", GivenFilter
        .Show
        Set GetSelectedFilesPathOnWindows = .SelectedItems
    End With

End Function

Public Function GetAllSelectedFilePath(ByVal GivenTitle As String _
                                       , ByVal GivenFilter As String) As Variant

    'This will give all the selected file path as string.
    'Example call : GetAllSelectedFilePath("Select Correct CSV","*.csv;*.txt")

    Const FILE_SELECTION_SUCCESSFUL As Long = -1
    Dim FilePicker As FileDialog
    Set FilePicker = Application.FileDialog(msoFileDialogFilePicker)
    Dim Result As Variant
    With FilePicker
        .AllowMultiSelect = True
        .Title = GivenTitle
        .InitialFileName = CurDir$()
        .Filters.Clear
        .Filters.Add "Filter : ", GivenFilter
        If .Show = FILE_SELECTION_SUCCESSFUL Then
            ReDim Result(1 To .SelectedItems.Count, 1 To 1)
            Dim Counter As Long
            For Counter = 1 To .SelectedItems.Count
                Result(Counter, 1) = .SelectedItems(Counter)
            Next Counter
        End If

    End With
    GetAllSelectedFilePath = Result

End Function

Public Function GetSelectedFolderPath(ByVal GivenTitle As String) As String

    'This will give the selected folder path as string.
    'Example call : GetSelectedFolderPath("Select Correct Folder")

    #If Mac Then
        GetSelectedFolderPath = GetSelectFolderOnMac(GivenTitle)
        Exit Function
    #End If

    Const FOLDER_SELECTION_SUCCESSFUL As Long = -1
    Dim FolderPicker As FileDialog
    Set FolderPicker = Application.FileDialog(msoFileDialogFolderPicker)

    With FolderPicker
        .AllowMultiSelect = False
        .Title = GivenTitle
        .InitialFileName = CurDir$()

        If .Show = FOLDER_SELECTION_SUCCESSFUL Then
            GetSelectedFolderPath = .SelectedItems(1)
        End If

    End With

End Function

Public Function GetSelectedFolderPathOnMac(ByVal Prompt As String) As String

    Dim Script As String
    Script = "set folderPath to POSIX path of (choose folder with prompt """ & Prompt & """)"
    On Error GoTo UserCancelledFolderSelection
    Dim Path As String
    Path = MacScript(Script)

    If Path = vbNullString Then
        Path = Left$(Path, Len(Path) - Len(Application.PathSeparator))
    End If

    GetSelectedFolderPathOnMac = Path
    Exit Function

UserCancelledFolderSelection:
    GetSelectedFolderPathOnMac = vbNullString

End Function

Public Function AddPathSeperatorIfNotExists(ByVal FolderPath As String) As String

    Dim Path As String
    Path = FolderPath
    If Right$(Path, Len(Application.PathSeparator)) <> Application.PathSeparator Then
        Path = Path & Application.PathSeparator
    End If

    AddPathSeperatorIfNotExists = Path

End Function

Public Function GetQueryList(ByVal GivenWorkbook As Workbook) As Variant

    'This function is for getting list of all the WorkbookConnection in the given workbook.
    'Example call : GetQueryList(Thisworkbook) or getQueryList(workbooks("Name"))
    'Dependency : WorkbookConnectionTypeCollection function

    With GivenWorkbook

        Dim ConnectionType As Collection
        Set ConnectionType = WorkbookConnectionTypeCollection

        Dim ConnectionList() As Variant
        ReDim ConnectionList(1 To .Connections.Count, 1 To 2)
        Dim CurrentQuery As WorkbookConnection
        Dim Counter As Long
        For Each CurrentQuery In .Connections
            Counter = Counter + 1
            ConnectionList(Counter, 1) = CurrentQuery.Name
            ConnectionList(Counter, 2) = ConnectionType.Item(CurrentQuery.Type)
        Next CurrentQuery

    End With

    GetQueryList = ConnectionList

End Function

Public Function GetWorkBookQueryList(ByVal GivenWorkbook As Workbook) As Variant

    '@Description("This will extract the queries  name from the givenworkbook")
    '@Dependency("No Dependency")
    '@ExampleCall : GetWorkBookQueryList(ThisWorkbook)
    '@Date : 06 March 2022 12:42:52 AM
    '@PossibleError:

    Dim ConnectionList() As Variant
    If GivenWorkbook.Queries.Count = 0 Then Exit Function
    ReDim ConnectionList(1 To GivenWorkbook.Queries.Count, 1 To 1)
    Dim CurrentQuery As WorkbookQuery
    Dim Counter As Long
    For Each CurrentQuery In GivenWorkbook.Queries
        Counter = Counter + 1
        ConnectionList(Counter, 1) = CurrentQuery.Name
    Next CurrentQuery
    GetWorkBookQueryList = ConnectionList

End Function

Private Function WorkbookConnectionTypeCollection() As Collection

    'This is just a collection of XlConnectionType enumeration
    ' MSDN : https://docs.microsoft.com/en-us/office/vba/api/excel.xlconnectiontype
    'This will return enum number as key and description as item so that we can get description using key
    'Dependency : No Dependency

    Dim AllConnectionType As Collection
    Set AllConnectionType = New Collection
    AllConnectionType.Add Item:="OLEDB", Key:=CStr(1)
    AllConnectionType.Add Item:="ODBC", Key:=CStr(2)
    AllConnectionType.Add Item:="XML MAP", Key:=CStr(3)
    AllConnectionType.Add Item:="Text", Key:=CStr(4)
    AllConnectionType.Add Item:="Web", Key:=CStr(5)
    AllConnectionType.Add Item:="Data Feed", Key:=CStr(6)
    AllConnectionType.Add Item:="PowerPivot Model", Key:=CStr(7)
    AllConnectionType.Add Item:="Worksheet", Key:=CStr(8)
    AllConnectionType.Add Item:="No source", Key:=CStr(9)

    Set WorkbookConnectionTypeCollection = AllConnectionType
    Set AllConnectionType = Nothing

End Function

Public Function GetOpenWorkbookNameList() As String()

    'Get all the open workbooks name

    Dim TotalOpenWorkbook As Long
    TotalOpenWorkbook = Application.Workbooks.Count
    Dim CurrentWorkbook As Workbook
    Dim AllWorkbook() As String
    ReDim AllWorkbook(1 To TotalOpenWorkbook, 1 To 1)
    Dim Counter As Long
    For Each CurrentWorkbook In Application.Workbooks
        Counter = Counter + 1
        AllWorkbook(Counter, 1) = CurrentWorkbook.Name
    Next CurrentWorkbook
    GetOpenWorkbookNameList = AllWorkbook

End Function

Public Function GetOpenWorkbookList() As Collection

    'Get all the open workbooks

    Dim AllWorkbook As Collection
    Set AllWorkbook = New Collection
    Dim CurrentWorkbook As Workbook
    For Each CurrentWorkbook In Application.Workbooks
        AllWorkbook.Add CurrentWorkbook, CurrentWorkbook.Name
    Next CurrentWorkbook

    Set GetOpenWorkbookList = AllWorkbook

End Function

Public Function GetOpenWorkbookIncludingAddInList() As Collection

    'Get all the open workbooks

    Dim AllWorkbook As Collection
    Set AllWorkbook = New Collection
    Dim CurrentWorkbook As Workbook

    ' Loop through all the workbooks
    For Each CurrentWorkbook In Application.Workbooks
        AllWorkbook.Add CurrentWorkbook, CurrentWorkbook.Name
    Next CurrentWorkbook

    ' Loop through all the add-ins
    Dim CurrentAddIn As AddIn
    For Each CurrentAddIn In Application.AddIns2
        Dim IsXlaAddIn As Boolean
        IsXlaAddIn = CurrentAddIn.Name Like "*.xla*"
        If IsXlaAddIn And CurrentAddIn.IsOpen And CurrentAddIn.Installed Then
            AllWorkbook.Add Application.Workbooks(CurrentAddIn.Name), CurrentAddIn.Name
        End If
    Next CurrentAddIn

    Set GetOpenWorkbookIncludingAddInList = AllWorkbook

End Function

Public Function GetOpenXLAAddInList() As Collection

    Dim AllWorkbook As Collection
    Set AllWorkbook = New Collection

    ' Loop through all the add-ins
    Dim CurrentAddIn As AddIn
    For Each CurrentAddIn In Application.AddIns2
        Dim IsXlaAddIn As Boolean
        IsXlaAddIn = CurrentAddIn.Name Like "*.xla*"
        If IsXlaAddIn And CurrentAddIn.IsOpen And CurrentAddIn.Installed Then
            AllWorkbook.Add Application.Workbooks(CurrentAddIn.Name), CurrentAddIn.Name
        End If
    Next CurrentAddIn

    Set GetOpenXLAAddInList = AllWorkbook

End Function

Public Function GetTableNameAsConstDeclaration(Optional ByVal GivenWorkbook As Workbook) As String

    '@Description("This will return all tables name as const declaration")
    '@Dependency("MakeValidConstName function")
    '@ExampleCall : GetTableNameAsConstDeclaration(ThisWorkbook)
    '@Date : 14 October 2021 10:20:10 PM

    If GivenWorkbook Is Nothing Then Set GivenWorkbook = ActiveWorkbook
    Dim Output As String
    Dim CurrentTable As ListObject
    Dim CurrentSheet As Worksheet
    Dim ValidConstName As String
    For Each CurrentSheet In GivenWorkbook.Worksheets
        For Each CurrentTable In CurrentSheet.ListObjects
            ValidConstName = MakeValidConstName(CurrentTable.Name)
            Output = Output & "Public Const " & ValidConstName _
                     & " As String =""" & CurrentTable.Name & """" & vbNewLine
        Next CurrentTable
    Next CurrentSheet
    If Output <> vbNullString Then Output = Left$(Output, Len(Output) - Len(vbNewLine))
    GetTableNameAsConstDeclaration = Output

End Function

Public Function GetNamedRangeNameAsConstDeclaration(Optional ByVal GivenWorkbook As Workbook) As String

    If GivenWorkbook Is Nothing Then Set GivenWorkbook = ActiveWorkbook
    Dim Output As String
    Dim CurrentName As Name
    Dim CurrentSheet As Worksheet
    Dim ValidConstName As String

    For Each CurrentName In GivenWorkbook.Names
        If CurrentName.Visible Then
            ValidConstName = MakeValidConstName(CurrentName.Name)
            Output = Output & "Public Const " & ValidConstName _
                     & " As String =""" & CurrentName.Name & """" & vbNewLine
        End If
    Next CurrentName

    For Each CurrentSheet In GivenWorkbook.Worksheets
        For Each CurrentName In CurrentSheet.Names
            If CurrentName.Visible Then
                ValidConstName = MakeValidConstName(CurrentName.Name)
                Output = Output & "Public Const " & ValidConstName _
                         & " As String =""" & CurrentName.Name & """" & vbNewLine
            End If
        Next CurrentName
    Next CurrentSheet
    If Output <> vbNullString Then Output = Left$(Output, Len(Output) - Len(vbNewLine))
    GetNamedRangeNameAsConstDeclaration = Output

End Function

Public Function GetPivotTableNameAsConstDeclaration(Optional ByVal GivenWorkbook As Workbook) As String

    '@Description("This will return all tables name as const declaration")
    '@Dependency("MakeValidConstName function")
    '@ExampleCall : GetPivotTableNameAsConstDeclaration(ThisWorkbook)
    '@Date : 14 October 2021 10:20:10 PM

    If GivenWorkbook Is Nothing Then Set GivenWorkbook = ActiveWorkbook
    Dim Output As String
    Dim CurrentPivotTable As PivotTable
    Dim CurrentSheet As Worksheet
    Dim ValidConstName As String
    For Each CurrentSheet In GivenWorkbook.Worksheets
        For Each CurrentPivotTable In CurrentSheet.PivotTables
            ValidConstName = MakeValidConstName(CurrentPivotTable.Name)
            Output = Output & "Public Const " & ValidConstName _
                     & " As String =""" & CurrentPivotTable.Name & """" & vbNewLine
        Next CurrentPivotTable
    Next CurrentSheet
    If Output <> vbNullString Then Output = Left$(Output, Len(Output) - Len(vbNewLine))
    GetPivotTableNameAsConstDeclaration = Output

End Function

Public Function ConvertStringToConstDeclaration(ByVal GivenText As String) As String

    Dim ValidConstName As String
    ValidConstName = MakeValidConstName(GivenText)
    Const CONST_DECLARATION_PATTERN As String = "Public Const {0} As String = ""{1}"""
    ConvertStringToConstDeclaration = BeautifyString(CONST_DECLARATION_PATTERN, Array(ValidConstName, GivenText))

End Function

Public Function MakeValidName(ByVal GivenName As Variant) As String

    'This function just remove all the space so that it can be used as control name.
    'This givenName should not be start with number as control name must need to be start with string.

    MakeValidName = Replace(GivenName, Space(1), vbNullString)

End Function

Public Function InsertNNewLine(ByVal NumberOfNewLine As Long) As String

    '@Description("This will return N number of newline as string")
    '@Dependency("No Dependency")
    '@ExampleCall : InsertNNewLine(3)
    '@Date : 14 October 2021 07:00:04 PM

    Dim Output As String
    Dim Counter As Long
    For Counter = 1 To NumberOfNewLine
        Output = Output & vbNewLine
    Next Counter
    InsertNNewLine = Output

End Function

Public Function CopySheet(ByVal SourceSheet As Worksheet _
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

''Dependency : Microsoft ActiveX Data Objects 6.1 Library Or other version(6.0,2.8,2.6 etc)
''Example call : Set NewRecord = ReadDataFromExcelADO("C:\Users\Ismail\Desktop\Test\", "ABC.xlsx", "Sheet1","$A:B" )
''Last argument dollar sign is important. If you miss that then you will get :
''       Error Number : -2147217865
''       Error Description : The Microsoft Access database engine could not find the object 'Sheet1A:B'. Make sure the object exists and
''                                    that you spell its name and the path name correctly. If 'Sheet1A:B' is not a local object, check your
''                                    network connection or contact the server administrator.
'
'Public Function ReadDataFromExcelADO(FolderPath As String, FileName As String, SheetName As String, _
'                                     RangeAddress As String) As ADODB.Recordset
'
'    Dim DataRecord As ADODB.Recordset
'    Dim ConnectionString As String
'    If Right(FolderPath, 1) = Application.PathSeparator Then
'        ConnectionString = FolderPath & FileName
'    Else
'        ConnectionString = FolderPath & Application.PathSeparator & FileName
'    End If
'    ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
'                       "Data Source=" & ConnectionString & ";" & "Extended Properties=" & """" & "Excel 12.0 Xml;HDR=NO;IMEX=1" & """"
'    Dim SqlQuery As String
'    SqlQuery = "SELECT * FROM [" & SheetName & RangeAddress & "]"
'    Set DataRecord = New ADODB.Recordset
'    Dim Connector As ADODB.Connection
'    Set Connector = New ADODB.Connection
'    Connector.Open ConnectionString
'
'    DataRecord.Open SqlQuery, Connector, adOpenStatic, adLockReadOnly
'    Set ReadDataFromExcelADO = DataRecord
'
'    Set Connector = Nothing
'    Set DataRecord = Nothing
'
'End Function



Public Function IsSheetExist(ByVal SheetTabName As String _
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

Public Function GetSheetsToDictionary(Optional ByVal GivenWorkbook As Workbook) As Object

    '@Dependency : Microsoft Scripting Runtime. To Enable that reference goto Tools>>Reference> Check Microsoft Scripting Runtime
    '@ExampleCall : GetSheetsToDictionary() if you want from thisworkbook
    '@ExampleCall : GetSheetsToDictionary(Workbooks("Name of the workbook"))
    '@EarlyBinding : To use Early binding you need to reference scrunn.dll following first process.

    'For Late Binding

    Dim AllSheet As Object
    Set AllSheet = CreateObject("Scripting.Dictionary")

    'For Early Binding
    'Public Function GetSheetsCollection(Optional GivenWorkbook As Workbook) As Scripting.Dictionary
    '
    '    Dim AllSheet As Scripting.Dictionary
    '    Set AllSheet = New Scripting.Dictionary

    Dim CurrentWorksheet As Worksheet
    If GivenWorkbook Is Nothing Then
        Set GivenWorkbook = ThisWorkbook
    End If
    For Each CurrentWorksheet In GivenWorkbook.Worksheets
        AllSheet.Add CurrentWorksheet.Name, CurrentWorksheet
    Next CurrentWorksheet
    Set GetSheetsToDictionary = AllSheet
    Set AllSheet = Nothing

End Function

Public Function GetSheetsIntoCollection(Optional ByVal GivenWorkbook As Workbook) As Collection

    '@Description("Get all the worksheet into a Collection using their name as key.")
    '@Dependency("No Dependency")
    '@ExampleCall :GetSheetsIntoCollection() if you want from thisworkbook
    '@ExampleCall : GetSheetsIntoCollection(Workbooks("Name of the workbook"))
    '@Date : 30 November 2021 11:14:52 PM

    Dim AllSheet As Collection
    Set AllSheet = New Collection

    Dim CurrentWorksheet As Worksheet
    If GivenWorkbook Is Nothing Then
        Set GivenWorkbook = ThisWorkbook
    End If
    For Each CurrentWorksheet In GivenWorkbook.Worksheets
        AllSheet.Add CurrentWorksheet, CurrentWorksheet.Name
    Next CurrentWorksheet
    Set GetSheetsIntoCollection = AllSheet
    Set AllSheet = Nothing

End Function

Public Function IsExistInCollection(ByVal GivenCollection As Collection _
                                    , ByVal Key As String) As Boolean

    '@Description("This is for testing if a key is present in a collection or not.")
    '@Dependency("No Dependency")
    '@ExampleCall : IsExistInCollection(InputCollection,"Key")
    '@Date : 14 October 2021 07:04:18 PM

    On Error GoTo NotExist
    'If item is not present then it will throw error on the first line.
    GivenCollection.Item Key
    IsExistInCollection = True
    Exit Function

NotExist:
    IsExistInCollection = False
    Err.Clear
    On Error GoTo 0

End Function

Public Function GetTextFileContent(ByVal FullFilePath As String) As String

    '@Description("This will return the content of text file as String")
    '@Dependency("IsFileExist")
    '@ExampleCall : GetTextFileContent("C:\Users\Ismail\Desktop\Test\Outlook Image Email.txt")
    '@Date : 14 October 2021 07:08:33 PM

    If Not IsFileExist(FullFilePath) Then
        GetTextFileContent = vbNullString
        Exit Function
    End If
    Dim FileNo As Integer
    FileNo = FreeFile
    Open FullFilePath For Input As #FileNo
    GetTextFileContent = Input$(LOF(FileNo), FileNo)
    Close #FileNo

End Function

Public Function GetTextFileContentIntoArray(ByVal FullFilePath As String) As Variant

    '@Description("Read each line of the text file into an array.")
    '@Dependency("FXCollection.IsFileExist,FXCollection.GetTextFileContentIntoCollection,FXCollection.CollectionToArray")
    '@ExampleCall : FXCollection.GetTextFileContentIntoArray(SelectedFilePath)
    '@Date : 19 December 2021 08:17:21 PM

    If Not IsFileExist(FullFilePath) Then
        GetTextFileContentIntoArray = vbEmpty
        Exit Function
    End If

    Dim AllLines As Collection
    Set AllLines = GetTextFileContentIntoCollection(FullFilePath)
    GetTextFileContentIntoArray = CollectionToArray(AllLines)

End Function

Public Function DictionaryToArray(ByVal GivenDictionary As Object) As Variant

    '@Description("This will return the key and item as an array from dictionary.")
    ''@Dependency("Microsoft scripting Runtime. If you want to use early binding then you need to reference that.")
    '@ExampleCall: DictionaryToArray(Dict)
    '@Date : 14 October 2021 07:11:46 PM

    'Public Function DictionaryKeyToArray(ByVal GivenDictionary As Scripting.Dictionary) As Variant

    Dim Result As Variant
    ReDim Result(1 To GivenDictionary.Count, 1 To 2) As Variant
    Dim Counter As Long
    Counter = 1
    Dim CurrentKey As Variant
    For Each CurrentKey In GivenDictionary.Keys
        Result(Counter, 1) = CurrentKey
        Result(Counter, 2) = GivenDictionary.Item(CurrentKey)
        Counter = Counter + 1
    Next CurrentKey
    DictionaryToArray = Result

End Function

Public Function GetTextFileContentIntoCollection(ByVal FullFilePath As String) As Collection

    '@Description : This function will read all the line from that text file and add into collection.
    '@Dependency : IsFileExist
    '@ExampleCall: GetTextFileContentIntoCollection("C:\Users\Ismail\Desktop\Test\Read Text File Project\Cleaned_Data.txt")

    If Not IsFileExist(FullFilePath) Then
        Set GetTextFileContentIntoCollection = New Collection
        Exit Function
    End If

    Dim AllLines As Collection
    Set AllLines = New Collection
    Dim FileNo As Integer
    FileNo = FreeFile
    Dim CurrentLineText As String
    Open FullFilePath For Input As #FileNo
    Do While Not EOF(FileNo)
        Line Input #FileNo, CurrentLineText
        AllLines.Add CurrentLineText
    Loop
    Close #FileNo
    Set GetTextFileContentIntoCollection = AllLines

End Function

Public Function OLEObjectsNameAsConstDeclartion(Optional ByVal GivenWorkbook As Workbook) As String

    '@Description("It will return all the OLEObject name as const.")
    '@Dependency("MakeValidConstName and BeautifyString function")
    '@ExampleCall : OLEObjectsNameAsConstDeclartion(Thisworkbook)
    '@Date : 14 October 2021 07:22:41 PM

    If GivenWorkbook Is Nothing Then
        Set GivenWorkbook = ActiveWorkbook
    End If

    Dim Result As String
    Dim CurrentSheet As Worksheet
    For Each CurrentSheet In GivenWorkbook.Worksheets
        Const CONST_DECLARATION_PATTERN As String = "Public Const {0} As String = ""{1}""" & vbNewLine
        Dim CurrentOLEObject As OLEObject
        Dim ValidConstName As String
        For Each CurrentOLEObject In CurrentSheet.OLEObjects
            ValidConstName = MakeValidConstName(CurrentOLEObject.Name)
            Result = Result & BeautifyString(CONST_DECLARATION_PATTERN, Array(ValidConstName, CurrentOLEObject.Name))
        Next CurrentOLEObject
    Next CurrentSheet
    If Result <> vbNullString Then Result = Left$(Result, Len(Result) - Len(vbNewLine))
    OLEObjectsNameAsConstDeclartion = Result

End Function

Public Function ShapesNameAsConst(Optional ByVal GivenWorkbook As Workbook) As String

    '@Description("It will give the shapes name as const")
    '@Dependency("MakeValidConstName and BeautifyString function")
    '@ExampleCall : ShapesNameAsConst(Thisworkbook)
    '@Date : 03 December 2021 11:24:20 AM

    If GivenWorkbook Is Nothing Then
        Set GivenWorkbook = ActiveWorkbook
    End If
    Dim Result As String
    Dim CurrentSheet As Worksheet
    For Each CurrentSheet In GivenWorkbook.Worksheets
        Const CONST_DECLARATION_PATTERN As String = "Public Const {0} As String = ""{1}""" & vbNewLine
        Dim CurrentShape As Shape
        Dim ValidConstName As String
        For Each CurrentShape In CurrentSheet.Shapes
            ValidConstName = MakeValidConstName(CurrentShape.Name)
            Result = Result & BeautifyString(CONST_DECLARATION_PATTERN, Array(ValidConstName, CurrentShape.Name))
        Next CurrentShape
    Next CurrentSheet
    If Result <> vbNullString Then Result = Left$(Result, Len(Result) - Len(vbNewLine))
    ShapesNameAsConst = Result

End Function

Public Function MakeValidConstName(ByVal GivenName As String) As String

    '@Description("This will make a valid constant name according to covention")
    '@Dependency("IsCapitalLetter,IsTextPresent function")
    '@ExampleCall : MakeValidConstName("SettingsTable") >> SETTINGS_TABLE
    '@Date : 14 October 2021 08:05:41 PM
    'Input Param : "CamelCASERules","IndexID","CamelCASE","aID","theIDForUSGovAndDOD","TheID_","_IDOne"
    'Expected Output : CAMEL_CASE_RULES,INDEX_ID,CAMEL_CASE,A_ID,THE_ID_FOR_US_GOV_AND_DOD,THE_ID,ID_ONE
    
    Dim Result As String
    Result = GivenName
    Result = GetOnlyAlphanumericCharcter(Result, vbNullString, True, True)
    If UCase$(Result) = Result Then
        MakeValidConstName = Replace(Result, Space(1), "_")
        Exit Function
    End If

    If Trim(Result) = vbNullString Then
        Err.Raise 13, "MakeValidConstName Function", "Constant Name can't be Nullstring"
    ElseIf Not (Left$(Result, 1) Like "[A-Za-z]") Then
        Err.Raise 13, "MakeValidConstName Function", "Constant Name should be start with A-Z or a-z."
    End If

    Result = ConvertVariableNameToSentence(Result)
    MakeValidConstName = UCase$(Replace(Result, Space(1), "_"))

End Function

Public Function Trim(ByVal Text As Variant) As String
    Trim = Application.WorksheetFunction.Trim(Text)
End Function

Public Function ConvertVariableNameToSentence(ByVal GivenName As String) As String

    '@Description("Convert VBA or other Variable declaration to sentence.")
    '@Dependency("No Dependency")
    '@ExampleCall :ConvertVariableNameToSentence("theIDForUSGovAndDOD") >>the ID For US Gov And DOD
    '@Date : 18 May 2023 10:59:50 AM
    '@PossibleError:

    GivenName = GetOnlyAlphanumericCharcter(GivenName, vbNullString)
    If UCase$(GivenName) = GivenName Then
        ConvertVariableNameToSentence = GivenName
        Exit Function
    End If

    'Put One Space in between Last Caps one . e.g. > CourierByDHL > CourierBy DHL
    Dim LastCAPSWord As String
    Dim Counter As Long
    Dim CurrentCharacter As String
    For Counter = Len(GivenName) To 1
        CurrentCharacter = Mid$(GivenName, Counter, 1)
        If IsCapitalLetter(CurrentCharacter) Then
            LastCAPSWord = CurrentCharacter & LastCAPSWord
        End If
    Next Counter

    If LastCAPSWord <> vbNullString Then
        GivenName = Left$(GivenName, Len(GivenName) - Len(LastCAPSWord)) _
                    & " " & LastCAPSWord
    End If

    'Now for each word we need to put space when lower case to upper case transition is happening. e.g. CamelCASERules > Camel CASERules

    Dim Words As Variant
    Words = Split(GivenName, " ")
    GivenName = vbNullString
    Dim CurrentWord As String
    For Counter = LBound(Words) To UBound(Words)
        CurrentWord = Words(Counter)
        CurrentWord = PutSpaceOnLowerCaseToUpperCaseTransition(CurrentWord)
        GivenName = GivenName & Space(1) & CurrentWord
    Next Counter
    'Remove Initial Space
    If GivenName <> vbNullString Then GivenName = Mid$(GivenName, 2)

    Words = Split(GivenName, " ")
    GivenName = vbNullString
    For Counter = LBound(Words) To UBound(Words)
        CurrentWord = Words(Counter)
        If UCase$(CurrentWord) <> CurrentWord Then
            CurrentWord = PutSpaceBeforeLastCapsFromStart(CurrentWord)
        End If
        GivenName = GivenName & Space(1) & CurrentWord
    Next Counter
    'Remove Initial Space
    If GivenName <> vbNullString Then GivenName = Mid$(GivenName, 2)
    ConvertVariableNameToSentence = GivenName

End Function

Public Function PutSpaceOnLowerCaseToUpperCaseTransition(ByVal CurrentWord As String) As String

    Dim Result As String
    Dim Index As Long
    Dim CurrentCharacter As String
    Dim NextCharacter As String
    For Index = 1 To Len(CurrentWord) - 1
        CurrentCharacter = Mid$(CurrentWord, Index, 1)
        NextCharacter = Mid$(CurrentWord, Index + 1, 1)
        Result = Result & CurrentCharacter
        If Not IsCapitalLetter(CurrentCharacter) And IsAlphabet(CurrentCharacter) _
           And IsCapitalLetter(NextCharacter) Then
            Result = Result & Space(1)
        End If
    Next Index
    If CurrentWord <> vbNullString Then Result = Result & Right$(CurrentWord, 1)
    PutSpaceOnLowerCaseToUpperCaseTransition = Result

End Function

Public Function IsAlphabet(ByVal Char As String) As Boolean

    Dim CharCode As Long
    CharCode = Asc(LCase$(Char))
    IsAlphabet = (CharCode >= Asc("a") And CharCode <= Asc("z"))

End Function

Public Function PutSpaceBeforeLastCapsFromStart(ByVal CurrentWord As String) As String

    'PutSpaceBeforeLastCapsFromStart("CASERules") >> "CASE Rules"

    If CurrentWord = vbNullString Then Exit Function

    If UCase$(CurrentWord) = CurrentWord Then
        PutSpaceBeforeLastCapsFromStart = CurrentWord
        Exit Function
    End If

    Dim Result As String
    Dim Index As Long
    Dim CurrentCharacter As String
    Dim NextCharacter As String
    Result = Left$(CurrentWord, 1)
    If Not IsCapitalLetter(Result) Then
        PutSpaceBeforeLastCapsFromStart = CurrentWord
        Exit Function
    End If

    For Index = 2 To Len(CurrentWord) - 1
        CurrentCharacter = Mid$(CurrentWord, Index, 1)
        NextCharacter = Mid$(CurrentWord, Index + 1, 1)
        If IsCapitalLetter(CurrentCharacter) And Not IsCapitalLetter(NextCharacter) Then
            Result = Result & Space(1)
        End If
        Result = Result & CurrentCharacter
    Next Index
    If CurrentWord <> vbNullString Then Result = Result & Right$(CurrentWord, 1)
    PutSpaceBeforeLastCapsFromStart = Result

End Function

Public Function GetOnlyAlphanumericCharcter(ByVal FromText As String _
                                            , Optional ByVal ReplaceOtherCharacterWith As String = " " _
                                             , Optional ByVal IsKeepSpace As Boolean = True _
                                             , Optional ByVal IsTrimAfter As Boolean = False) As String

    Dim Result As String
    Dim Index As Long
    Dim CurrentCharacter As String

    For Index = 1 To Len(FromText)
        CurrentCharacter = Mid$(FromText, Index, 1)
        If (CurrentCharacter Like "[A-Za-z0-9]") Or (CurrentCharacter = Space(1) And IsKeepSpace) Then
            Result = Result & CurrentCharacter
        Else
            Result = Result & ReplaceOtherCharacterWith
        End If
    Next Index

    If IsTrimAfter Then Result = Trim(Result)

    GetOnlyAlphanumericCharcter = Result

End Function

Public Function IsCapitalLetter(ByVal GivenLetter As String) As Boolean

    '@Description("This will check if a given character is Capital letter > A-Z..It will throw error if length of the letter is more than 1")
    '@Dependency("No Dependency")
    '@ExampleCall : IsCapitalLetter(CurrentCharacter)
    '@Date : 14 October 2021 10:23:19 PM

    If Len(GivenLetter) > 1 Then
        Err.Raise 13, "IsCapitalLetter Function", "Given Letter need to be one character String"
    End If
    If GivenLetter = vbNullString Then
        Err.Raise 5, "IsCapitalLetter Function", "Given Letter can't be nullstring"
    End If

    Const ASCII_CODE_FOR_A As Integer = 65
    Const ASCII_CODE_FOR_Z As Integer = 90
    Dim ASCIICodeForGivenLetter As Integer
    ASCIICodeForGivenLetter = Asc(GivenLetter)
    IsCapitalLetter = (ASCIICodeForGivenLetter >= ASCII_CODE_FOR_A _
                       And ASCIICodeForGivenLetter <= ASCII_CODE_FOR_Z)

End Function

Public Function IsTextPresent(ByVal SearchInText As String, ByVal SearchForText As String) As Boolean

    '@Description : This function will check if a text is present in another on text or not. Comparing is case In Sensitive

    IsTextPresent = (InStr(1, SearchInText, SearchForText, vbTextCompare) <> 0)

End Function

Public Function TextBetweenDelimiter(ByVal GivenText As String _
                                     , ByVal FirstDelimiter As String _
                                      , ByVal SecondDelimiter As String) As String

    '@Description : This function will retrive text between two given delimiter. Comparing is case In Sensitive
    '@This function will return vbNullString if any of the argument is vbNullString or if any delimiter isn't present.
    '@ExampleCall : TextBetweenDelimiter("1997ismail.hosen@gmail.com", "@", ".")

    Dim Result As String
    Dim FirstDelimiterPosition As Long
    If GivenText = vbNullString Or FirstDelimiter = vbNullString Or SecondDelimiter = vbNullString Then
        Result = vbNullString
    Else
        FirstDelimiterPosition = InStr(1, GivenText, FirstDelimiter, vbTextCompare)
        If FirstDelimiterPosition <> 0 Then
            Dim SecondDelimiterPosition As Long
            SecondDelimiterPosition = InStr(FirstDelimiterPosition + Len(FirstDelimiter), GivenText, SecondDelimiter, vbTextCompare)
            If SecondDelimiterPosition <> 0 Then
                Result = Mid$(GivenText, FirstDelimiterPosition + Len(FirstDelimiter), SecondDelimiterPosition - FirstDelimiterPosition - Len(FirstDelimiter))
            End If
        End If
    End If
    TextBetweenDelimiter = Result

End Function

Public Function TextBeforeDelimiter(ByVal GivenText As String _
                                    , ByVal GivenDelimiter As String) As String

    '@Description : This function will retrive text before that given delimiter. Comparing is case In Sensitive
    '@This will return vbNullString if any argument is vbNullString or if that delimiter isn't present
    '@ExampleCall : TextBeforeDelimiter("1997ismail.hosen@gmail.com", "@")

    Dim Result As String
    Dim DelimiterPosition As Long
    If GivenText = vbNullString Or GivenDelimiter = vbNullString Then
        Result = vbNullString
    Else
        DelimiterPosition = InStr(1, GivenText, GivenDelimiter, vbTextCompare)
        If DelimiterPosition <> 0 Then
            Result = Left$(GivenText, DelimiterPosition - 1)
        End If
    End If
    TextBeforeDelimiter = Result

End Function

Public Function TextAfterDelimiter(ByVal GivenText As String _
                                   , ByVal GivenDelimiter As String) As String

    '@Description : This function will retrive text after that given delimiter. Comparing is case In Sensitive
    '@This will return vbNullString if any argument is vbNullString or if that delimiter isn't present
    '@ExampleCall : TextAfterDelimiter("1997ismail.hosen@gmail.com", "@")

    Dim Result As String
    Dim DelimiterPosition As Long
    If GivenText = vbNullString Or GivenDelimiter = vbNullString Then
        Result = vbNullString
    Else
        DelimiterPosition = InStr(1, GivenText, GivenDelimiter, vbTextCompare)
        If DelimiterPosition <> 0 Then
            Result = Right$(GivenText, Len(GivenText) - DelimiterPosition - Len(GivenDelimiter) + 1)
        End If
    End If
    TextAfterDelimiter = Result

End Function

Public Function OnlyDigitPart(ByVal GivenText As String) As Collection

    '@Description : This will return all the digit as collection. It will consider a consequitive number as a item of collection.
    '@Dependency : IsDigit function.
    '@ExampleCall : OnlyDigitPart ("1997ismail.hosen234@gmail.com") >> 1997,234

    Dim Result As Collection
    Set Result = New Collection
    Dim IsNewStart As Boolean
    Dim CurrentCharIndex As Long
    Dim LastDigitSeenAtIndex As Long
    Dim CurrentChar As String
    For CurrentCharIndex = 1 To Len(GivenText)
        CurrentChar = Mid$(GivenText, CurrentCharIndex, 1)
        If IsDigit(CurrentChar) Then
            If Not IsNewStart Then
                IsNewStart = True
                LastDigitSeenAtIndex = CurrentCharIndex
            End If
        Else
            If IsNewStart And CurrentChar <> "." Then
                IsNewStart = False
                Result.Add Mid$(GivenText, LastDigitSeenAtIndex, CurrentCharIndex - LastDigitSeenAtIndex)
            End If
        End If
    Next CurrentCharIndex
    If IsNewStart Then
        Result.Add Mid$(GivenText, LastDigitSeenAtIndex, CurrentCharIndex - LastDigitSeenAtIndex)
    End If
    Set OnlyDigitPart = Result

End Function

Public Function OnlyNonDigitPart(ByVal GivenText As String) As Collection

    '@Description : This will return all the non digit as collection. It will consider a consequitive charset as a item of collection.
    '@Dependency : IsDigit function.
    '@ExampleCall : OnlyNonDigitPart("1997ismail.hosen234@gmail.com") >> ismail.hosen,@gmail.com

    Dim Result As Collection
    Set Result = New Collection
    Dim IsNewStart As Boolean
    Dim CurrentCharIndex As Long
    Dim LastNonDigitSeenAtIndex As Long
    For CurrentCharIndex = 1 To Len(GivenText)
        If Not IsDigit(Mid$(GivenText, CurrentCharIndex, 1)) Then
            If Not IsNewStart Then
                IsNewStart = True
                LastNonDigitSeenAtIndex = CurrentCharIndex
            End If
        Else
            If IsNewStart Then
                IsNewStart = False
                Result.Add Mid$(GivenText, LastNonDigitSeenAtIndex, CurrentCharIndex - LastNonDigitSeenAtIndex)
            End If
        End If
    Next CurrentCharIndex
    If IsNewStart Then
        Result.Add Mid$(GivenText, LastNonDigitSeenAtIndex, CurrentCharIndex - LastNonDigitSeenAtIndex)
    End If
    Set OnlyNonDigitPart = Result

End Function

Public Function IsDigit(ByVal GivenCharacter As String) As Boolean

    '@Description : This will check if a given character is digit or not.
    '@ExampleCall : IsDigit("8")

    Const ASCIICodeForZero As Long = 48
    Const ASCIICodeForNine As Long = 57
    Dim GivenCharacterCode As Integer
    GivenCharacterCode = Asc(GivenCharacter)
    IsDigit = (GivenCharacterCode >= ASCIICodeForZero And GivenCharacterCode <= ASCIICodeForNine)

End Function

Public Function SplitDigitAndNonDigit(ByVal GivenText As String) As Collection

    '@Description : This will return all the digit and non digit as collection. It will consider a consequitive charset as a item of collection.
    '@Dependency : IsDigit function.
    '@ExampleCall : SplitDigitAndNonDigit("1997ismail.hosen234@gmail.com") >> 1997,ismail.hosen,234,@gmail.com

    Dim Result As Collection
    Set Result = New Collection
    Dim IsNewDigitStart As Boolean
    Dim IsNewNonDigitStart As Boolean
    Dim CurrentCharIndex As Long
    Dim LastNonDigitSeenAtIndex As Long
    Dim LastDigitSeenAtIndex As Long
    For CurrentCharIndex = 1 To Len(GivenText)
        If IsDigit(Mid$(GivenText, CurrentCharIndex, 1)) Then
            If IsNewNonDigitStart Then
                IsNewDigitStart = True
                LastDigitSeenAtIndex = CurrentCharIndex
                Result.Add Mid$(GivenText, LastNonDigitSeenAtIndex, CurrentCharIndex - LastNonDigitSeenAtIndex)
                IsNewNonDigitStart = False
            ElseIf Not IsNewDigitStart Then
                IsNewDigitStart = True
                LastDigitSeenAtIndex = CurrentCharIndex
            End If
        Else
            If IsNewDigitStart And Mid$(GivenText, CurrentCharIndex, 1) = "." Then
                Debug.Print "Period Present"
            ElseIf IsNewDigitStart Then
                IsNewNonDigitStart = True
                LastNonDigitSeenAtIndex = CurrentCharIndex
                Result.Add Mid$(GivenText, LastDigitSeenAtIndex, CurrentCharIndex - LastDigitSeenAtIndex)
                IsNewDigitStart = False
            ElseIf Not IsNewNonDigitStart Then
                IsNewNonDigitStart = True
                LastNonDigitSeenAtIndex = CurrentCharIndex
            End If
        End If
    Next CurrentCharIndex

    If IsNewNonDigitStart Then
        Result.Add Mid$(GivenText, LastNonDigitSeenAtIndex, CurrentCharIndex - LastNonDigitSeenAtIndex)
    End If
    If IsNewDigitStart Then
        Result.Add Mid$(GivenText, LastDigitSeenAtIndex, CurrentCharIndex - LastDigitSeenAtIndex)
    End If

    Set SplitDigitAndNonDigit = Result

End Function

Public Function IsEmptyFolder(ByVal FolderPath As String) As Boolean

    '@Description: Check if a folder is empty(No File and no sub folder in it)
    '@Dependency : Microsoft scripting Runtime. If you want to use early binding then you need to reference that.
    '@ExampleCall: IsEmptyFolder("D:\Downloads\PNG File")

    If Right$(FolderPath, Len(Application.PathSeparator)) <> Application.PathSeparator Then
        FolderPath = FolderPath & Application.PathSeparator
    End If

    Dim Manager As Object
    Set Manager = CreateObject("Scripting.FileSystemObject")
    Dim RootFolder As Object
    Set RootFolder = Manager.GetFolder(FolderPath)
    If RootFolder.SubFolders.Count = 0 And RootFolder.Files.Count = 0 Then
        IsEmptyFolder = True
    End If

End Function

Public Function InsetNTab(ByVal NumberOfTab As Long) As String


    '@Description : It will return string of given number of tab
    '@ExampleCall : InsetNTab(3)

    Dim Output As String
    Dim Counter As Long
    For Counter = 1 To NumberOfTab
        Output = Output & vbTab
    Next Counter
    InsetNTab = Output

End Function

Public Function IsMacOS() As Boolean

    '@Description("This Short function will let you know if the current OS is MAC or Windows.")
    '@Dependency("No Dependency")
    '@ExampleCall : IsMacOS
    '@Date : 13 October 2021

    Const WindowsIdentifierPattern As String = "*Windows*"
    IsMacOS = Not (Application.OperatingSystem Like WindowsIdentifierPattern)

End Function

Public Function ReplaceKeyWords(ByVal DoReplaceOn As Variant, ByVal FromAndToMapping As Object, _
                                Optional ByVal IsWholeWordMatch As Boolean = True _
                                , Optional ByVal ComparisionType As VbCompareMethod = vbBinaryCompare) As Variant

    '@Description("This function will do the the ReplaceAll for all item of the given array.")
    '@Dependency("Microsoft Scripting Runtime,FXCollection.ReplaceAll")
    '@ExampleCall : ReplaceKeyWords(TestOnArray, Map, True, vbTextCompare)
    '@ExampleCall : ReplaceKeyWords(TestOnArray, Map, True, vbBinaryCompare)
    '@ExampleCall : ReplaceKeyWords(TestOnArray, Map, False, vbTextCompare)
    '@ExampleCall : ReplaceKeyWords(TestOnArray, Map, False, vbBinaryCompare)
    '@Date : 14 October 2021

    Dim Text As String
    If IsArray(DoReplaceOn) Then
        Dim CurrentRowIndex As Long
        For CurrentRowIndex = LBound(DoReplaceOn, 1) To UBound(DoReplaceOn, 1)
            Dim CurrentColumnIndex As Long
            For CurrentColumnIndex = LBound(DoReplaceOn, 2) To UBound(DoReplaceOn, 2)
                Text = DoReplaceOn(CurrentRowIndex, CurrentColumnIndex)
                DoReplaceOn(CurrentRowIndex, CurrentColumnIndex) = ReplaceAll(Text, FromAndToMapping, _
                                                                              IsWholeWordMatch, ComparisionType)
            Next CurrentColumnIndex
        Next CurrentRowIndex
    Else
        Text = CStr(DoReplaceOn)
        DoReplaceOn = ReplaceAll(Text, FromAndToMapping, IsWholeWordMatch, ComparisionType)
    End If
    ReplaceKeyWords = DoReplaceOn

End Function

Public Function ReplaceAll(ByVal DoReplaceOn As String, ByVal FromAndToMapping As Object, _
                           Optional ByVal IsWholeWordMatch As Boolean = True _
                           , Optional ByVal ComparisionType As VbCompareMethod = vbBinaryCompare) As String

    '@Description("This function will replace all the found keyword with item of dictionary based on other parameter.")
    '@Dependency("Microsoft Scripting Runtime")
    '@ExampleCall : ReplaceAll(SUT, Map, True, vbTextCompare)
    '@ExampleCall : ReplaceAll(SUT, Map, True, vbBinaryCompare)
    '@ExampleCall : ReplaceAll(SUT, Map, False, vbTextCompare)
    '@ExampleCall : ReplaceAll(SUT, Map, False, vbBinaryCompare)
    '@Date : 13 October 2021
    '@SpecialConcern: As this is doing replace on the main text so it may happen that replace is done on the previous replace.
    '--For Example : "Md.Ismail Hosen" and i want to replace "Ho" with "Hosen" and "sensen" with "ssain" then Output will be like this:
    '--Md.Ismail Hosen > Md.Ismail Hosensen >> Md.Ismail Hossain. So be careful about it.

    Dim IsDoReplace As Boolean
    Dim CurrentKey As Variant
    For Each CurrentKey In FromAndToMapping.Keys
        IsDoReplace = True
        If IsWholeWordMatch Then
            IsDoReplace = (StrComp(DoReplaceOn, CurrentKey, ComparisionType) = 0)
        End If
        If IsDoReplace Then
            DoReplaceOn = Replace(DoReplaceOn, CurrentKey, FromAndToMapping.Item(CurrentKey), Compare:=ComparisionType)
        End If
    Next CurrentKey
    ReplaceAll = DoReplaceOn

End Function

Public Function GetTableFromRange(ByVal GivenRange As Range) As ListObject

    '@Description("This function will determine if a cell is under any table and return that table. If not under table then it will return nothing")
    '@Dependency("No Dependency")
    '@ExampleCall :  GetTableFromRange(Sheet3.Range("H10"))
    '@Date : 18 October 2021 01:39:09 PM

    On Error Resume Next
    Set GetTableFromRange = GivenRange.ListObject
    On Error GoTo 0

End Function

Public Function GetObjectsPropertyValue(ByVal ObjectCollection As Collection _
                                        , ByVal PropertiesName As Collection _
                                         , Optional ByVal IsHaveHeader As Boolean = True) As Variant

    '@Description("This function will flatten the collection object and return the properties of that given object collection as an array")
    '@Dependency("IsValidObjectInput function")
    '@ExampleCall : GetObjectsPropertyValue(ObjectCollection, PropertyName)
    '@Date : 05 November 2021 08:36:11 PM
    '@PossibleErrorCase: Both collection is nothing or having no item in them(Error Number 13),Private accessor (450) and Property doesn't exit (438)

    Dim Output As Variant
    If Not IsValidObjectInput(ObjectCollection, PropertiesName) Then
        Err.Raise 13, "Wrong Input argument", "Check if you given a proper input arguemnt or not"
    End If

    Dim Counter As Long
    Dim PropertyName As Variant
    If IsHaveHeader Then
        ReDim Output(1 To ObjectCollection.Count + 1, 1 To PropertiesName.Count)
        For Each PropertyName In PropertiesName
            Counter = Counter + 1
            Output(1, Counter) = PropertyName
        Next PropertyName
        Counter = 2
    Else
        ReDim Output(1 To ObjectCollection.Count, 1 To PropertiesName.Count)
        Counter = 1
    End If

    Dim TotalPropertyCount As Long
    TotalPropertyCount = PropertiesName.Count
    On Error GoTo HandleError
    Dim CurrentObject As Object
    For Each CurrentObject In ObjectCollection
        Dim CurrentColumnIndex As Long
        For CurrentColumnIndex = 1 To TotalPropertyCount
            Output(Counter, CurrentColumnIndex) = CallByName(CurrentObject, PropertiesName.Item(CurrentColumnIndex), VbGet)
        Next CurrentColumnIndex
        Counter = Counter + 1
    Next CurrentObject
    GetObjectsPropertyValue = Output
Cleanup:
    Exit Function

HandleError:
    Select Case Err.Number
        Case 450
            Err.Raise 450, "Property Access Problem", "Check If you have valid Property access or not."
        Case 438
            Err.Raise Err.Number, "Property doesn't exist", "Check if your given property exist or not. Also check spelling"
        Case Else
            Err.Raise Err.Number, Err.Source, Err.Description
    End Select
    Resume Cleanup
    'This is only for debugging purpose.
    Resume

End Function

Private Function IsValidObjectInput(ByVal ObjectCollection As Collection _
                                    , ByVal PropertiesName As Collection) As Boolean

    '@Description("This is helper function for GetObjectsPropertyValue function")
    '@Dependency("No Dependency")
    '@ExampleCall : IsValidObjectInput(ObjectCollection, PropertiesName)
    '@Date : 05 November 2021 08:35:29 PM

    Dim Valid As Boolean
    If ObjectCollection Is Nothing Then
        Valid = False
    ElseIf PropertiesName Is Nothing Then
        Valid = False
    ElseIf ObjectCollection.Count = 0 Then
        Valid = False
    ElseIf PropertiesName.Count = 0 Then
        Valid = False
    ElseIf Not IsObject(ObjectCollection.Item(1)) Then
        Valid = False
    Else
        Valid = True
    End If
    IsValidObjectInput = Valid

End Function

Public Function BeautifyString(ByVal Pattern As String, ByVal PlaceHolderValues As Variant _
                                                       , Optional ByVal StartNumber As Long = 0 _
                                                        , Optional ByVal PlaceHolder As String = "{#}") As String

    '@Description("This function will replace placeholder text with appropriate values")
    '@Dependency("No Dependency")
    '@ExampleCall : BeautifyString("Your Name : {1}   Your Age: {2}",Array("Md.Ismail Hosen", 24),1) >> Your Name : Md.Ismail Hosen   Your Age: 24
    '@ExampleCall : BeautifyString("Your Name : {0}","Md.Ismail Hosen") >> Your Name : Md.Ismail Hosen
    '@ExampleCall : BeautifyString("Your Name : {}","Md.Ismail Hosen") >> Your Name : {} >> Because no place holder value..So if your text has {} this will help
    '@Date : 25 November 2021 07:06:34 PM

    Dim PlaceHolderValue As Variant
    Dim CurrentPlaceHolder As String
    If IsArray(PlaceHolderValues) Then
        For Each PlaceHolderValue In PlaceHolderValues
            CurrentPlaceHolder = Replace(PlaceHolder, "#", StartNumber)
            Pattern = Replace(Pattern, CurrentPlaceHolder, PlaceHolderValue)
            StartNumber = StartNumber + 1
        Next PlaceHolderValue
    Else
        CurrentPlaceHolder = Replace(PlaceHolder, "#", StartNumber)
        Pattern = Replace(Pattern, CurrentPlaceHolder, PlaceHolderValues)
    End If
    BeautifyString = Pattern

End Function

Public Function GetTableCollection(ByVal GivenWorkbook As Workbook) As Collection

    '@Description("This will return a table collection of a given workbook")
    '@Dependency("No Dependency")
    '@ExampleCall : GetTableCollection(ActiveWorkbook)
    '@Date : 03 December 2021 07:34:12 PM

    Dim CurrentSheet As Worksheet
    Dim TableCollection As Collection
    Set TableCollection = New Collection
    For Each CurrentSheet In GivenWorkbook.Worksheets
        Dim Table As ListObject
        For Each Table In CurrentSheet.ListObjects
            TableCollection.Add Table, CurrentSheet.Name & "-" & Table.Name
        Next Table
    Next CurrentSheet
    Set GetTableCollection = TableCollection

End Function

Public Function nPr(ByVal TotalNumberOfItem As Long, ByVal NumberOfItemTaken As Long) As Double

    '@Description("This will calculate the number of permutation(Mathematically nPr")
    '@Dependency("No Dependency")
    '@ExampleCall : nPr(20, 7) >> 390700800
    '@Date : 29 December 2021 09:46:03 PM
    '@Error : Overflow error if cross double variable value limit
    '               Type Mismatch if invalid nPr invoke

    If TotalNumberOfItem < NumberOfItemTaken Or TotalNumberOfItem < 0 Or NumberOfItemTaken < 0 Then
        Err.Raise 13, "Invalid input argument", "Total Number Of Item need to be greater than or equal to Number Of Item Taken"
    End If
    Dim Result As Double
    Result = 1
    Dim Counter As Long

    On Error GoTo MaximumLimitReached

    'nPr = n*(n-1)*(n-2)*(n-3).....(n-r+1)

    For Counter = TotalNumberOfItem To (TotalNumberOfItem - NumberOfItemTaken + 1) Step -1
        Result = Result * Counter
    Next Counter

    Debug.Print "TotalNumberOfItem: " & TotalNumberOfItem & " NumberOfItemTaken: " & NumberOfItemTaken & " nPr: " & Result

    nPr = Result
    Exit Function

MaximumLimitReached:
    HandleOverFlowError Err.Number, Err.Source, Err.Description
    nPr = -1

End Function

Public Function nCr(ByVal TotalNumberOfItem As Long, ByVal NumberOfItemTaken As Long) As Double

    '@Description("This will calculate the number of permutation(Mathematically nPr")
    '@Dependency("No Dependency")
    '@ExampleCall : nCr(20, 7) >> 77520
    '@Date : 29 December 2021 09:46:03 PM
    '@Error : Overflow error if cross double variable value limit
    '               Type Mismatch if invalid nCr invoke

    If TotalNumberOfItem < NumberOfItemTaken Or TotalNumberOfItem < 0 Or NumberOfItemTaken < 0 Then
        Err.Raise 13, "Invalid input argument", "Total Number Of Item need to be greater than or equal to Number Of Item Taken"
    End If

    If NumberOfItemTaken > TotalNumberOfItem - NumberOfItemTaken Then
        NumberOfItemTaken = TotalNumberOfItem - NumberOfItemTaken
        'nCr=nC(n-r)
    End If

    Dim Counter As Long
    Dim Result As Double
    Result = 1

    On Error GoTo MaximumLimitReached

    For Counter = 1 To NumberOfItemTaken
        'nCr = (n(n-1)*(n-2)...(n-r+1))/(1*2*3....r)
        Result = Result * (TotalNumberOfItem - Counter + 1)
        Result = Result / Counter
    Next Counter

    Debug.Print "TotalNumberOfItem: " & TotalNumberOfItem & " NumberOfItemTaken: " & NumberOfItemTaken & " nCr: " & Result

    nCr = Result
    Exit Function

MaximumLimitReached:
    HandleOverFlowError Err.Number, Err.Source, Err.Description
    nCr = -1

End Function

Public Function Factorial(ByVal GivenNumber As Long) As Double

    '@Description("This will calculate the number of permutation(Mathematically nPr")
    '@Dependency("No Dependency")
    '@ExampleCall : Factorial(10) >>3628800
    '@Date : 29 December 2021 09:46:03 PM
    '@Error : Overflow error if cross double variable value limit.
    '               Type Mismatch if invalid number given.

    If GivenNumber = 0 Or GivenNumber = 1 Then
        Factorial = 1
        Exit Function
    ElseIf GivenNumber < 0 Then
        Err.Raise 13, "Factorial", "Need positive number to calculate factorial"
    End If

    On Error GoTo MaximumLimitReached

    Dim Result As Double
    Result = 1
    Dim CurrentNumber As Long
    For CurrentNumber = GivenNumber To 1 Step -1
        Result = Result * CurrentNumber
    Next CurrentNumber

    Factorial = Result
    Exit Function

MaximumLimitReached:
    HandleOverFlowError Err.Number, Err.Source, Err.Description

End Function

Private Sub HandleOverFlowError(ByVal ErrorNumber As Long, ByVal ErrorSource As String _
                                                          , ByVal ErrorDescription As String)

    Select Case ErrorNumber
        Case 0
            MsgBox "As it is error handling so it should not come here..You are doing something bad to handle error. Check Error Handling code and also check if you use Exit Procedure on Cleanup."
            'Todo : Uncomment and Use proper error number and handling
        Case 6
            MsgBox "You have reached the maximum number limit of  a double variable."
        Case Else
            Err.Raise ErrorNumber, ErrorSource, ErrorDescription
    End Select

End Sub

Public Function ChangeDimension(ByVal InputArray As Variant _
                                , ByVal ForwardDirection As ProcessDirection _
                                 , ByVal NumberOfRowOrColumn As Long _
                                  , ByVal FixedIn As ChangedTo) As Variant

    '@Description("This function will change the array dimension to your expected. You can fixed number of row or column")
    '@Dependency("ChangeToNumberOfRow,ChangeToNumberOfColumn")
    '@ExampleCall :  FXCollection.ChangeDimension(InputArray, TopToBottomThenRight, 4, FixedColumn)
    '@Date : 07 January 2022 12:46:58 PM
    '@PossibleError:

    If FixedIn = FixedRow Then
        ChangeDimension = ChangeToNumberOfRow(InputArray, ForwardDirection, NumberOfRowOrColumn)
    Else
        ChangeDimension = ChangeToNumberOfColumn(InputArray, ForwardDirection, NumberOfRowOrColumn)
    End If

End Function

Private Function ChangeToNumberOfColumn(ByVal InputArray As Variant _
                                        , ByVal ForwardDirection As ProcessDirection _
                                         , ByVal NumberOfColumn As Long) As Variant

    Dim NumberOfPossibleRow As Long
    NumberOfPossibleRow = TotalNumberOfDataInArray(InputArray) / NumberOfColumn
    ChangeToNumberOfColumn = ProcessBasedOnDirection(ForwardDirection, InputArray, NumberOfPossibleRow, NumberOfColumn)

End Function

Private Function ChangeToNumberOfRow(ByVal InputArray As Variant _
                                     , ByVal ForwardDirection As ProcessDirection _
                                      , ByVal NumberOfRow As Long) As Variant

    Dim NumberOfPossibleColumn As Long
    NumberOfPossibleColumn = TotalNumberOfDataInArray(InputArray) / NumberOfRow
    ChangeToNumberOfRow = ProcessBasedOnDirection(ForwardDirection, InputArray, NumberOfRow, NumberOfPossibleColumn)

End Function

Private Function ProcessBasedOnDirection(ByVal ForwardDirection As ProcessDirection _
                                         , ByVal InputArray As Variant _
                                          , ByVal NumberOfRow As Long _
                                           , ByVal NumberOfColumn As Long) As Variant

    If ForwardDirection = LeftToRightThenBottom Then
        ProcessBasedOnDirection = ProcessFromLeftToRightThenBottom(InputArray, NumberOfRow, NumberOfColumn)
    Else
        ProcessBasedOnDirection = ProcessFromTopToBottomThenRight(InputArray, NumberOfRow, NumberOfColumn)
    End If

End Function

Private Function TotalNumberOfDataInArray(ByVal InputArray As Variant) As Long

    Dim NumberOfRow As Long
    NumberOfRow = UBound(InputArray, 1) - LBound(InputArray, 1) + 1
    Dim NumberOfColumn As Long
    NumberOfColumn = UBound(InputArray, 2) - LBound(InputArray, 2) + 1
    TotalNumberOfDataInArray = NumberOfRow * NumberOfColumn

End Function

Private Function ProcessFromLeftToRightThenBottom(ByVal InputArray As Variant _
                                                  , ByVal NumberOfRow As Long _
                                                   , ByVal NumberOfColumn As Long) As Variant

    Dim Result As Variant
    ReDim Result(1 To NumberOfRow, 1 To NumberOfColumn)
    Dim CurrentRowNumber As Long
    Dim CurrentColumnNumber As Long
    CurrentRowNumber = 1
    CurrentColumnNumber = 1
    Dim CurrentRowIndex As Long
    For CurrentRowIndex = LBound(InputArray, 1) To UBound(InputArray, 1)
        Dim CurrentColumnIndex As Long
        For CurrentColumnIndex = LBound(InputArray, 2) To UBound(InputArray, 2)
            Result(CurrentRowNumber, CurrentColumnNumber) = InputArray(CurrentRowIndex, CurrentColumnIndex)
            If CurrentColumnNumber = NumberOfColumn Then
                CurrentColumnNumber = 1
                CurrentRowNumber = CurrentRowNumber + 1
            Else
                CurrentColumnNumber = CurrentColumnNumber + 1
            End If
        Next CurrentColumnIndex
    Next CurrentRowIndex
    ProcessFromLeftToRightThenBottom = Result

End Function

Private Function ProcessFromTopToBottomThenRight(ByVal InputArray As Variant _
                                                 , ByVal NumberOfRow As Long _
                                                  , ByVal NumberOfColumn As Long) As Variant

    Dim Result As Variant
    ReDim Result(1 To NumberOfRow, 1 To NumberOfColumn)
    Dim CurrentRowNumber As Long
    Dim CurrentColumnNumber As Long
    CurrentRowNumber = 1
    CurrentColumnNumber = 1
    Dim Item As Variant
    For Each Item In InputArray
        Result(CurrentRowNumber, CurrentColumnNumber) = Item
        If CurrentColumnNumber = NumberOfColumn Then
            CurrentColumnNumber = 1
            CurrentRowNumber = CurrentRowNumber + 1
        Else
            CurrentColumnNumber = CurrentColumnNumber + 1
        End If
    Next Item

    ProcessFromTopToBottomThenRight = Result

End Function

Public Function GetFileName(ByVal FilePath As String) As String

    Dim LastSeparatorIndex As Long
    LastSeparatorIndex = InStrRev(FilePath, Application.PathSeparator)
    If LastSeparatorIndex = 0 Then Exit Function
    GetFileName = Mid$(FilePath, LastSeparatorIndex + 1)

End Function

Public Function GetFileNameWithoutExtension(ByVal FilePath As String) As String

    '@Description("This will extract the file name from the folder path")
    '@Dependency("No Dependency")
    '@ExampleCall : GetFileNameWithoutExtension("C:\Users\Ismail\Desktop\New Project\Wise_existing recipients sample.csv")>> Wise_existing recipients sample
    '@Date : 04 March 2022 11:11:48 AM
    '@PossibleError:

    Dim LastSeparatorIndex As Long
    LastSeparatorIndex = InStrRev(FilePath, Application.PathSeparator)
    Dim LastDotIndex As Long
    LastDotIndex = InStrRev(FilePath, ".")
    GetFileNameWithoutExtension = Mid$(FilePath, LastSeparatorIndex + 1, LastDotIndex - LastSeparatorIndex - 1)

End Function

Public Function IsStartsWith(ByVal TestOnText As String, ByVal TextToMatch As String) As Boolean

    '@Description("Test if a text startswith some partial text.")
    '@Dependency("No Dependency")
    '@ExampleCall : IsStartsWith("Ismail","is") >> True>> IsStartsWith("Ismail","some") >>False
    '@Date : 04 March 2022 11:16:45 AM
    '@PossibleError:

    IsStartsWith = (UCase$(Left$(TestOnText, Len(TextToMatch))) = UCase$(TextToMatch))

End Function

Public Function IsEndsWith(ByVal TestOnText As String, ByVal TextToMatch As String) As Boolean

    '@Description("Test if a text endswith some partial text.")
    '@Dependency("No Dependency")
    '@ExampleCall : IsEndsWith("Ismail","il") >> True>> IsEndsWith("Ismail","al") >>False
    '@Date : 04 March 2022 11:16:45 AM
    '@PossibleError:
    IsEndsWith = (UCase$(Right$(TestOnText, Len(TextToMatch))) = UCase$(TextToMatch))

End Function

Public Function RetriveDataFromFirstSheet(ByVal FilePath As String _
                                          , ByVal StartRow As Long _
                                           , ByVal FindLastRowFromColumn As Long _
                                            , ByVal FindLastColumnFromRow As Long) As Variant

    '@Description("This will retrive data from first sheet of the given FilePath workbook starting from A1")
    '@Dependency("FXCollection.FindLastRowNumber,FXCollection.FindLastColumnNumber")
    '@ExampleCall : RetriveDataFromFirstSheet("C:\Users\Ismail\Desktop\New Project\Wise_existing recipients sample.csv")
    '@Date : 04 March 2022 11:57:50 AM
    '@PossibleError:

    Dim DataBook As Workbook
    Set DataBook = Application.Workbooks.Open(FilePath)
    Dim DataSheet As Worksheet
    Set DataSheet = DataBook.Worksheets(1)
    Dim LastRow As Long
    LastRow = LastUsedRowNumber(DataSheet, FindLastRowFromColumn)
    Dim LastColumn As Long
    LastColumn = LastUsedColumnNumber(DataSheet, FindLastColumnFromRow)
    RetriveDataFromFirstSheet = DataSheet.Range(DataSheet.Cells(StartRow, 1), DataSheet.Cells(LastRow, LastColumn)).Value
    DataBook.Close

End Function

Public Function GetFileExtension(ByVal FileNameOrPath As String) As String

    '@Description("This will extract the file extension from the file path or file name")
    '@Dependency("No Dependency")
    '@ExampleCall : GetFileExtension("A.bas") >> bas
    '@Date : 05 March 2022 11:45:20 PM
    '@PossibleError:

    Dim DotIndex As Long
    DotIndex = InStrRev(FileNameOrPath, ".")
    GetFileExtension = Right$(FileNameOrPath, Len(FileNameOrPath) - DotIndex)

End Function

Public Function HasDynamicFormula(ByVal SelectionRange As Range) As Boolean

    '@Description("This will check if a range has a Dynamic formula or not. If your selection cross dynamic formula section then it will return false. It has to be either a single cell or part of the dynamic formula output range")
    '@Dependency("No Dependency")
    '@ExampleCall : HasDynamicFormula(Sheet1.Range("P5:P10"))=True, HasDynamicFormula(Sheet1.Range("P5:P11"))=False as we have formula upto P10
    '@Date : 23 March 2022 07:41:19 PM
    '@PossibleError:

    On Error Resume Next
    HasDynamicFormula = SelectionRange.HasSpill
    On Error GoTo 0

End Function

Public Function ConvertToReference(ByVal DataSource As Range _
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

Public Function MakeValidDefinedName(ByVal GivenDefinedName As String) As String


    '@Description("This will create a valid name for name range")
    '@Dependency("No Dependency")
    '@ExampleCall :MakeValidDefinedName("1ABC1") = "_1ABC1",MakeValidDefinedName("AB C1") = "_ABC1",MakeValidDefinedName(vbNullString) = "_DefaultName"
    '@Date : 23 March 2022 07:36:46 PM
    '@PossibleError:
    
    Dim Result As String
    If Trim(GivenDefinedName) = vbNullString Then
        Result = "DefaultName"
        GoTo ClenExit
    End If
    
    Result = GivenDefinedName
    If Not Left$(GivenDefinedName, 1) Like "[A-Za-z]" Then
        Result = "_" & Result
    End If
    
    Result = Replace(Result, Space(1), vbNullString)
    
    'Handle Name like(AB25) cell reference.It will add one more underscore if given name is name range
    On Error GoTo Done:
    If Not Range(Result) Is Nothing Then
        If Range(Result).Address(False, False) = Replace(Result, "$", vbNullString) Then
            Result = "_" & Result
        End If
    End If

ClenExit:
    MakeValidDefinedName = Result
    Exit Function
    
Done:
    GoTo ClenExit
    
End Function

Private Function IsInsideTableHeader(ByVal GivenRange As Range) As Boolean


    '@Description("This will check if given range from header of a table.")
    '@Dependency("GetTableFromRange")
    '@ExampleCall : IsInsideTableHeader Selection
    '@Date : 26 March 2022 06:36:15 PM
    '@PossibleError:

    Dim ActiveTable As ListObject
    Set ActiveTable = GetTableFromRange(GivenRange)
    If ActiveTable Is Nothing Then
        IsInsideTableHeader = False
    ElseIf Intersect(ActiveTable.HeaderRowRange, GivenRange) Is Nothing Then
        IsInsideTableHeader = False
    Else
        IsInsideTableHeader = True
    End If

End Function

Public Function IsInsideNamedRange(ByVal GivenRange As Range) As Boolean


    '@Description("This will check if given range is inside of a named range or not. If any cell is fall under any name range then it return true.")
    '@Dependency("No Dependency")
    '@ExampleCall : IsInsideNamedRange Selection
    '@Date : 26 March 2022 06:37:13 PM
    '@PossibleError:

    Dim CurrentNameRange As Name
    Dim ReferredRange As Range
    For Each CurrentNameRange In GivenRange.Parent.Parent.Names
        On Error Resume Next
        Set ReferredRange = CurrentNameRange.RefersToRange
        On Error GoTo 0
        If Not ReferredRange Is Nothing Then
            If ReferredRange.Parent.Name = GivenRange.Parent.Name Then
                If Not Intersect(ReferredRange, GivenRange) Is Nothing Then
                    IsInsideNamedRange = True
                    Exit Function
                End If
            End If
        End If
    Next CurrentNameRange
    IsInsideNamedRange = False

End Function

Public Function SubString(ByVal GivenText As String, ByVal StartIndex As Long _
                                                    , Optional ByVal EndIndex As Long = -1) As String

    '@Description("This will extract part of a string like java.If You don't mention EndIndex then it will extract upto last part.")
    '@Dependency("No Dependency")
    '@ExampleCall : SubString("Ismail Hosen",8)>Hosen
    '@Date : 27 March 2022 11:46:28 AM
    '@PossibleError:Type Mistmatch error if end index less than startindex

    If EndIndex = -1 Then EndIndex = Len(GivenText)
    If StartIndex > EndIndex Then
        Err.Raise 13, "SubString", "StartIndex need to be less or equal to EndIndex"
    End If
    SubString = Mid$(GivenText, StartIndex, EndIndex - StartIndex + 1)

End Function

Public Function GetFormattedValue(ByVal GivenRange As Range) As Variant

    '@Description("This will return the value of the given range and keep the formatting from cell.")
    '@Dependency("No Dependency")
    '@ExampleCall : GetFormattedValue Sheet11.Range("B44:D48")
    '@Date : 04 April 2022 10:18:56 PM
    '@PossibleError:

    Dim Result As Variant
    ReDim Result(1 To GivenRange.Rows.Count, 1 To GivenRange.Columns.Count)
    Dim RowIndex As Long
    For RowIndex = 1 To GivenRange.Rows.Count
        Dim ColumnIndex As Long
        For ColumnIndex = 1 To GivenRange.Columns.Count
            Result(RowIndex, ColumnIndex) = GivenRange.Cells(RowIndex, ColumnIndex).Text
        Next ColumnIndex
    Next RowIndex
    GetFormattedValue = Result

End Function

Public Function CreateObjectsFromArray(ByVal PropertyNameWithValues As Variant _
                                       , ByVal InstanceOfClassHavingEmptyConstructor As Object _
                                        , ByVal CreatorMethodName As String _
                                         , Optional ByVal KeyColumnIndex As Long = -1) As Collection

    '@Description("This is a function which take an array with property name in top row and value in rest of the row
    '                           and create objects with those value. InstanceOfClassHavingEmptyConstructor is just new object of that class
    '                           which has a method by CreatorMethodName to create a new object of the same type.
    '                           possible function signature is below ")
    '
    '                           Public Function CreateMe() As ClassName
    '                               Set CreateMe = New ClassName
    '                           End Function

    '@Dependency("ReasonToBeInValidObjectInstance")
    '@ExampleCall : CreateObjectsFromArray(PropertyNameWithValues,new ClassName,"CreateMe",3)
    '@Date : 15 April 2022 03:03:54 PM
    '@PossibleError:Type Mismatch error

    Dim Reason As String
    Reason = ReasonToBeInValidObjectInstance(InstanceOfClassHavingEmptyConstructor, CreatorMethodName)
    If Reason <> vbNullString Then
        Err.Raise 13, "Invalid Call of CreateObjectsFromArray", Reason
        Exit Function
    ElseIf Not IsArray(PropertyNameWithValues) Then
        Err.Raise 13, "Invalid Array Data", "PropertyNameWithValues should be an array with property name at the top row"
        Exit Function
    End If

    On Error GoTo HandleError

    Dim KeyToObjectMap As Collection
    Set KeyToObjectMap = New Collection
    Dim CurrentRowIndex As Long
    Dim CurrentObject As Object
    For CurrentRowIndex = LBound(PropertyNameWithValues, 1) + 1 To UBound(PropertyNameWithValues, 1)
        Set CurrentObject = CallByName(InstanceOfClassHavingEmptyConstructor, CreatorMethodName, VbMethod)
        Dim FirstRowIndex As Long
        FirstRowIndex = LBound(PropertyNameWithValues, 1)
        Dim CurrentColumnIndex As Long
        For CurrentColumnIndex = LBound(PropertyNameWithValues, 2) To UBound(PropertyNameWithValues, 2)
            Dim PropertyName As String
            PropertyName = CStr(PropertyNameWithValues(FirstRowIndex, CurrentColumnIndex))
            Dim PropertyValue As Variant
            PropertyValue = PropertyNameWithValues(CurrentRowIndex, CurrentColumnIndex)
            CallByName CurrentObject, PropertyName, VbLet, PropertyValue
        Next CurrentColumnIndex
        If KeyColumnIndex = -1 Then
            KeyToObjectMap.Add CurrentObject
        Else
            KeyToObjectMap.Add CurrentObject, CStr(PropertyNameWithValues(CurrentRowIndex, KeyColumnIndex))
        End If
    Next CurrentRowIndex
    Set CreateObjectsFromArray = KeyToObjectMap
Cleanup:
    Exit Function

HandleError:
    Select Case Err.Number
        Case 450
            Debug.Print "Property Access Problem in " & PropertyName & ". Check If you have valid Property access or not."
        Case 438
            Err.Raise Err.Number, "Property doesn't exist", "Check if your given property exist or not. Also check spelling"
        Case Else
            Err.Raise Err.Number, Err.Source, Err.Description
    End Select
    Resume Cleanup
    'This is only for debugging purpose.
    Resume

End Function

Private Function ReasonToBeInValidObjectInstance(ByVal InstanceOfClassHavingEmptyConstructor As Object _
                                                 , ByVal CreatorMethodName As String) As String


    '@Description("This is helper function for CreateObjectsFromArray function")
    '@Dependency("No Dependency")
    '@ExampleCall : ReasonToBeInValidObjectInstance(InstanceOfClassHavingEmptyConstructor, CreatorMethodName)
    '@Date : 15 April 2022 02:58:52 PM

    On Error GoTo HandleError
    If InstanceOfClassHavingEmptyConstructor Is Nothing Then
        ReasonToBeInValidObjectInstance = "InstanceOfClassHavingEmptyConstructor is nothing."
    Else
        Dim NewObject As Object
        Set NewObject = CallByName(InstanceOfClassHavingEmptyConstructor, CreatorMethodName, VbMethod)
        If TypeName(NewObject) <> TypeName(InstanceOfClassHavingEmptyConstructor) Then
            ReasonToBeInValidObjectInstance = "Creator method return different type of object than it should be."
        End If
    End If
Cleanup:
    Exit Function
HandleError:
    Select Case Err.Number
        Case 450
            Err.Raise 450, "Property Access Problem", "Check If you have valid Property access or not."
        Case 438
            ReasonToBeInValidObjectInstance = "Property doesn't exist Check if your given property exist or not. Also check spelling"
            GoTo Cleanup
        Case Else
            Err.Raise Err.Number, Err.Source, Err.Description
    End Select
    Resume Cleanup
    'This is only for debugging purpose.
    Resume

End Function

Public Function ReverseIndexCollection(ByVal InputCollection As Collection) As Collection

    '@Description("This function just reverse the collection. It also doesn't add key to the output  collection")
    '@Dependency("No Dependency")
    '@ExampleCall : ReverseIndexCollection AllIndex
    '@Date : 28 April 2022 03:33:38 PM
    '@PossibleError:

    Dim ReversedIndex As Collection
    Set ReversedIndex = New Collection

    Dim CurrentItemIndex As Long
    For CurrentItemIndex = InputCollection.Count To 1 Step -1
        ReversedIndex.Add InputCollection.Item(CurrentItemIndex)
    Next CurrentItemIndex
    Set ReverseIndexCollection = ReversedIndex

End Function

Public Function UnionOfNonExistableRange(ByVal FirstRange As Range, ByVal SecondRange As Range) As Range


    '@Description("Normal Union throw error if any of those range is nothing. This will not. That is the use of this function.")
    '@Dependency("No Dependency")
    '@ExampleCall :UnionOfNonExistableRange(FirstRange,Nothing)

    If FirstRange Is Nothing And SecondRange Is Nothing Then
        Set UnionOfNonExistableRange = Nothing
    ElseIf FirstRange Is Nothing Then
        Set UnionOfNonExistableRange = SecondRange
    ElseIf SecondRange Is Nothing Then
        Set UnionOfNonExistableRange = FirstRange
    Else
        Set UnionOfNonExistableRange = Union(FirstRange, SecondRange)
    End If

End Function

Public Function HasMergeCells(ByVal GivenRange As Range) As Boolean

    '@Description("To check if a range has any merge cells.")
    '@Dependency("No Dependency")
    '@ExampleCall :
    '@Date : 05 November 2022 12:02:31 AM
    '@PossibleError:

    HasMergeCells = (IsNull(GivenRange.MergeCells) Or GivenRange.MergeCells)

End Function

Public Function IsFolderPath(ByVal GivenPath As String) As Boolean

    '@Description("Check if a path is Folder Path or not.")
    '@Dependency("No Dependency")
    '@ExampleCall : IsFolderPath("C:\Users\USER\Documents")
    '@Date : 23 November 2022 06:29:57 PM
    '@PossibleError:

    IsFolderPath = (Dir(GivenPath, vbNormal) = vbNullString _
                    And Dir(GivenPath, vbDirectory) <> vbNullString)

End Function

Public Function IsFilePath(ByVal GivenPath As String) As Boolean

    '@Description("Check if a path is File Path or not")
    '@Dependency("No Dependency")
    '@ExampleCall : IsFilePath("C:\Users\USER\Documents\Compare Folder - Copy.xlsm")
    '@Date : 23 November 2022 06:27:01 PM
    '@PossibleError:

    IsFilePath = (Dir(GivenPath, vbNormal) = Dir(GivenPath, vbDirectory) _
                  And Dir(GivenPath, vbNormal) <> vbNullString)

End Function

Public Function ConcatenateCollection(ByVal GivenCollection As Collection _
                                      , Optional ByVal Delimiter As String = ",") As String

    '@Description("This will concatenate all the item of a Collection if not an object Collection")
    '@Dependency("No Dependency")
    '@ExampleCall : ConcatenateCollection(ValidNameColl)
    '@Date : 21 January 2023 04:07:21 PM
    '@PossibleError:

    Dim Result As String
    Dim CurrentItem As Variant
    For Each CurrentItem In GivenCollection
        Result = Result & CStr(CurrentItem) & Delimiter
    Next CurrentItem

    If Result = vbNullString Then
        ConcatenateCollection = vbNullString
    Else
        ConcatenateCollection = Left$(Result, Len(Result) - Len(Delimiter))
    End If

End Function

Public Function IsFileExist(ByVal FilePath As String) As Boolean


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

Public Function IsFolderExist(ByVal FolderPath As String) As Boolean


    '@Description("This will check if folder is found or not. Sometimes Dir function gives false result specially for temp/appdata folder")
    '@Dependency("No Dependency")
    '@ExampleCall :
    '@Date : 05 February 2023 12:00:18 PM
    '@PossibleError:

    #If Mac Then
        IsFolderExist = IsFileOrFolderExistsOnMac(FolderPath, False)
    #Else
        Dim FolderManager As Object
        Set FolderManager = CreateObject("Scripting.FileSystemObject")
        IsFolderExist = FolderManager.FolderExists(FolderPath)
        Set FolderManager = Nothing
    #End If

End Function

Public Function MakeValidFileOrFolderNameByEncoding(ByVal FileName As String) As String

    '@Description("This will replace invalid chars with their encoded code.")
    '@Dependency("No Dependency")
    '@ExampleCall :MakeValidFileOrFolderNameByEncoding("Name:Example.pq") , It will return "Name%3AExample.pq"
    '@Date : 11 March 2023 12:33:03 AM
    '@PossibleError:

    Dim InvalidCharAndEncodedValueMap As Variant
    InvalidCharAndEncodedValueMap = InvalidFileOrFolderNameCharacterToEncodedMap()

    Dim ValidFileName As String
    ValidFileName = FileName
    Dim FirstColumnIndex  As Long
    FirstColumnIndex = LBound(InvalidCharAndEncodedValueMap, 2)
    Dim CurrentRowIndex As Long
    For CurrentRowIndex = LBound(InvalidCharAndEncodedValueMap, 1) To UBound(InvalidCharAndEncodedValueMap, 1)
        ValidFileName = VBA.Replace(ValidFileName, InvalidCharAndEncodedValueMap(CurrentRowIndex, FirstColumnIndex) _
                                                  , InvalidCharAndEncodedValueMap(CurrentRowIndex, FirstColumnIndex + 1))

    Next CurrentRowIndex
    MakeValidFileOrFolderNameByEncoding = ValidFileName

End Function

Public Function DecodeValidFileOrFolderName(ByVal FileName As String) As String

    '@Description("This will replace encoded character with their invalid chars")
    '@Dependency("No Dependency")
    '@ExampleCall :DecodeValidFileOrFolderName("Name%3AExample.pq") , It will return "Name:Example.pq"
    '@Date : 11 March 2023 12:33:03 AM
    '@PossibleError:

    Dim InvalidCharAndEncodedValueMap As Variant
    InvalidCharAndEncodedValueMap = InvalidFileOrFolderNameCharacterToEncodedMap()

    Dim ValidFileName As String
    ValidFileName = FileName
    Dim FirstColumnIndex  As Long
    FirstColumnIndex = LBound(InvalidCharAndEncodedValueMap, 2)
    Dim CurrentRowIndex As Long
    For CurrentRowIndex = LBound(InvalidCharAndEncodedValueMap, 1) To UBound(InvalidCharAndEncodedValueMap, 1)
        ValidFileName = VBA.Replace(ValidFileName, InvalidCharAndEncodedValueMap(CurrentRowIndex, FirstColumnIndex + 1) _
                                                  , InvalidCharAndEncodedValueMap(CurrentRowIndex, FirstColumnIndex))

    Next CurrentRowIndex
    DecodeValidFileOrFolderName = ValidFileName

End Function

Public Function GetFullyQualifiedRangeReference(ByVal ForRange As Range _
                                                , Optional ByVal IsAbsoluteRow As Boolean = True _
                                                 , Optional ByVal IsAbsoluteColumn As Boolean = True) As String

    '@Description("This will return fully qualified range ref")
    '@Dependency("No Dependency")
    '@ExampleCall :
    '@Date : 17 May 2023 07:05:24 PM
    '@PossibleError:

    If ForRange Is Nothing Then
        GetFullyQualifiedRangeReference = vbNullString
    Else
        GetFullyQualifiedRangeReference = "'[" & ForRange.Parent.Parent.Name & "]" & ForRange.Parent.Name & "'!" & ForRange.Address(IsAbsoluteRow, IsAbsoluteColumn)
    End If

End Function

Public Function GetSheetQualifiedRangeReference(ByVal ForRange As Range _
                                                , Optional ByVal IsAbsoluteRow As Boolean = True _
                                                 , Optional ByVal IsAbsoluteColumn As Boolean = True) As String

    If ForRange Is Nothing Then
        GetSheetQualifiedRangeReference = vbNullString
    Else

        If IsTextPresent(ForRange.Worksheet.Name, Space(1)) Then
            GetSheetQualifiedRangeReference = "'" & ForRange.Worksheet.Name & "'!" & ForRange.Address(IsAbsoluteRow, IsAbsoluteColumn)
        Else
            GetSheetQualifiedRangeReference = ForRange.Worksheet.Name & "!" & ForRange.Address(IsAbsoluteRow, IsAbsoluteColumn)
        End If

    End If

End Function

Public Function GetRandomBetweenTwoInteger(ByVal Low As Long, ByVal High As Long) As Long

    '@Description("This will generate a random integer between a range")
    '@Dependency("No Dependency")
    '@ExampleCall :GetRandomBetweenTwoInteger(100,200)
    '@Date : 17 May 2023 07:04:20 PM
    '@PossibleError:

    GetRandomBetweenTwoInteger = Int((High - Low + 1) * Rnd + Low)
End Function

Public Function CapitalizeFirstCharacterOfEachWord(ByVal CurrentText As String) As String

    '@Description("This will capitalize first character of each word. Difference between Proper Case and this is that it will not change case of other chars")
    '@Dependency("No Dependency")
    '@ExampleCall :CapitalizeFirstCharacterOfEachWord("Range rEf") >>Range REf
    '@Date : 17 May 2023 07:03:08 PM
    '@PossibleError:

    Dim ArrayOfSingleWords As Variant
    ArrayOfSingleWords = Split(CurrentText, " ")
    Dim ProperCaseText As String
    Dim SingleWord As Variant
    For Each SingleWord In ArrayOfSingleWords
        ProperCaseText = ProperCaseText & " " & UCase$(Left$(SingleWord, 1)) & Mid$(SingleWord, 2)
    Next SingleWord
    ProperCaseText = Mid$(ProperCaseText, 2)
    CapitalizeFirstCharacterOfEachWord = ProperCaseText

End Function

Private Function InvalidFileOrFolderNameCharacterToEncodedMap() As Variant

    Dim InvalidCharAndEncodedValueMap(1 To 9, 1 To 2) As Variant
    InvalidCharAndEncodedValueMap(1, 1) = "\"
    InvalidCharAndEncodedValueMap(1, 2) = "%5C"
    InvalidCharAndEncodedValueMap(2, 1) = "/"
    InvalidCharAndEncodedValueMap(2, 2) = "%2F"
    InvalidCharAndEncodedValueMap(3, 1) = "|"
    InvalidCharAndEncodedValueMap(3, 2) = "%7C"
    InvalidCharAndEncodedValueMap(4, 1) = """"
    InvalidCharAndEncodedValueMap(4, 2) = "%22"
    InvalidCharAndEncodedValueMap(5, 1) = "*"
    InvalidCharAndEncodedValueMap(5, 2) = "%2A"
    InvalidCharAndEncodedValueMap(6, 1) = "?"
    InvalidCharAndEncodedValueMap(6, 2) = "%3F"
    InvalidCharAndEncodedValueMap(7, 1) = "<"
    InvalidCharAndEncodedValueMap(7, 2) = "%3C"
    InvalidCharAndEncodedValueMap(8, 1) = ">"
    InvalidCharAndEncodedValueMap(8, 2) = "%3E"
    InvalidCharAndEncodedValueMap(9, 1) = ":"
    InvalidCharAndEncodedValueMap(9, 2) = "%3A"

    InvalidFileOrFolderNameCharacterToEncodedMap = InvalidCharAndEncodedValueMap

End Function

Public Function GetClipboardData() As String


    '@Description("This will return clipboard text from the clipboard.")
    '@Dependency("No Dependency")
    '@ExampleCall :GetClipboardData()
    '@Date : 18 May 2023 02:46:32 PM
    '@PossibleError:

    GetClipboardData = CreateObject("htmlfile").parentWindow.clipboardData.GetData("text")

End Function

Public Function ConvertArrayFunctionCodeTo1DArray(ByVal InputCodeLines As String) As String

    '@Description("This will generate VBA Code from Array declaration line. Like Properties = Array("A","B") - > ")
    '   ReDim Properties(1 To 2)
    '   Properties(1) = "A"
    '   Properties(2) = "B"
    '@Dependency("No Dependency")
    '@ExampleCall :ConvertArrayFunctionCodeTo1DArray("Properties = Array(""A"",""B"")")
    '@Date : 18 May 2023 02:52:46 PM
    '@PossibleError:

    'Clean up
    InputCodeLines = VBA.Replace(InputCodeLines, " _", vbNullString)
    InputCodeLines = VBA.Replace(InputCodeLines, vbNewLine, vbNullString)
    InputCodeLines = VBA.Replace(InputCodeLines, Chr$(10), vbNullString)

    Const EQUAL_SIGN As String = "="
    Dim ArrayVariableName As String
    ArrayVariableName = Trim(TextBeforeDelimiter(InputCodeLines, EQUAL_SIGN))

    Dim TextProperties As String
    TextProperties = TextBetweenDelimiter(InputCodeLines, "Array(", ")")
    Dim ArrayOfProperties As Variant
    ArrayOfProperties = VBA.Split(TextProperties, ",")
    Dim NumberOfItem As Long
    NumberOfItem = UBound(ArrayOfProperties) - LBound(ArrayOfProperties) + 1

    Dim FinalCode As String
    FinalCode = "    ReDim " & ArrayVariableName & "(1 To " & NumberOfItem & ")"

    Dim CurrentProperty As Variant
    Dim Count As Long
    For Each CurrentProperty In ArrayOfProperties
        Count = Count + 1
        Dim CurrentLineCode As String
        CurrentProperty = VBA.LTrim$(CurrentProperty)
        CurrentProperty = VBA.RTrim$(CurrentProperty)
        CurrentLineCode = ArrayVariableName & "(" & Count & ") = " & CurrentProperty

        FinalCode = FinalCode & vbNewLine & "    " & CurrentLineCode
    Next CurrentProperty

    ConvertArrayFunctionCodeTo1DArray = FinalCode

End Function

Public Function ConvertRangeDataToArrayCode(ByVal FromRange As Range _
                                            , ByVal ArrayVariableName As String) As String

    '@Description("This will generate Array declaration Code. See Example Call.")
    '@Dependency("No Dependency")
    '@ExampleCall : ConvertRangeDataToArrayCode(Range("A1:A3"),"Properties")
    '  Output :
    '    ReDim Properties(1 To 3, 1 To 1 )
    '    Properties(1, 1) = "RangeLabel"
    '    Properties(2, 1) = "RangeReference"
    '    Properties(3, 1) = "Level"
    'Where A1:A3 has "RangeLabel,RangeReference,Level"
    '@Date : 18 May 2023 02:54:42 PM
    '@PossibleError:

    Dim TotalColCount As Long
    TotalColCount = FromRange.Columns.Count
    Dim TotalRowCount As Long
    TotalRowCount = FromRange.Rows.Count

    Dim RowCounter As Long
    RowCounter = 1
    Dim ColumnCounter As Long
    ColumnCounter = 1

    Dim FinalCode As String
    FinalCode = "    ReDim " & ArrayVariableName & "(1 To " & TotalRowCount & ", 1 To " & TotalColCount & " )"

    Dim CurrentValue As Variant
    For Each CurrentValue In FromRange

        Const DOUBLE_QUOTES As String = """"
        If Not IsNumeric(CurrentValue) Then

            If Left$(CurrentValue, Len(DOUBLE_QUOTES)) <> DOUBLE_QUOTES Then
                CurrentValue = DOUBLE_QUOTES & CurrentValue
            End If

            If Right$(CurrentValue, Len(DOUBLE_QUOTES)) <> DOUBLE_QUOTES Then
                CurrentValue = CurrentValue & DOUBLE_QUOTES
            End If

        End If

        Dim CurrentLineCode As String
        CurrentLineCode = ArrayVariableName & "(" & RowCounter & "," & ColumnCounter & ") = " & CurrentValue
        FinalCode = FinalCode & vbNewLine & "    " & CurrentLineCode

        If ColumnCounter < TotalColCount Then
            ColumnCounter = ColumnCounter + 1
        ElseIf RowCounter < TotalRowCount Then
            RowCounter = RowCounter + 1
            ColumnCounter = 1
        End If
    Next CurrentValue
    ConvertRangeDataToArrayCode = FinalCode

End Function

Public Function IsCharacterOfSearchTextFoundInSequence(ByVal SearchInText As String _
                                                       , ByVal SearchForText As String) As Boolean

    '@Description("This will check if all the character of SearchForText is found in sequence in SearchInText.")
    '@Dependency("No Dependency")
    '@ExampleCall : IsCharacterOfSearchTextFoundInSequence("Dim CurrentCharacter As String", "DCuASt") > True
    '@Date : 19 May 2023 07:57:46 PM
    '@PossibleError:

    Dim Index As Long
    Dim CurrentCharacterOfSearchForText As String
    Dim PreviousCharIndex As Long
    PreviousCharIndex = 0
    Dim CurrentCharIndex As Long
    For Index = 1 To Len(SearchForText)
        CurrentCharacterOfSearchForText = Mid$(SearchForText, Index, 1)
        CurrentCharIndex = InStr(PreviousCharIndex + 1, SearchInText, CurrentCharacterOfSearchForText, vbTextCompare)
        If CurrentCharIndex = 0 Then
            IsCharacterOfSearchTextFoundInSequence = False
            Exit Function
        End If
        PreviousCharIndex = CurrentCharIndex
    Next Index
    IsCharacterOfSearchTextFoundInSequence = True

End Function

Public Function IsCharacterOfSearchTextFoundInSequenceAndAtTheStartOfEachWord(ByVal SearchInText As String _
                                                                              , ByVal SearchForText As String) As Boolean

    '@Description("This will try to match in each word.See example call. This function is not 100% perfect")
    '@Dependency("No Dependency")
    '@ExampleCall : IsCharacterOfSearchTextFoundInSequenceAndAtTheStartOfEachWord("Dim CurrentCharacter As String", "DCuAS")->True
    '               because D is found on "Dim" and Cu is found on "CurrentCharacter" and "As" is found on "As" and finally S is found on "String

    '@Date : 19 May 2023 07:54:27 PM
    '@PossibleError:

    SearchForText = Replace(SearchForText, Space(1), vbNullString)
    Dim Words As Variant
    Words = Split(SearchInText)
    Dim CurrentWord As Variant
    Dim Index As Long
    Index = 1
    For Each CurrentWord In Words
        Dim IndexInWord As Long
        IndexInWord = 1
        Do While IndexInWord <= Len(CurrentWord)
            Dim CurrentCharacterOfSearchForText As String
            CurrentCharacterOfSearchForText = LCase$(Mid$(SearchForText, Index, 1))
            Dim CurrentCharacterOfCurrentWord As String
            CurrentCharacterOfCurrentWord = LCase$(Mid$(CStr(CurrentWord), IndexInWord, 1))
            If CurrentCharacterOfCurrentWord = CurrentCharacterOfSearchForText Then
                Index = Index + 1
                IndexInWord = IndexInWord + 1
            Else
                Exit Do
            End If
        Loop
    Next CurrentWord
    'After Last Match we increment it by 1 . So subtract that.
    Index = Index - 1
    IsCharacterOfSearchTextFoundInSequenceAndAtTheStartOfEachWord = (Index = Len(SearchForText))

End Function

Public Function IsAllCharacterPresent(ByVal SearchInText As String, ByVal SearchForText As String) As Boolean

    '@Description("This will check if all character is present no matter in which order.")
    '@Dependency("No Dependency")
    '@ExampleCall :IsAllCharacterPresent("present","rp")>True because we found r and p in "present"
    '@Date : 19 May 2023 08:22:45 PM
    '@PossibleError:

    Dim Index As Long
    Dim CurrentCharacterOfSearchForText As String
    Dim CurrentCharIndex As Long
    For Index = 1 To Len(SearchForText)
        CurrentCharacterOfSearchForText = Mid$(SearchForText, Index, 1)
        CurrentCharIndex = InStr(1, SearchInText, CurrentCharacterOfSearchForText, vbTextCompare)
        If CurrentCharIndex = 0 Then
            IsAllCharacterPresent = False
            Exit Function
        End If
    Next Index
    IsAllCharacterPresent = True

End Function

Public Function IsTwoRangeEqualBasedOnValue(ByVal FirstRange As Range, ByVal SecondRange As Range) As Boolean

    '@Description("This will check if two range is equal or not based on their cell value.")
    '@Dependency("No Dependency")
    '@ExampleCall : IsTwoRangeEqualBasedOnValue(Selection.Areas(1),Selection.Areas(2))
    '@Date : 20 May 2023 02:54:38 PM
    '@PossibleError:

    If FirstRange Is Nothing Or SecondRange Is Nothing Then Exit Function

    If FirstRange.Rows.Count <> SecondRange.Rows.Count _
       Or FirstRange.Columns.Count <> SecondRange.Columns.Count Then
        Exit Function
    End If

    Dim FirstRangeValues As Variant
    FirstRangeValues = FirstRange.Value
    Dim SecondRangeValues As Variant
    SecondRangeValues = SecondRange.Value

    Dim CurrentRowIndex As Long
    Dim CurrentColumnIndex As Long
    For CurrentRowIndex = LBound(FirstRangeValues, 1) To UBound(FirstRangeValues, 1)
        For CurrentColumnIndex = LBound(FirstRangeValues, 2) To UBound(FirstRangeValues, 2)
            Dim FirstRangeCurrentValue As Variant
            FirstRangeCurrentValue = FirstRangeValues(CurrentRowIndex, CurrentColumnIndex)
            Dim SecondRangeCurrentValue As Variant
            SecondRangeCurrentValue = SecondRangeValues(CurrentRowIndex, CurrentColumnIndex)
            If FirstRangeCurrentValue <> SecondRangeCurrentValue Then Exit Function
        Next CurrentColumnIndex
    Next CurrentRowIndex

    IsTwoRangeEqualBasedOnValue = True

End Function

Public Function IsTwoSpillRangeEqualBasedOnValue(ByVal AnyCellInsideFirstSpill As Range _
                                                 , ByVal AnyCellInsideSecondSpill As Range) As Boolean

    '@Description("This will check if two spill range has same value or not. You can select any cell in those two spill")
    '@Dependency("IsTwoRangeEqualBasedOnValue")
    '@ExampleCall :IsTwoSpillRangeEqualBasedOnValue(Selection.Area(1),Selection.Area(2))
    '@Date : 20 May 2023 03:43:38 PM
    '@PossibleError:

    If AnyCellInsideFirstSpill Is Nothing Or AnyCellInsideSecondSpill Is Nothing Then Exit Function
    If Not AnyCellInsideFirstSpill.Cells(1).HasSpill Or Not AnyCellInsideSecondSpill.Cells(1).HasSpill Then Exit Function
    IsTwoSpillRangeEqualBasedOnValue = IsTwoRangeEqualBasedOnValue(AnyCellInsideFirstSpill.Cells(1).SpillParent.SpillingToRange _
                                                                   , AnyCellInsideSecondSpill.Cells(1).SpillParent.SpillingToRange)

End Function

Public Function ConvertEnumMemberNameCaseToConstantCase(ByVal OperationOnText As String _
                                                        , Optional ByVal IsFormat As Boolean = True) As String

    '@Description("This will convert enum declaration code to CONSTANT_CASE. You need to pass whole enum declaration code starting from Enum Name to End Enum")
    '@Dependency("MakeValidConstName,IsTextPresent,TextBeforeDelimiter,TextAfterDelimiter")
    '@ExampleCall :ConvertEnumMemberNameCaseToConstantCase("Enum Case" & vbnewline & "Proper=1" & vbnewline & "Upper" & vbnewline & "LOWER=3" & VBNEWLINE & "End Enum")
    'Output:
    'Enum Case
    '    Proper = 1
    '    Upper
    '    Lower = 3
    'End Enum
    '@Date : 21 May 2023 12:43:06 PM
    '@PossibleError:

    Dim SplittedLinesByNewLine As Variant
    OperationOnText = Replace(OperationOnText, vbNewLine, Chr$(10))
    SplittedLinesByNewLine = Split(OperationOnText, Chr$(10))
    Dim FinalOutputCode As String
    FinalOutputCode = SplittedLinesByNewLine(0)

    Dim CurrentLineIndex As Long
    Dim CurrentLineCode As String
    For CurrentLineIndex = LBound(SplittedLinesByNewLine) + 1 To UBound(SplittedLinesByNewLine) - 1

        'Remove space from end
        CurrentLineCode = RTrim$(SplittedLinesByNewLine(CurrentLineIndex))
        Dim TextBeforeEqualSign As String
        Dim TextAfterEqualSign As String
        Const EQUAL_SIGN As String = "="
        If IsTextPresent(CurrentLineCode, EQUAL_SIGN) Then
            TextBeforeEqualSign = RTrim$(TextBeforeDelimiter(CurrentLineCode, EQUAL_SIGN))
            TextAfterEqualSign = Trim(TextAfterDelimiter(CurrentLineCode, EQUAL_SIGN))
        Else
            TextBeforeEqualSign = RTrim$(CurrentLineCode)
            TextAfterEqualSign = vbNullString
        End If
        Const ONE_SPACE As String = " "
        Dim InitialSpaceCount As Long
        If IsFormat Then
            InitialSpaceCount = 4
        Else
            'Count initial space by replacing space with VbnullString
            InitialSpaceCount = Len(TextBeforeEqualSign) - Len(Replace(TextBeforeEqualSign, ONE_SPACE, vbNullString))
        End If

        TextBeforeEqualSign = MakeValidConstName(Trim(TextBeforeEqualSign))
        If InitialSpaceCount > 0 Then TextBeforeEqualSign = Space(InitialSpaceCount) & TextBeforeEqualSign
        If TextAfterEqualSign <> vbNullString Then TextAfterEqualSign = ONE_SPACE & EQUAL_SIGN & ONE_SPACE & TextAfterEqualSign
        CurrentLineCode = TextBeforeEqualSign & TextAfterEqualSign
        FinalOutputCode = FinalOutputCode & vbNewLine & CurrentLineCode

    Next CurrentLineIndex

    FinalOutputCode = FinalOutputCode & vbNewLine & SplittedLinesByNewLine(UBound(SplittedLinesByNewLine))
    ConvertEnumMemberNameCaseToConstantCase = FinalOutputCode

End Function

Public Function IsTwoTextFileHasSameContent(ByVal FirstFilePath As String _
                                            , ByVal SecondFilePath As String) As Boolean

    '@Description("This will check if two text file has same content or not. If both doesn't exist then it will return true.")
    '@Dependency("GetTextFileContent")
    '@ExampleCall :
    '@Date : 22 May 2023 11:37:39 AM
    '@PossibleError:

    IsTwoTextFileHasSameContent = (GetTextFileContent(FirstFilePath) = GetTextFileContent(SecondFilePath))

End Function

Public Function ConvertSentenceToVariableName(ByVal FromText As String) As String

    '@Description("This will convert a sentence to be a proper variable name.")
    '@Dependency("No Dependency")
    '@ExampleCall : ConvertSentenceToVariableName("This is a Var") >> ThisIsAVar
    '@Date : 05 June 2023 04:14:27 PM
    '@PossibleError:
    
    Dim TextWithoutNumericFromBeginning As String
    TextWithoutNumericFromBeginning = FromText
    
    Dim CurrentCharIndex As Long
    For CurrentCharIndex = 1 To Len(FromText)
        If Mid$(FromText, CurrentCharIndex, 1) Like "[!A-Za-z]" Then
            TextWithoutNumericFromBeginning = Mid$(FromText, CurrentCharIndex + 1)
        Else
            Exit For
        End If
    Next CurrentCharIndex

    Dim ValidVarName As String
    ValidVarName = CapitalizeFirstCharacterOfEachWord(TextWithoutNumericFromBeginning)
    'Invalid Characters URL: https://learn.microsoft.com/en-us/office/vba/language/concepts/getting-started/visual-basic-naming-rules
    Dim InvalidCharacters As Variant
    InvalidCharacters = Array(" ", ".", "!", "@", "&", "$", "#")

    CurrentCharIndex = 0
    For CurrentCharIndex = LBound(InvalidCharacters) To UBound(InvalidCharacters)
        If IsTextPresent(ValidVarName, InvalidCharacters(CurrentCharIndex)) Then
            Dim VariableNameWithoutInvalidCharacters As String
            VariableNameWithoutInvalidCharacters = VBA.Replace(ValidVarName, InvalidCharacters(CurrentCharIndex), vbNullString)
            ValidVarName = VariableNameWithoutInvalidCharacters
        End If
    Next CurrentCharIndex
    
    If Len(ValidVarName) > 255 Then
        ValidVarName = Left$(ValidVarName, 255)
    End If
    
    If IsNumeric(ValidVarName) Then ValidVarName = vbNullString
    
    ' Although based on above url docs it seems other characters are allowed, I am going to limit to only alphanumeric.
    ValidVarName = GetOnlyAlphanumericCharcter(ValidVarName, "", False, True)
    
    ConvertSentenceToVariableName = ValidVarName

End Function

Public Function GetRangeInfoAsJSON(ByVal FromRange As Range) As String

    Dim RangeInfoJSON As String
    RangeInfoJSON = "{"

    Dim RangeInfoPropertyNameVsValueMap As Object
    Set RangeInfoPropertyNameVsValueMap = CreateObject("Scripting.Dictionary")

    With RangeInfoPropertyNameVsValueMap

        .Add "Absolute Address", FromRange.Address
        .Add "Relative Address", FromRange.Address(False, False)
        .Add "Row Number", FromRange.Row
        .Add "Column Number", FromRange.Column
        .Add "Total Rows", FromRange.Rows.Count
        .Add "Total Columns", FromRange.Columns.Count

        'Is Inside Table
        If FromRange.ListObject Is Nothing Then
            .Add "Is Inside Table", "false"
            .Add "Table Name If Inside", vbNullString
        Else
            .Add "Is Inside Table", "true"
            .Add "Table Name If Inside", FromRange.ListObject.Name
        End If

        'Is Inside Pivot Table
        If IsInsidePivotTable(FromRange) Then
            .Add "Is Inside Pivot Table", "true"
            .Add "Pivot Table Name If Inside", FromRange.PivotTable.Name
        Else
            .Add "Is Inside Pivot Table", "false"
            .Add "Pivot Table Name If Inside", vbNullString
        End If

        'Is Inside Named Range
        If IsInsideNamedRange(FromRange) Then
            .Add "Is Inside Named Range", "true"
            .Add "Named Range Name", GetNamedRangeNameIfInside(FromRange)
        Else
            .Add "Is Inside Named Range", "false"
            .Add "Named Range Name", vbNullString
        End If

        'Value (Only If Single Cell)
        If FromRange.Cells.Count = 1 Then
            .Add "Value (Only If Single Cell)", IIf(FromRange.Value, FromRange.Value, vbNullString)
        End If

        'First Cell Formula
        .Add "Formula (First Cell Formula)", FromRange.Cells(1).Formula

        'Number format (First Cell Only)
        .Add "Number format (First Cell Only)", FromRange.Cells(1).NumberFormat

        'Fully Qualified Range Reference
        .Add "Fully Qualified Range Reference", GetFullyQualifiedRangeReference(FromRange)

        'Area Count
        .Add "Area Count.", FromRange.Areas.Count

        Dim CurrentKey As Variant
        For Each CurrentKey In RangeInfoPropertyNameVsValueMap
            Dim CurrentItemValue As Variant
            CurrentItemValue = RangeInfoPropertyNameVsValueMap(CurrentKey)

            If IsNumeric(CurrentItemValue) Or CurrentItemValue = "true" Or CurrentItemValue = "false" Then
                RangeInfoJSON = RangeInfoJSON & vbNewLine & Space(4) & _
                                QUOTATION_SIGN _
                                & CurrentKey & QUOTATION_SIGN _
                                & " : " & CurrentItemValue & ","
            Else
                RangeInfoJSON = RangeInfoJSON & vbNewLine & Space(4) & QUOTATION_SIGN & CurrentKey & _
                                QUOTATION_SIGN & " : " _
                                & QUOTATION_SIGN _
                                & CurrentItemValue & QUOTATION_SIGN & ","
            End If

        Next CurrentKey

        RangeInfoJSON = Left$(RangeInfoJSON, Len(RangeInfoJSON) - 1) & vbNewLine & "}"

        GetRangeInfoAsJSON = RangeInfoJSON
    End With

End Function

Private Function GetNamedRangeNameIfInside(ByVal FromRange As Range) As String

    Dim CurrentNamedRange As Name
    Dim Temp As Range
    For Each CurrentNamedRange In FromRange.Parent.Parent.Names
        If Not CurrentNamedRange.RefersToRange Is Nothing Then
            Set Temp = Intersect(FromRange, CurrentNamedRange.RefersToRange)
            If Not Temp Is Nothing Then
                If Temp.Address = FromRange.Address Then
                    GetNamedRangeNameIfInside = CurrentNamedRange.Name
                    Exit Function
                End If
            End If
        End If
    Next CurrentNamedRange

End Function

Public Function IsInsidePivotTable(ByVal FromRange As Range) As Boolean

    Dim CurrentPivotTable As PivotTable
    For Each CurrentPivotTable In ActiveSheet.PivotTables
        If Not Intersect(FromRange, CurrentPivotTable.TableRange2) Is Nothing Then
            IsInsidePivotTable = True
            Exit For
        End If
    Next CurrentPivotTable

End Function

Public Function IsNothing(ByVal GivenObject As Object) As Boolean
    IsNothing = (GivenObject Is Nothing)
End Function

Public Function IsNotNothing(ByVal GivenObject As Object) As Boolean
    IsNotNothing = (Not GivenObject Is Nothing)
End Function

Public Function IsInsideTable(ByVal CheckForRange As Range) As Boolean
    IsInsideTable = (Not CheckForRange.ListObject Is Nothing)
End Function

Public Function GenerateRandomNumbers(ByVal NumberOfRows As Long, ByVal NumberOfColumns As Long) As Variant


    '@Description("This will generate random number array whose value will be in between 0 and 1")
    '@Dependency("No Dependency")
    '@ExampleCall :
    '@Date : 13 June 2023 09:21:53 AM
    '@PossibleError:

    If NumberOfRows <= 0 Or NumberOfColumns <= 0 Then Exit Function

    Dim ResultArray As Variant
    ReDim ResultArray(1 To NumberOfRows, 1 To NumberOfColumns)

    Dim CurrentRowIndex As Long
    For CurrentRowIndex = LBound(ResultArray, 1) To UBound(ResultArray, 1)
        Dim CurrentColumnIndex As Long
        For CurrentColumnIndex = LBound(ResultArray, 2) To UBound(ResultArray, 2)
            ResultArray(CurrentRowIndex, CurrentColumnIndex) = VBA.Rnd
        Next CurrentColumnIndex
    Next CurrentRowIndex

    GenerateRandomNumbers = ResultArray

End Function

Public Function GenerateRandomNumbersWithinRange(ByVal NumberOfRows As Long, ByVal NumberOfColumns As Long _
                                                                            , ByVal MinValue As Double _
                                                                             , ByVal MaxValue As Double) As Variant

    '@Description("Generate random number array between two number")
    '@Dependency("No Dependency")
    '@ExampleCall :
    '@Date : 13 June 2023 09:21:31 AM
    '@PossibleError:

    If NumberOfRows <= 0 Or NumberOfColumns <= 0 Then Exit Function

    Dim ResultArray As Variant
    ReDim ResultArray(1 To NumberOfRows, 1 To NumberOfColumns)

    Dim CurrentRowIndex As Long
    For CurrentRowIndex = LBound(ResultArray, 1) To UBound(ResultArray, 1)
        Dim CurrentColumnIndex As Long
        For CurrentColumnIndex = LBound(ResultArray, 2) To UBound(ResultArray, 2)
            ResultArray(CurrentRowIndex, CurrentColumnIndex) = (MaxValue - MinValue + 1) * Rnd + MinValue
        Next CurrentColumnIndex
    Next CurrentRowIndex

    GenerateRandomNumbersWithinRange = ResultArray

End Function

Public Function GenerateRandomIntegerNumbersWithinRange(ByVal NumberOfRows As Long, ByVal NumberOfColumns As Long _
                                                                                   , ByVal MinValue As Integer _
                                                                                    , ByVal MaxValue As Integer) As Variant

    '@Description("Generate a random number array between two integer number")
    '@Dependency("No Dependency")
    '@ExampleCall :
    '@Date : 13 June 2023 09:21:11 AM
    '@PossibleError:

    If NumberOfRows <= 0 Or NumberOfColumns <= 0 Then Exit Function
    Dim ResultArray As Variant
    ReDim ResultArray(1 To NumberOfRows, 1 To NumberOfColumns)

    Dim CurrentRowIndex As Long
    For CurrentRowIndex = LBound(ResultArray, 1) To UBound(ResultArray, 1)
        Dim CurrentColumnIndex As Long
        For CurrentColumnIndex = LBound(ResultArray, 2) To UBound(ResultArray, 2)
            ResultArray(CurrentRowIndex, CurrentColumnIndex) = Int((MaxValue - MinValue + 1) * Rnd + MinValue)
        Next CurrentColumnIndex
    Next CurrentRowIndex

    GenerateRandomIntegerNumbersWithinRange = ResultArray

End Function

Public Function Collection(ByVal IsUseKey As Boolean, ParamArray Items() As Variant) As Collection


    '@Description("Get Items into Collection like Array( function.")
    '@Dependency("No Dependency")
    '@ExampleCall :Collection(True, 2,3,4)
    '@Date : 10 June 2023 01:19:01 PM
    '@PossibleError:

    Dim Item As Variant
    Dim Result As Collection
    Set Result = New Collection
    If IsUseKey Then On Error Resume Next
    For Each Item In Items
        If IsUseKey Then
            Result.Add Item, CStr(Item)
        Else
            Result.Add Item
        End If
    Next Item
    If IsUseKey Then On Error GoTo 0
    Set Collection = Result

End Function

Public Function Dictionary(ParamArray Items() As Variant) As Object


    '@Description("Get Items into Dictionary like Array( function.")
    '@Dependency("No Dependency")
    '@ExampleCall :Dictionary(2,3,4)
    '@Date : 10 June 2023 01:19:01 PM
    '@PossibleError:

    Dim Item As Variant
    Dim Result As Object
    Set Result = CreateObject("Scripting.Dictionary")
    On Error Resume Next
    For Each Item In Items
        Result.Add Item, Item
    Next Item
    On Error GoTo 0

    Set Dictionary = Result

End Function

Public Function IsAlphanumericOnly(ByVal ForText As String) As Boolean


    '@Description("This will check if given text contains only alphanumeric character or not.")
    '@Dependency("No Dependency")
    '@ExampleCall :
    '@Date : 13 June 2023 09:20:32 AM
    '@PossibleError:

    Dim Index As Long
    Dim CurrentCharacter As String
    For Index = 1 To Len(ForText)
        CurrentCharacter = Mid$(ForText, Index, 1)
        If Not CurrentCharacter Like "[A-Za-z0-9]" Then
            IsAlphanumericOnly = False
            Exit Function
        End If
    Next Index
    IsAlphanumericOnly = True

End Function

Public Function IsUpperCase(ByVal ForText As String) As Boolean

    '@Description("This will check if given entire text is all upper case character or not")
    '@Dependency("No Dependency")
    '@ExampleCall :
    '@Date : 13 June 2023 09:19:39 AM
    '@PossibleError:

    IsUpperCase = (UCase$(ForText) = ForText)

End Function

Public Function IsLowerCase(ByVal ForText As String) As Boolean

    '@Description("This will check if given entire text is all lower case character or not")
    '@Dependency("No Dependency")
    '@ExampleCall :
    '@Date : 13 June 2023 09:19:39 AM
    '@PossibleError:

    IsLowerCase = (LCase$(ForText) = ForText)

End Function

Public Function IsProperCase(ByVal ForText As String) As Boolean

    '@Description("This will check if given entire text is all proper case character or not")
    '@Dependency("No Dependency")
    '@ExampleCall :
    '@Date : 13 June 2023 09:19:39 AM
    '@PossibleError:

    IsProperCase = (StrConv(ForText, vbProperCase) = ForText)

End Function

Public Function IsPrime(ByVal NumberToCheck As Long) As Boolean

    '@Description("This will check if a given number is prime or not.")
    '@Dependency("No Dependency")
    '@ExampleCall :
    '@Date : 13 June 2023 09:19:21 AM
    '@PossibleError:

    If NumberToCheck < 2 Then Exit Function
    Dim Counter As Long
    For Counter = 2 To Sqr(NumberToCheck)
        If NumberToCheck Mod Counter = 0 Then Exit Function
    Next Counter
    IsPrime = True

End Function

Public Function GetKeyAndValueFromDictionary(ByVal FromDictionary As Object) As Variant

    '@Description("Extract Key and Value from a Dictionary as an array.")
    '@Dependency("No Dependency")
    '@ExampleCall :
    '@Date : 13 June 2023 09:18:57 AM
    '@PossibleError:

    If IsNothing(FromDictionary) Then Exit Function
    If FromDictionary.Count = 0 Then Exit Function

    Dim Result As Variant
    ReDim Result(1 To FromDictionary.Count, 1 To 2)
    Dim Counter As Long
    Dim CurrentKey As Variant
    For Each CurrentKey In FromDictionary.Keys
        Counter = Counter + 1
        Result(Counter, 1) = CurrentKey
        Result(Counter, 2) = FromDictionary.Item(CurrentKey)
    Next CurrentKey
    GetKeyAndValueFromDictionary = Result

End Function

Public Function IsValidEmail(ByVal EmailAddress As String) As Boolean

    '@Description("This will check if a given email address is valid or not using regex.")
    '@Dependency("No Dependency")
    '@ExampleCall :
    '@Date : 13 June 2023 09:18:37 AM
    '@PossibleError:

    Dim RegEx As Object
    Set RegEx = CreateObject("VBScript.RegExp")
    RegEx.Pattern = "^([a-zA-Z0-9_\-\.]+)@((\[[0-9]{1,3}" & _
                    "\.[0-9]{1,3}\.[0-9]{1,3}\.)|(([a-zA-Z0-9\-]+\" & _
                    ".)+))([a-zA-Z]{2,4}|[0-9]{1,3})(\]?)$"
    IsValidEmail = RegEx.Test(EmailAddress)

End Function

Public Function IsValidIP(ByVal IP As String) As Boolean

    '@Description("This will check if a given IP address is valid or not.")
    '@Dependency("No Dependency")
    '@ExampleCall :
    '@Date : 13 June 2023 09:18:18 AM
    '@PossibleError:

    Dim Octets() As String
    Octets = Split(IP, ".")

    If UBound(Octets) <> 3 Then
        IsValidIP = False
        Exit Function
    End If

    Dim Index As Integer
    For Index = 0 To 3
        If Not IsNumeric(Octets(Index)) Then
            IsValidIP = False
            Exit Function
        End If

        If Octets(Index) < 0 Or Octets(Index) > 255 Then
            IsValidIP = False
            Exit Function
        End If
    Next Index
    IsValidIP = True

End Function

Public Function GetButtonRibonXML(ByVal Label As String, ByVal OnAction As String _
                                                        , Optional ByVal ID As String _
                                                         , Optional ByVal Description As String _
                                                          , Optional ByVal Size As String _
                                                           , Optional ByVal Image As String _
                                                            , Optional ByVal ImageMso As String _
                                                             , Optional ByVal Supertip As String) As String

    '@Description("This will generate Ribbon Button XML")
    '@Dependency("GetKeyValueCodeIfValueIsNotVbNullString")
    '@ExampleCall :
    '@Date : 13 June 2023 09:17:45 AM
    '@PossibleError:

    Const ONE_SPACE As String = " "
    Dim XML As String
    XML = "<button" & ONE_SPACE
    XML = XML & GetKeyValueCodeIfValueIsNotVbNullString("id", ID)
    XML = XML & GetKeyValueCodeIfValueIsNotVbNullString("label", Label)
    XML = XML & GetKeyValueCodeIfValueIsNotVbNullString("onAction", OnAction)
    XML = XML & GetKeyValueCodeIfValueIsNotVbNullString("size", Size)
    XML = XML & GetKeyValueCodeIfValueIsNotVbNullString("image", Image)
    If Image = vbNullString Then
        XML = XML & GetKeyValueCodeIfValueIsNotVbNullString("imageMso", ImageMso)
    End If
    XML = XML & GetKeyValueCodeIfValueIsNotVbNullString("description", Description)
    XML = XML & GetKeyValueCodeIfValueIsNotVbNullString("supertip", Supertip)
    XML = XML & "/>"
    GetButtonRibonXML = XML

End Function

Private Function GetKeyValueCodeIfValueIsNotVbNullString(ByVal Key As String, ByVal Value As String) As String
    If Value <> vbNullString Then GetKeyValueCodeIfValueIsNotVbNullString = Key & "=" & """" & Value & """" & Space(1)
End Function

Private Function CorrectForNewLine(ByVal CellData As Variant) As String
    CorrectForNewLine = Replace(Replace(CellData, vbNewLine, "<br>"), Chr$(10), "<br>")
End Function

Public Function MakeValidSubOrFunctionName(ByVal FromText As String) As String
    
    Dim Result As String
    Result = ConvertSentenceToVariableName(FromText)
    MakeValidSubOrFunctionName = Result
    
End Function

Public Function GetMultipleValueFromUser() As Variant

    Dim AllValue As Collection
    Set AllValue = New Collection

    Dim UserInput As Variant
    UserInput = InputBox("Enter first value:")

    Do While UserInput <> vbNullString
        AllValue.Add UserInput
        UserInput = InputBox("Enter next vale Or cancel to submit:")
    Loop

    If AllValue.Count > 0 Then
        GetMultipleValueFromUser = CollectionToArray(AllValue)
    End If

End Function

Public Function IsColumnExist(ByVal Table As ListObject, ByVal ColumnName As String) As Boolean

    If Table Is Nothing Then Exit Function
    If Table.ListColumns.Count = 0 Then Exit Function
    Dim CurrentColumn As ListColumn
    For Each CurrentColumn In Table.ListColumns
        If CurrentColumn.Name = ColumnName Then
            IsColumnExist = True
            Exit Function
        End If
    Next CurrentColumn

End Function

Public Function CreateURLByAppendingQueryParams(ByVal BaseURL As String _
                                                , ByVal QueryParameters As Variant _
                                                 , ByVal QueryValues As Variant) As String

    '@Description("This will create url from query param by concatenating.")
    '@Dependency("No Dependency")
    '@ExampleCall :CreateURLByAppendingQueryParams("https://www.alphavantage.co/query" _
    '                                              , Array("function", "symbol", "interval", "slice", "apikey") _
    '                                              , Array("TIME_SERIES_INTRADAY", "AAPL", "15min", "year1month1", "Demo"))
    '                                              Output: https://www.alphavantage.co/query?function=TIME_SERIES_INTRADAY&symbol=AAPL&interval=15min&slice=year1month1&apikey=Demo
    '@Date : 21 June 2023 08:30:35 PM
    '@PossibleError:

    If IsArray(QueryParameters) Then
        Dim Counter As Long
        Dim ParamName As Variant
        BaseURL = BaseURL & "?"
        For Each ParamName In QueryParameters
            BaseURL = BaseURL & ParamName & "=" & QueryValues(Counter) & "&"
            Counter = Counter + 1
        Next ParamName
        BaseURL = Left$(BaseURL, Len(BaseURL) - 1)
    Else
        BaseURL = BaseURL & "?" & QueryParameters & "=" & QueryValues
    End If
    CreateURLByAppendingQueryParams = BaseURL

End Function

Public Function FindOldestModifiedFileFullPathFromFolder(ByVal FolderPath As String) As String
    FindOldestModifiedFileFullPathFromFolder = GetLatestOrOldestModifiedFileFullPathFromFolder(FolderPath, False)
End Function

Public Function FindLatestModifiedFileFullPathFromFolder(ByVal FolderPath As String) As String
    FindLatestModifiedFileFullPathFromFolder = GetLatestOrOldestModifiedFileFullPathFromFolder(FolderPath, True)
End Function

Private Function GetLatestOrOldestModifiedFileFullPathFromFolder(ByVal FolderPath As String _
                                                                 , Optional ByVal IsLatest As Boolean = True) As String

    Dim FileManager As Object
    Set FileManager = CreateObject("Scripting.FileSystemObject")

    If Not IsFolderExist(FolderPath) Then
        MsgBox "Folder doesn't exists."
        Exit Function
    End If

    Dim NewestDate As Date
    NewestDate = DateSerial(1900, 1, 1)

    Dim OldestDate As Date
    OldestDate = DateSerial(9999, 12, 31)

    Dim FilePath As String

    Dim File As Object
    For Each File In FileManager.GetFolder(FolderPath).Files

        If IsLatest Then
            If File.DateLastModified > NewestDate Then
                FilePath = File.Path
                NewestDate = File.DateLastModified
            End If
        Else
            If File.DateLastModified < OldestDate Then
                FilePath = File.Path
                OldestDate = File.DateLastModified
            End If
        End If

    Next File

    GetLatestOrOldestModifiedFileFullPathFromFolder = FilePath

End Function

Public Function MakeValidSheetName(ByVal SheetName As String) As String

    If SheetName = vbNullString Then SheetName = "_Blank"

    Dim ValidSheetName As String
    Dim Index As Long
    Dim CurrentCharacter As String
    For Index = 1 To Len(SheetName)

        CurrentCharacter = Mid$(SheetName, Index, 1)
        Select Case CurrentCharacter
            Case "[", "]", "*", "/", "\", "?", ":"
                ValidSheetName = ValidSheetName
            Case Else
                ValidSheetName = ValidSheetName & CurrentCharacter
        End Select

    Next Index

    Const MAX_SHEET_NAME_LENGTH As Long = 31
    If Len(ValidSheetName) > MAX_SHEET_NAME_LENGTH Then
        ValidSheetName = Left$(ValidSheetName, MAX_SHEET_NAME_LENGTH)
    End If

    MakeValidSheetName = ValidSheetName

End Function

Public Function GetDesktopFolderPath() As String

    '@Description("Get Desktop folder path from registry or using mac script. Both Mac and Windows compatible")
    '@Dependency("No Dependency")
    '@ExampleCall :GetDesktopFolderPath()
    '@Date : 02 November 2023 10:52:49 PM
    '@PossibleError:

    #If Mac Then
        GetDesktopFolderPath = GetSpecialFolderPathInMac("desktop folder")
    #Else
        GetDesktopFolderPath = GetValueFromRegistry(HKEY_CURRENT_USER, SHELL_FOLDER_KEY, "Desktop")
    #End If

End Function

Public Function GetDocumentsFolderPath() As String

    '@Description("Get Documents folder path from registry or using mac script. Both Mac and Windows compatible")
    '@Dependency("No Dependency")
    '@ExampleCall :GetDocumentsFolderPath()
    '@Date : 02 November 2023 10:52:49 PM
    '@PossibleError:

    #If Mac Then
        GetDocumentsFolderPath = GetSpecialFolderPathInMac("documents folder")
    #Else
        GetDocumentsFolderPath = GetValueFromRegistry(HKEY_CURRENT_USER, SHELL_FOLDER_KEY, "Personal")
    #End If

End Function

Public Function GetDownloadsFolderPath() As String

    '@Description("Get Downloads folder path from registry or using mac script. Both Mac and Windows compatible")
    '@Dependency("No Dependency")
    '@ExampleCall :GetDownloadsFolderPath()
    '@Date : 02 November 2023 10:52:49 PM
    '@PossibleError:

    #If Mac Then
        GetDownloadsFolderPath = GetSpecialFolderPathInMac("downloads folder")
    #Else
        Const DOWNLOADS_FOLDER_VALUE_NAME As String = "{374DE290-123F-4565-9164-39C4925E467B}"
        GetDownloadsFolderPath = GetValueFromRegistry(HKEY_CURRENT_USER, SHELL_FOLDER_KEY, DOWNLOADS_FOLDER_VALUE_NAME)
    #End If

End Function

Public Function GetSpecialFolderPathInMac(ByVal NameFolder As String) As String

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

Public Function GetValueFromRegistry(ByVal BaseKey As BASE_KEY _
                                     , ByVal KeyName As String _
                                      , ByVal ValueName As String) As String

    On Error GoTo HandleError
    Const KEY_ALL_ACCESS As Long = &H3F
    Dim ReturnValueOfDLL  As Long
    Dim KeyHandle As Long
    ReturnValueOfDLL = RegOpenKeyEx(hKey:=BaseKey, lpSubKey:=KeyName _
                                                              , ulOptions:=0& _
                                                                            , samDesired:=KEY_ALL_ACCESS, phkResult:=KeyHandle)

    If KeyHandle = 0 Then
        GetValueFromRegistry = vbNullString
        Exit Function
    End If

    Const MAX_DATA_BUFFER_SIZE As Long = 1024
    Dim StringData As String
    StringData = String$(MAX_DATA_BUFFER_SIZE, vbNullChar)
    Dim LenStringData As Long
    LenStringData = Len(StringData)
    Const REG_SZ As Long = 1                     ' String


    ReturnValueOfDLL = RegQueryValueExStr(hKey:=KeyHandle, lpValueName:=ValueName, lpReserved:=0&, _
                                          lpType:=REG_SZ, szData:=StringData, lpcbData:=LenStringData)

    Const ERROR_SUCCESS As Long = 0&
    If ReturnValueOfDLL <> ERROR_SUCCESS Then
        RegCloseKey KeyHandle
        Exit Function
    End If

    Dim PositionOfNull As Long
    PositionOfNull = InStr(1, StringData, vbNullChar, vbTextCompare)

    If PositionOfNull <> 0 Then
        StringData = Left$(StringData, PositionOfNull - 1)
    End If

    GetValueFromRegistry = StringData
    Exit Function

HandleError:
    GetValueFromRegistry = vbNullString

End Function

Public Function GetFilteredRangeData(ByVal UnFilteredRange As Range) As Variant

    ' Extract data from visible range only. For example if we have a range
    ' like $A$1:$HG$902 but only $A$1:$D$1,$F$1:$HG$1,$A$694:$D$902,$F$694:$HG$902 these are visible
    ' Then this function will bring data from only the visible part. It can handle both hidden row and col


    If UnFilteredRange Is Nothing Then
        GetFilteredRangeData = vbEmpty
        Exit Function
    End If

    Dim DataRange As Range
    Set DataRange = UnFilteredRange.SpecialCells(xlCellTypeVisible)
    If UnFilteredRange.Address = DataRange.Address Then
        GetFilteredRangeData = UnFilteredRange.Value
        Exit Function
    End If

    Dim FirstAreaStartRow As Long
    FirstAreaStartRow = DataRange.Areas(1).Row
    Dim FirstAreaStartCol As Long
    FirstAreaStartCol = DataRange.Areas(1).Column

    Dim AreaCountInFirstRow As Long
    Dim AreaCountInFirstCol As Long
    Dim OutputArrayColCount As Long
    Dim OutputArrayRowCount As Long

    Dim CurrentRange As Range
    For Each CurrentRange In DataRange.Areas

        If CurrentRange.Row = FirstAreaStartRow Then
            AreaCountInFirstRow = AreaCountInFirstRow + 1
            OutputArrayColCount = OutputArrayColCount + CurrentRange.Columns.Count
        End If

        If CurrentRange.Column = FirstAreaStartCol Then
            AreaCountInFirstCol = AreaCountInFirstCol + 1
            OutputArrayRowCount = OutputArrayRowCount + CurrentRange.Rows.Count
        End If

    Next CurrentRange

    Dim Output As Variant
    ReDim Output(1 To OutputArrayRowCount, 1 To OutputArrayColCount)

    Dim RowIndex As Long
    Dim ColIndex As Long

    Dim i As Long, j As Long
    For i = 1 To AreaCountInFirstCol

        For j = 1 To AreaCountInFirstRow

            Dim CurrentArea As Variant
            CurrentArea = DataRange.Areas((i - 1) * AreaCountInFirstRow + j).Value
            Dim AreaRowIndex As Long
            Dim AreaColIndex As Long
            For AreaColIndex = LBound(CurrentArea, 2) To UBound(CurrentArea, 2)
                ColIndex = ColIndex + 1
                For AreaRowIndex = LBound(CurrentArea, 1) To UBound(CurrentArea, 1)
                    Output(RowIndex + AreaRowIndex - LBound(CurrentArea, 1) + 1, ColIndex) = CurrentArea(AreaRowIndex, AreaColIndex)
                Next AreaRowIndex
            Next AreaColIndex
        Next j
        RowIndex = RowIndex + DataRange.Areas((i - 1) * AreaCountInFirstRow + 1).Rows.Count
        ColIndex = 0
    Next i

    GetFilteredRangeData = Output

End Function

Private Function CollectionToArray(ByVal GivenCollection As Collection) As Variant

    '@Description("This will return the item of collection as variant array.")
    '@Dependency("No Dependency")
    '@ExampleCall : CollectionToArray(InputArray)
    '@Date : 14 October 2021 06:56:54 PM

    Dim Result() As Variant
    ReDim Result(1 To GivenCollection.Count, 1 To 1)
    Dim CurrentElement As Variant
    Dim CurrentIndex As Long
    For Each CurrentElement In GivenCollection
        CurrentIndex = CurrentIndex + 1
        If IsObject(CurrentElement) Then
            Set Result(CurrentIndex, 1) = CurrentElement
        Else
            Result(CurrentIndex, 1) = CurrentElement
        End If
    Next CurrentElement
    CollectionToArray = Result

End Function

Public Function GetColHeaderVsIndexMap(ByVal Data As Variant) As Object

    Dim Map As Object
    Set Map = CreateObject("Scripting.Dictionary")

    Dim FirstRowIndex As Long
    FirstRowIndex = LBound(Data, 1)
    Dim ColumnIndex As Long
    For ColumnIndex = LBound(Data, 2) To UBound(Data, 2)
        Map.Add Data(FirstRowIndex, ColumnIndex), ColumnIndex
    Next ColumnIndex

    Set GetColHeaderVsIndexMap = Map

End Function

Public Function ReadFromWorkbook(ByVal FilePath As String _
                                 , ByVal SheetName As String _
                                  , Optional ByVal FromRange As String) As Variant
    On Error GoTo HandleError
    Dim Connection As Object
    Set Connection = CreateObject("ADODB.Connection")
    With Connection
        .ConnectionString = _
                          "Provider=Microsoft.ACE.OLEDB.12.0;" & _
                          "Data Source=" & FilePath & ";" & _
                          "Extended Properties='Excel 12.0 Xml;HDR=YES;IMEX=1';"
        .Open
    End With

    Dim RecordSet As Object
    Set RecordSet = CreateObject("ADODB.RecordSet")
    With RecordSet
        .ActiveConnection = Connection
        .CursorType = 3
        If FromRange = vbNullString Then
            .Source = "SELECT * FROM [" & SheetName & "$]"
        Else
            .Source = "SELECT * FROM [" & SheetName & "$" & FromRange & "]"
        End If
        .Open
    End With

    ' We may have more than 255 columns. so try that.
    Const MAX_COLS_OF_ADDODB As Long = 255
    If RecordSet.Fields.Count = MAX_COLS_OF_ADDODB Then
        RecordSet.Close
        ReadFromWorkbook = ReadFromWorkbookIfMoreThan255Cols(SheetName, RecordSet, FromRange)
    Else
        ReadFromWorkbook = RecordSetTo2DArray(RecordSet)
    End If

ReleaseObjects:

    On Error Resume Next
    RecordSet.Close
    Connection.Close
    On Error GoTo 0

    Exit Function

HandleError:
    ReadFromWorkbook = vbEmpty
    GoTo ReleaseObjects

End Function

Private Function ReadFromWorkbookIfMoreThan255Cols(ByVal SheetName As String _
                                                   , ByVal RecordSet As Object _
                                                    , Optional ByVal FromRange As String) As Variant

    Dim StartColIndex As Long
    Dim EndColIndex As Long

    If FromRange = vbNullString Then
        StartColIndex = 1
        EndColIndex = ActiveSheet.Columns.Count
    Else
        StartColIndex = Range(FromRange).Column
        EndColIndex = Range(FromRange).Column + Range(FromRange).Columns.Count
    End If

    ' This is the limit for excel provider (ADODB)
    Const MAX_COLS_PER_GROUP As Long = 250

    Dim TotalGroup As Long
    If (EndColIndex - StartColIndex + 1) Mod MAX_COLS_PER_GROUP = 0 Then
        TotalGroup = (EndColIndex - StartColIndex + 1) / MAX_COLS_PER_GROUP
    Else
        TotalGroup = Int((EndColIndex - StartColIndex + 1) / MAX_COLS_PER_GROUP) + 1
    End If

    Dim Result As Variant

    Dim Counter As Long
    For Counter = 1 To TotalGroup

        Dim GroupRangeAddress As String
        GroupRangeAddress = Range(Cells(1, (Counter - 1) * MAX_COLS_PER_GROUP + 1) _
                                  , Cells(1, Counter * MAX_COLS_PER_GROUP)).EntireColumn.Address(False, False)
        Dim SQL As String
        SQL = "SELECT * FROM [" & SheetName & "$" & GroupRangeAddress & "]"
        RecordSet.Source = SQL
        '        Debug.Print SQL
        RecordSet.Open
        DoEvents
        Dim CurrentGroupData As Variant
        CurrentGroupData = RecordSetTo2DArray(RecordSet)
        If Counter = 1 Then
            Result = CurrentGroupData
        Else
            Result = HStack(Result, CurrentGroupData)
            DoEvents
        End If

        If RecordSet.Fields.Count < MAX_COLS_PER_GROUP Then
            RecordSet.Close
            Exit For
        End If
        RecordSet.Close

    Next Counter

    ReadFromWorkbookIfMoreThan255Cols = Result

End Function

Private Function RecordSetTo2DArray(ByVal GivenRecordSet As Object) As Variant

    '@Description("Convert a Recordset to Array")
    '@Dependency("No Dependency")
    '@ExampleCall : RecordSetToArray(GivenRecordSet)
    '@Date : 18 November 2022 09:10:01 PM
    '@PossibleError :


    'Some Guard Clause
    If GivenRecordSet Is Nothing Then
        Exit Function
    ElseIf GivenRecordSet.RecordCount = 0 Then
        Exit Function
    End If

    Dim Result As Variant
    Dim RecordCount As Long
    RecordCount = CLng(GivenRecordSet.RecordCount)
    ReDim Result(1 To RecordCount + 1, 1 To GivenRecordSet.Fields.Count)
    Dim ColIndex As Long

    'Fill Header
    For ColIndex = 1 To GivenRecordSet.Fields.Count
        Result(1, ColIndex) = GivenRecordSet.Fields(ColIndex - 1).Name
    Next ColIndex
    Dim RowIndex As Long
    RowIndex = 1

    'Fill with Data
    GivenRecordSet.MoveFirst
    Do While Not GivenRecordSet.EOF
        RowIndex = RowIndex + 1
        For ColIndex = 1 To GivenRecordSet.Fields.Count
            Result(RowIndex, ColIndex) = GivenRecordSet.Fields(Result(1, ColIndex)).Value
        Next ColIndex                            '
        GivenRecordSet.MoveNext
    Loop

    RecordSetTo2DArray = Result

End Function

Public Function GetLocalPathFromOneDrivePath(ByVal Path As String) As String

    ' Get local computer location path of a one drive file path.
    ' Sample Input:  https://d.docs.live.net/6edd704b8f8c537b/Documents/Stocks/Improved Studies/Project Automate
    ' Sample Output: C:\Users\USER\OneDrive\Documents\Stocks\Improved Studies\Project Automate
    ' Worked for both folder and file path.
    ' Read info from registry.
    ' Tested for personal OneDrive.

    Const ONE_DRIVE_PATH_PREFIX As String = "https://d.docs.live.net"
    If InStr(1, Path, ONE_DRIVE_PATH_PREFIX) = 0 Then
        GetLocalPathFromOneDrivePath = Path
        Exit Function
    End If

    Dim LocalRootOnedrivePath As String
    Dim URLNameSpace As String
    Dim CID As String

    Const REG_PATH As String = "SOFTWARE\SyncEngines\Providers\OneDrive\"

    Dim RegistryManager As Object
    Dim arSubKeys As Variant

    Set RegistryManager = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")

    With RegistryManager

        .EnumKey &H80000001, REG_PATH, arSubKeys

        Dim Index As Long
        For Index = 0 To UBound(arSubKeys)

            .GetStringValue &H80000001, REG_PATH & arSubKeys(Index), "MountPoint", LocalRootOnedrivePath
            .GetStringValue &H80000001, REG_PATH & arSubKeys(Index), "UrlNamespace", URLNameSpace
            .GetStringValue &H80000001, REG_PATH & arSubKeys(Index), "CID", CID

            CID = IIf(CID = vbNullString, vbNullString, "/" & CID)

            If InStr(1, Path, URLNameSpace & CID) = 1 Then
                GetLocalPathFromOneDrivePath = Replace(Replace(Path, URLNameSpace & CID, LocalRootOnedrivePath), "/", "\")
                Exit Function
            End If

        Next Index

    End With

End Function

Public Function GetOneDriveRootFolderPath() As String

    Dim LocalRootOnedrivePath As String
    Const REG_PATH As String = "SOFTWARE\SyncEngines\Providers\OneDrive\"

    Dim RegistryManager As Object
    Dim arSubKeys As Variant

    Set RegistryManager = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")

    With RegistryManager

        .EnumKey &H80000001, REG_PATH, arSubKeys

        Dim Index As Long
        For Index = 0 To UBound(arSubKeys)

            .GetStringValue &H80000001, REG_PATH & arSubKeys(Index), "MountPoint", LocalRootOnedrivePath
            If LocalRootOnedrivePath <> vbNullString Then Exit For
        Next Index

    End With

    GetOneDriveRootFolderPath = LocalRootOnedrivePath

End Function

Public Function GetFolderDetails(ByVal ParentFolderPath As String _
                                 , ByVal IsRecursively As Boolean _
                                  , Optional ByVal FolderFilter As FolderItemFilter = VISIBLE_AND_HIDDEN _
                                   , Optional ByVal IgnoreFolderFilterOnRootFolder As Boolean = True _
                                    , Optional ByVal ShowStatusBarMessage As Boolean = True _
                                     , Optional ByVal IgnoreFoldersName As Collection) As Variant

    If Not IsFolderExist(ParentFolderPath) Then Exit Function

    Dim AllValidFoldersData As Collection
    Set AllValidFoldersData = New Collection

    If ShowStatusBarMessage Then
        Application.StatusBar = "Folders data are being Collected. Please wait..."
    End If

    ScanAndRetriveFoldersData AllValidFoldersData, ParentFolderPath, IsRecursively _
                                                                    , FolderFilter _
                                                                     , IgnoreFolderFilterOnRootFolder _
                                                                      , ShowStatusBarMessage, IgnoreFoldersName

    If ShowStatusBarMessage Then Application.StatusBar = False

    Dim Result As Variant
    If AllValidFoldersData.Count = 0 Then Exit Function

    Dim NonBlankPropertiesIndexAndName As Object
    Set NonBlankPropertiesIndexAndName = GetAllPropertiesName(AllValidFoldersData)

    Dim PropertiesCount As Long
    PropertiesCount = UBound(NonBlankPropertiesIndexAndName.Keys) - LBound(NonBlankPropertiesIndexAndName.Keys) + 1
    ReDim Result(1 To AllValidFoldersData.Count + 1, 1 To PropertiesCount) As Variant

    Dim ColIndex As Long
    Dim Key As Variant
    For Each Key In NonBlankPropertiesIndexAndName.Keys
        Result(1, NonBlankPropertiesIndexAndName.Item(Key)) = Key
    Next Key

    Dim RowIndex As Long
    For RowIndex = 1 To AllValidFoldersData.Count
        Dim Properties As Object
        Set Properties = AllValidFoldersData.Item(RowIndex)
        Dim CurrentProperty As Variant
        For Each CurrentProperty In Properties.Keys
            ColIndex = NonBlankPropertiesIndexAndName.Item(CurrentProperty)
            Result(RowIndex + 1, ColIndex) = Properties.Item(CurrentProperty)
        Next CurrentProperty

    Next RowIndex

    Dim ImportantCols(1 To 7) As Variant
    ImportantCols(1) = "Path"
    ImportantCols(2) = "Name"
    ImportantCols(3) = "Date created"
    ImportantCols(4) = "Date modified"
    ImportantCols(5) = "Date accessed"
    ImportantCols(6) = "Folder name"
    ImportantCols(7) = "Folder path"
    Result = MoveColsToStart(RemoveBlankColumns(Result, True), ImportantCols, True, True)

    GetFolderDetails = Result

End Function

Public Function GetFilesDetails(ByVal ParentFolderPath As String _
                                , ByVal IsRecursively As Boolean _
                                 , Optional ByVal FolderFilter As FolderItemFilter = VISIBLE_AND_HIDDEN _
                                  , Optional ByVal FileFilter As FolderItemFilter = VISIBLE_AND_HIDDEN _
                                   , Optional ByVal IgnoreFolderFilterOnRootFolder As Boolean = True _
                                    , Optional ByVal ShowStatusBarMessage As Boolean = True _
                                    , Optional ByVal IgnoreFoldersName As Collection) As Variant

    If Not IsFolderExist(ParentFolderPath) Then
        MsgBox "Folder doesn't Exist.", , "Folder is Missing"
        Exit Function
    End If

    Dim AllValidFilesData As Collection
    Set AllValidFilesData = New Collection

    If ShowStatusBarMessage Then
        Application.StatusBar = "Files data are being Collected. Please wait..."
    End If
    ScanAndRetriveFilesData AllValidFilesData, ParentFolderPath, IsRecursively _
                                                                , FolderFilter _
                                                                 , FileFilter _
                                                                  , IgnoreFolderFilterOnRootFolder _
                                                                   , ShowStatusBarMessage, IgnoreFoldersName
    If ShowStatusBarMessage Then Application.StatusBar = False

    Dim Result As Variant
    If AllValidFilesData.Count = 0 Then
        Result = vbEmpty
    Else

        Dim NonBlankPropertiesIndexAndName As Object
        Set NonBlankPropertiesIndexAndName = GetAllPropertiesName(AllValidFilesData)

        Dim PropertiesCount As Long
        PropertiesCount = UBound(NonBlankPropertiesIndexAndName.Keys) - LBound(NonBlankPropertiesIndexAndName.Keys) + 1
        ReDim Result(1 To AllValidFilesData.Count + 1, 1 To PropertiesCount) As Variant

        Dim ColIndex As Long
        Dim Key As Variant
        For Each Key In NonBlankPropertiesIndexAndName.Keys
            Result(1, NonBlankPropertiesIndexAndName.Item(Key)) = Key
        Next Key

        Dim RowIndex As Long
        For RowIndex = 1 To AllValidFilesData.Count
            Dim Properties As Object
            Set Properties = AllValidFilesData.Item(RowIndex)
            Dim CurrentProperty As Variant
            For Each CurrentProperty In Properties.Keys
                ColIndex = NonBlankPropertiesIndexAndName.Item(CurrentProperty)
                Result(RowIndex + 1, ColIndex) = Properties.Item(CurrentProperty)
            Next CurrentProperty

        Next RowIndex

        Dim ImportantCols(1 To 14) As Variant
        ImportantCols(1) = "Path"
        ImportantCols(2) = "Name"
        ImportantCols(3) = "File extension"
        ImportantCols(4) = "Size"
        ImportantCols(5) = "Kind"
        ImportantCols(6) = "Item type"
        ImportantCols(7) = "Authors"
        ImportantCols(8) = "Date created"
        ImportantCols(9) = "Date modified"
        ImportantCols(10) = "Date accessed"
        ImportantCols(11) = "Date last saved"
        ImportantCols(12) = "Last printed"
        ImportantCols(13) = "Folder name"
        ImportantCols(14) = "Folder path"
        Result = MoveColsToStart(RemoveBlankColumns(Result, True), ImportantCols, True, True)

    End If

    GetFilesDetails = Result

End Function

Private Function GetAllPropertiesName(ByVal AllValidItemData As Variant) As Variant

    Dim NonBlankPropertiesName As Object
    Set NonBlankPropertiesName = CreateObject("Scripting.Dictionary")

    Dim Counter As Long
    Counter = 1
    Dim PropNameVsValue As Object
    For Each PropNameVsValue In AllValidItemData

        Dim Key As Variant
        For Each Key In PropNameVsValue.Keys
            If Not NonBlankPropertiesName.Exists(Key) Then
                NonBlankPropertiesName.Add Key, Counter
                Counter = Counter + 1
            End If
        Next Key

    Next PropNameVsValue

    Set GetAllPropertiesName = NonBlankPropertiesName

End Function

Private Sub ScanAndRetriveFoldersData(ByRef AllValidFoldersData As Collection _
                                      , ByVal ParentFolderPath As String _
                                       , ByVal IsRecursively As Boolean _
                                        , Optional ByVal FolderFilter As FolderItemFilter = VISIBLE_AND_HIDDEN _
                                         , Optional ByVal IgnoreFolderFilterOnRootFolder As Boolean = True _
                                          , Optional ByVal ShowStatusBarMessage As Boolean = True _
                                           , Optional ByVal IgnoreFoldersName As Collection)

    If Right$(ParentFolderPath, Len(Application.PathSeparator)) <> Application.PathSeparator Then
        ParentFolderPath = ParentFolderPath & Application.PathSeparator
    End If

    ' If folder is hidden but we are looking for only visible then exit or vice versa.
    If Not IgnoreFolderFilterOnRootFolder Then
        If IsHiddenFolderOrFile(ParentFolderPath) And FolderFilter = VISIBLE_ONLY Then
            Exit Sub
        ElseIf Not IsHiddenFolderOrFile(ParentFolderPath) And FolderFilter = HIDDEN_ONLY Then
            Exit Sub
        End If
    End If

    Static FSO As Object
    Static ShellApp As Object
    If FSO Is Nothing Then
        Set FSO = CreateObject("Scripting.FileSystemObject")
        Set ShellApp = CreateObject("Shell.Application")
    End If

    Dim ParentFolder As Object
    Set ParentFolder = FSO.GetFolder(ParentFolderPath)


    Dim ParentFolderInShell As Object
    Set ParentFolderInShell = ShellApp.Namespace("" & ParentFolderPath & "")

    Dim IsScanSubFolder As Boolean
    Dim CurrentSubFolder As Object
    If IsRecursively Then

        For Each CurrentSubFolder In ParentFolder.SubFolders

            If IgnoreFoldersName Is Nothing Then
                IsScanSubFolder = True
            ElseIf IsExistInCollection(IgnoreFoldersName, CurrentSubFolder.Name) Then
                IsScanSubFolder = False
            Else
                IsScanSubFolder = True
            End If

            If IsScanSubFolder Then
                ScanAndRetriveFoldersData AllValidFoldersData, CurrentSubFolder.Path, IsRecursively _
                                                                                     , FolderFilter _
                                                                                      , False, ShowStatusBarMessage
            End If

        Next CurrentSubFolder

    Else

        ' If not recursive then only process the first depth folders.
        For Each CurrentSubFolder In ParentFolder.SubFolders

            If IgnoreFoldersName Is Nothing Then
                IsScanSubFolder = True
            ElseIf IsExistInCollection(IgnoreFoldersName, CurrentSubFolder.Name) Then
                IsScanSubFolder = False
            Else
                IsScanSubFolder = True
            End If

            If IsScanSubFolder Then
                AllValidFoldersData.Add GetAllPropertiesForCurrentItem(ParentFolderInShell.ParseName(CurrentSubFolder.Name), ShowStatusBarMessage)
            End If

        Next CurrentSubFolder

    End If

End Sub

Private Sub ScanAndRetriveFilesData(ByRef AllValidFilesData As Collection _
                                    , ByVal ParentFolderPath As String _
                                     , ByVal IsRecursively As Boolean _
                                      , Optional ByVal FolderFilter As FolderItemFilter = VISIBLE_AND_HIDDEN _
                                       , Optional ByVal FileFilter As FolderItemFilter = VISIBLE_AND_HIDDEN _
                                        , Optional ByVal IgnoreFolderFilterOnRootFolder As Boolean = True _
                                         , Optional ByVal ShowStatusBarMessage As Boolean = True _
                                          , Optional ByVal IgnoreFoldersName As Collection)

    If Right$(ParentFolderPath, Len(Application.PathSeparator)) <> Application.PathSeparator Then
        ParentFolderPath = ParentFolderPath & Application.PathSeparator
    End If

    ' If folder is hidden but we are looking for only visible then exit or vice versa.
    If Not IgnoreFolderFilterOnRootFolder Then
        If IsHiddenFolderOrFile(ParentFolderPath) And FolderFilter = VISIBLE_ONLY Then
            Exit Sub
        ElseIf Not IsHiddenFolderOrFile(ParentFolderPath) And FolderFilter = HIDDEN_ONLY Then
            Exit Sub
        End If
    End If

    Static FSO As Object
    Static ShellApp As Object
    If FSO Is Nothing Then
        Set FSO = CreateObject("Scripting.FileSystemObject")
        Set ShellApp = CreateObject("Shell.Application")
    End If

    Dim ParentFolder As Object
    Set ParentFolder = FSO.GetFolder(ParentFolderPath)


    Dim ParentFolderInShell As Object
    Set ParentFolderInShell = ShellApp.Namespace("" & ParentFolderPath & "")

    ' Loop through all the files in that folder.

    Dim FileInShell As Object

    Dim CurrentFile As Object
    For Each CurrentFile In ParentFolder.Files

        Dim IsFileHidden As Boolean
        IsFileHidden = IsHiddenFolderOrFile(CurrentFile.Path)
        Dim IsValidFileToExtractDetails As Boolean
        IsValidFileToExtractDetails = Not ( _
                                      (IsFileHidden And FileFilter = VISIBLE_ONLY) _
                                      Or (Not IsFileHidden And FileFilter = HIDDEN_ONLY) _
                                      )

        If IsValidFileToExtractDetails Then
            Set FileInShell = ParentFolderInShell.ParseName(CurrentFile.Name)
            Dim PropertiesValue As Object
            Set PropertiesValue = GetAllPropertiesForCurrentItem(FileInShell, ShowStatusBarMessage)
            AllValidFilesData.Add PropertiesValue
            DoEvents
        End If

    Next CurrentFile

    If Not IsRecursively Then Exit Sub

    Dim CurrentSubFolder As Object
    For Each CurrentSubFolder In ParentFolder.SubFolders

        Dim IsScanSubFolder As Boolean
        If IgnoreFoldersName Is Nothing Then
            IsScanSubFolder = True
        ElseIf IsExistInCollection(IgnoreFoldersName, CurrentSubFolder.Name) Then
            IsScanSubFolder = False
        Else
            IsScanSubFolder = True
        End If

        If IsScanSubFolder Then
            ScanAndRetriveFilesData AllValidFilesData, CurrentSubFolder.Path _
                                                      , IsRecursively _
                                                       , FolderFilter _
                                                        , FileFilter, False, ShowStatusBarMessage
        End If

    Next CurrentSubFolder

End Sub

Public Function IsHiddenFolderOrFile(ByVal FullFilePath As String) As Boolean

    ' Find if a specific file or folder is hidden or not.

    Dim FileAttrIntegerValue As Integer
    FileAttrIntegerValue = GetAttr(FullFilePath)
    IsHiddenFolderOrFile = (FileAttrIntegerValue And vbHidden) <> 0

End Function

Public Function GetAllPropertiesForCurrentItem(ByVal CurrentItem As Object _
                                               , ByVal ShowStatusBarMessage As Boolean) As Object

    Dim PropertyNameVsValueMap As Object
    Set PropertyNameVsValueMap = CreateObject("Scripting.Dictionary")

    Dim ParentItem As Object
    Set ParentItem = CurrentItem.Parent

    If CurrentItem.IsFolder And ShowStatusBarMessage Then
        Application.StatusBar = "Collecting details for: " & CurrentItem.Path
    ElseIf ShowStatusBarMessage Then
        Const STATUS_BAR_MAX_LEN_LIMIT As Long = 255
        Application.StatusBar = Right$("Collecting details for: " & CurrentItem.Path _
                                      & Application.PathSeparator & CurrentItem.Name _
                                      , STATUS_BAR_MAX_LEN_LIMIT)
    End If

    Const TOTAL_PROPERTIES_COUNT As Long = 321
    Dim Index As Long
    For Index = 0 To TOTAL_PROPERTIES_COUNT - 1
        Dim PropertyName As String
        PropertyName = ParentItem.GetDetailsOf(Null, Index)
        If PropertyName <> vbNullString Then
            If PropertyNameVsValueMap.Exists(PropertyName) Then
                If PropertyNameVsValueMap.Item(PropertyName) = vbNullString Then
                    PropertyNameVsValueMap.Item(PropertyName) = ParentItem.GetDetailsOf(CurrentItem, Index)
                End If
            Else
                PropertyNameVsValueMap.Add PropertyName, ParentItem.GetDetailsOf(CurrentItem, Index)
            End If
        End If
    Next Index

    Set GetAllPropertiesForCurrentItem = PropertyNameVsValueMap

End Function

Private Function RemoveBlankColumns(ByVal InputArr As Variant, Optional ByVal IsUsingHeader As Boolean = False) As Variant

    If Not Is2DArray(InputArr) Then
        Err.Raise 13, "Type mismatch.", "InputArr need to be a 2D array."
    End If

    Dim FirstRowIndex As Long
    FirstRowIndex = LBound(InputArr, 1) + IIf(IsUsingHeader, 1, 0)

    Dim NonBlankColIndexes As Collection
    Set NonBlankColIndexes = New Collection

    Dim ColIndex As Long

    For ColIndex = LBound(InputArr, 2) To UBound(InputArr, 2)

        Dim Counter As Long
        For Counter = FirstRowIndex To UBound(InputArr, 1)
            If InputArr(Counter, ColIndex) <> vbNullString Then
                NonBlankColIndexes.Add ColIndex, CStr(ColIndex)
                Exit For
            End If
        Next Counter

    Next ColIndex

    RemoveBlankColumns = ChooseCols(InputArr, CollectionToVector(NonBlankColIndexes), False)

End Function

Private Function MoveColsToStart(ByVal InputArr As Variant, ByVal ColIndexOrHeader As Variant _
                                                           , ByVal IsUsingHeader As Boolean _
                                                            , ByVal IsIgnoreMissingCol As Boolean) As Variant

    If Not Is2DArray(InputArr) Then
        Err.Raise 13, "Type mismatch.", "InputArr need to be a 2D array."
    End If

    Dim ValidColIndexes As Collection
    Set ValidColIndexes = New Collection

    Dim IndexOrHeader As Variant
    For Each IndexOrHeader In ColIndexOrHeader
        Dim ColIndex As Long
        If IsUsingHeader Then
            ColIndex = FindColIndex(InputArr, IndexOrHeader)
            If ColIndex = -1 Then
                If Not IsIgnoreMissingCol Then
                    Err.Raise 13, "Column not found.", "Column " & IndexOrHeader & " not found."
                End If
            Else
                ValidColIndexes.Add ColIndex, CStr(ColIndex)
            End If
        Else
            ColIndex = IndexOrHeader
            If ColIndex < LBound(InputArr, 2) Or ColIndex > UBound(InputArr, 2) Then
                Err.Raise 13, "Type mismatch.", "ColIndex is out of bound."
            End If
            ValidColIndexes.Add ColIndex, CStr(ColIndex)
        End If
    Next IndexOrHeader

    Dim Counter As Long
    For Counter = LBound(InputArr, 2) To UBound(InputArr, 2)
        If Not IsExistInCollection(ValidColIndexes, CStr(Counter)) Then
            ValidColIndexes.Add Counter, CStr(Counter)
        End If
    Next Counter

    MoveColsToStart = ChooseCols(InputArr, CollectionToVector(ValidColIndexes), False)

End Function

Public Function IsVector(ByVal InputArr As Variant) As Boolean
    IsVector = (NumberOfArrayDimensions(InputArr) = 1)
End Function

Private Function Is2DArray(ByVal InputArray As Variant) As Boolean
    ' It just check if 2D array or not. It doesn't gurantee that in both dimension there will be more than one element.
    Is2DArray = (NumberOfArrayDimensions(InputArray) = 2)
End Function

Private Function NumberOfArrayDimensions(ByVal InputArray As Variant) As Byte

    ' This function returns the number of dimensions of an array. An unallocated dynamic array
    ' has 0 dimensions. This condition can also be tested with IsArrayEmpty.
    ' The output of this function is byte data type because VBA allow maximum of 60 dimension
    ' and Byte can hold upto 256. So byte
    Dim Ndx As Byte
    Dim Res As Long
    On Error Resume Next
    ' Loop, increasing the dimension index Ndx, until an error occurs.
    ' An error will occur when Ndx exceeds the number of dimension
    ' in the array. Return Ndx - 1.
    Do
        Ndx = Ndx + 1
        Res = UBound(InputArray, Ndx)
    Loop Until Err.Number <> 0
    On Error GoTo 0
    'Return the dimension..-1 because we are increasing the value of Ndx before error occured..
    NumberOfArrayDimensions = Ndx - 1

End Function

Private Function ChooseCols(ByVal InputArray As Variant _
                            , ByVal ColumnsToSelect As Variant _
                             , ByVal IsUsingHeader As Boolean) As Variant

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Choose specific columns from an array. You can choose columns using col header or using index
    ' If you choose option to use colindex then you have a option to choose from end using negative number.
    ' So if you want to choose last col then pass -1 if second last then -2 etc.
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    If Not IsArrayAllocated(InputArray) Then
        Err.Raise ERR_UN_ALLOCATED_ARRAY, "Invalid Input", MSG_UN_ALLOCATED_ARRAY
        Exit Function
    End If

    Dim ColIndexToSelect As Variant
    If IsUsingHeader Then
        ColIndexToSelect = FindAllColIndex(InputArray, ColumnsToSelect)
    Else
        ColIndexToSelect = ConvertToArrayIfNotArray(ColumnsToSelect, True)
    End If

    Dim Result As Variant
    ReDim Result(LBound(InputArray, 1) To UBound(InputArray, 1), LBound(ColIndexToSelect, 1) To UBound(ColIndexToSelect, 1))

    Dim CurrentRowIndex As Long
    For CurrentRowIndex = LBound(InputArray, 1) To UBound(InputArray, 1)

        Dim ColIndex As Long
        For ColIndex = LBound(ColIndexToSelect, 1) To UBound(ColIndexToSelect, 1)
            Dim InputArrayColIndex As Long
            InputArrayColIndex = ColIndexToSelect(ColIndex, LBound(ColIndexToSelect, 2))

            ' If using Index and if pass negative number then extract from end
            If InputArrayColIndex < LBound(InputArray, 2) And Not IsUsingHeader Then
                InputArrayColIndex = UBound(InputArray, 2) + InputArrayColIndex + 1
            End If

            Result(CurrentRowIndex, ColIndex) = InputArray(CurrentRowIndex _
                                                           , InputArrayColIndex)
        Next ColIndex

    Next CurrentRowIndex

    ChooseCols = Result

End Function

Private Function IsArrayAllocated(ByVal InputArray As Variant) As Boolean

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' IsArrayAllocated. Returns TRUE if the array is allocated (either a static array or a dynamic array that has been
    ' sized with Redim) or FALSE if the array is not allocated (a dynamic that has not yet
    ' been sized with Redim, or a dynamic array that has been Erased). Static arrays are always allocated.
    '
    ' The VBA IsArray function indicates whether a variable is an array, but it does not
    ' distinguish between allocated and unallocated arrays. It will return TRUE for both
    ' allocated and unallocated arrays. This function tests whether the array has actually
    ' been allocated.This function is just the reverse of IsArrayEmpty.
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Dim N As Long
    On Error Resume Next
    If Not IsArray(InputArray) Then
        IsArrayAllocated = False
        Exit Function
    End If

    ' Attempt to get the UBound of the array. If the array has not been allocated,
    ' an error will occur. Test Err.Number to see if an error occurred.
    N = UBound(InputArray, 1)
    If (Err.Number = 0) Then
        ''''''''''''''''''''''''''''''''''''''
        ' Under some circumstances, if an array
        ' is not allocated, Err.Number will be
        ' 0. To acccomodate this case, we test
        ' whether LBound <= Ubound. If this
        ' is True, the array is allocated. Otherwise,
        ' the array is not allocated.
        '''''''''''''''''''''''''''''''''''''''
        IsArrayAllocated = (LBound(InputArray) <= UBound(InputArray))
    Else
        IsArrayAllocated = False                 ' error. unallocated array
    End If
    On Error GoTo 0

End Function

Private Function FindAllColIndex(ByVal InputArray As Variant _
                                 , ByVal ColumnsHeader As Variant) As Variant

    Dim ColIndexToSelect As Variant
    ColIndexToSelect = ConvertToArrayIfNotArray(ColumnsHeader, True)

    Dim FirstColumnIndex  As Long
    FirstColumnIndex = LBound(ColIndexToSelect, 2)
    Dim CurrentRowIndex As Long
    For CurrentRowIndex = LBound(ColIndexToSelect, 1) To UBound(ColIndexToSelect, 1)
        Dim ColIndex As Long
        ColIndex = FindColIndex(InputArray, ColIndexToSelect(CurrentRowIndex, FirstColumnIndex))
        ColIndexToSelect(CurrentRowIndex, FirstColumnIndex) = ColIndex
    Next CurrentRowIndex
    FindAllColIndex = ColIndexToSelect

End Function

Private Function FindColIndex(ByVal InputArray As Variant, ByVal ColHeader As String) As Long

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '   Find Column index of a specific header from first row of the array.
    '   If not found or invalid input arguments then it will return -1
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    If IsArrayAllocated(InputArray) Then

        Dim FirstRowIndex As Long
        FirstRowIndex = LBound(InputArray, 1)
        Dim CurrentColumnIndex As Long
        For CurrentColumnIndex = LBound(InputArray, 2) To UBound(InputArray, 2)
            If InputArray(FirstRowIndex, CurrentColumnIndex) = ColHeader Then
                FindColIndex = CurrentColumnIndex
                Exit Function
            End If
        Next CurrentColumnIndex

    End If
    FindColIndex = -1

End Function

Private Function ConvertToArrayIfNotArray(ByVal InputArray As Variant _
, ByVal ConvertTo2D As Boolean) As Variant

    Dim Result As Variant
    If IsArrayAllocated(InputArray) Then
        If Not ConvertTo2D Then
            Result = InputArray
        ElseIf NumberOfArrayDimensions(InputArray) = 2 Then
            Result = InputArray
        ElseIf NumberOfArrayDimensions(InputArray) = 1 And ConvertTo2D Then
            ReDim Result(LBound(InputArray) To UBound(InputArray), 1 To 1)
            Dim CurrentIndex As Long
            For CurrentIndex = LBound(InputArray) To UBound(InputArray)
                Result(CurrentIndex, 1) = InputArray(CurrentIndex)
            Next CurrentIndex
        End If
    Else
        If ConvertTo2D Then
            ReDim Result(1 To 1, 1 To 1)
            Result(1, 1) = InputArray
        Else
            Result = Array(InputArray)
        End If
    End If
    ConvertToArrayIfNotArray = Result

End Function

Private Function CollectionToVector(ByVal InputCollection As Collection) As Variant

    If InputCollection.Count = 0 Then
        CollectionToVector = vbEmpty
        Exit Function
    End If

    Dim Result As Variant
    ReDim Result(1 To InputCollection.Count)

    Dim Index As Long
    For Index = 1 To InputCollection.Count
        Result(Index) = InputCollection.Item(Index)
    Next Index

    CollectionToVector = Result

End Function

Private Function HStack(ByVal FirstArray As Variant, ByVal SecondArray As Variant) As Variant

    Dim TotalRowCount As Long
    Dim TotalColCount As Long
    TotalColCount = NumberOfColumnIn2DArray(FirstArray) + NumberOfColumnIn2DArray(SecondArray)
    TotalRowCount = Max(NumberOfRowIn2DArray(FirstArray), NumberOfRowIn2DArray(SecondArray))
    FirstArray = Expand(FirstArray, TotalRowCount)
    SecondArray = Expand(SecondArray, TotalRowCount)

    Dim Result As Variant
    ReDim Result(1 To TotalRowCount, 1 To TotalColCount)

    Dim FirstArrayColCount As Long
    FirstArrayColCount = NumberOfColumnIn2DArray(FirstArray)

    Dim RowIndex As Long
    Dim ColIndex As Long
    For RowIndex = 1 To TotalRowCount

        For ColIndex = 1 To TotalColCount
            If ColIndex <= FirstArrayColCount Then
                Result(RowIndex, ColIndex) = FirstArray(LBound(FirstArray, 1) + RowIndex - 1 _
                                                        , LBound(FirstArray, 2) + ColIndex - 1)
            Else
                Result(RowIndex, ColIndex) = SecondArray(LBound(SecondArray, 1) + RowIndex - 1 _
                                                         , LBound(SecondArray, 2) + ColIndex - FirstArrayColCount - 1)
            End If
        Next ColIndex

    Next RowIndex
    HStack = Result

End Function

Public Function Max(ByVal X As Variant, ByVal Y As Variant) As Variant
    Max = IIf(X > Y, X, Y)
End Function

Private Function NumberOfColumnIn2DArray(ByVal GivenArray As Variant) As Long

    '@Description("This will calculate the number of column of a 2D Array")
    '@Dependency("No Dependency")
    '@ExampleCall : NumberOfColumnIn2DArray(GivenArray)
    '@Date : 26 December 2021 01:38:56 AM

    NumberOfColumnIn2DArray = UBound(GivenArray, 2) - LBound(GivenArray, 2) + 1

End Function

Private Function NumberOfRowIn2DArray(ByVal GivenArray As Variant) As Long

    '@Description("This will calculate the number of row of a 2D Array")
    '@Dependency("No Dependency")
    '@ExampleCall : NumberOfRowIn2DArray(GivenArray)
    '@Date : 26 December 2021 01:38:56 AM

    NumberOfRowIn2DArray = UBound(GivenArray, 1) - LBound(GivenArray, 1) + 1

End Function

Private Function Expand(ByVal InputArray As Variant _
                        , Optional ByVal NewRowCount As Long _
                         , Optional ByVal NewColCount As Long _
                          , Optional ByVal PadWith As Variant) As Variant

    If Not Is2DArray(InputArray) Then
        Err.Raise vbObjectError + 1, "Invalid Arguments", "Input array need to be 2D array."
    End If

    If NewRowCount = 0 Then NewRowCount = NumberOfRowIn2DArray(InputArray)
    If NewColCount = 0 Then NewColCount = NumberOfColumnIn2DArray(InputArray)
    If IsMissing(PadWith) Then PadWith = "#N/A"

    Dim OldRowCount As Long
    OldRowCount = NumberOfRowIn2DArray(InputArray)
    Dim OldColCount As Long
    OldColCount = NumberOfColumnIn2DArray(InputArray)

    Dim Result As Variant
    ReDim Result(1 To NewRowCount, 1 To NewColCount)
    Dim RowIndex As Long
    For RowIndex = 1 To NewRowCount

        Dim ColIndex As Long
        For ColIndex = 1 To NewColCount
            If RowIndex > OldRowCount Or ColIndex > OldColCount Then
                Result(RowIndex, ColIndex) = PadWith
            Else
                Result(RowIndex, ColIndex) = InputArray(LBound(InputArray, 1) + RowIndex - 1, LBound(InputArray, 2) + ColIndex - 1)
            End If
        Next ColIndex

    Next RowIndex
    Expand = Result

End Function

Public Function ReplaceNewlineWith(ByVal FromText As String, ByVal ReplaceWith As String) As String
    ReplaceNewlineWith = Replace(Replace(FromText, vbNewLine, ReplaceWith), Chr$(10), ReplaceWith)
End Function

Public Function GetQueryDataFromDataModel(ByVal QueryName As String _
                                          , ByVal FromBook As Workbook _
                                           , Optional ByVal IsRefresh As Boolean) As Variant

    ' Get Query data which you have loaded to data model into a 2d array.
    ' Use query name that you have used when developing the M code.
    ' No need to use Query - prefix.

    Dim CurrentModelTable As ModelTable
    Set CurrentModelTable = FromBook.Model.ModelTables.Item(QueryName)
    If CurrentModelTable Is Nothing Then Exit Function
    If IsRefresh Then
        CurrentModelTable.Refresh
        Set CurrentModelTable = FromBook.Model.ModelTables.Item(QueryName)
    End If

    Dim AdoConn As Object
    Set AdoConn = ThisWorkbook.Model.DataModelConnection.ModelConnection.ADOConnection

    Dim RecordSet As Object
    Set RecordSet = CreateObject("ADODB.RecordSet")
    RecordSet.Open "SELECT * From [$" & QueryName & "].[$" & QueryName & "]", AdoConn
    Dim Result As Variant
    Result = RecordSetTo2DArray(RecordSet)

    ' Clean up header row.
    RemoveColHeaderPrefixAndSuffix Result, "[$" & QueryName & "].[", "]"
    Const ROW_NUMBER_COL_HEADER As String = "__XL_RowNumber"
    Result = RemoveCols(Result, ROW_NUMBER_COL_HEADER, True)

    GetQueryDataFromDataModel = Result

End Function

Private Sub RemoveColHeaderPrefixAndSuffix(ByRef FromArr As Variant, ByVal Prefix As String, ByVal Suffix As String)

    Dim PrefixLength As Long
    PrefixLength = Len(Prefix)

    Dim FirstRowIndex As Long
    FirstRowIndex = LBound(FromArr, 1)

    Dim ColIndex As Long
    For ColIndex = LBound(FromArr, 2) To UBound(FromArr, 2)
        Dim Header As String
        Header = FromArr(FirstRowIndex, ColIndex)
        Header = Mid$(Header, PrefixLength + 1, Len(Header) - PrefixLength - Len(Suffix))
        FromArr(FirstRowIndex, ColIndex) = Header
    Next ColIndex

End Sub

Private Function RemoveCols(ByVal InputArray As Variant _
                            , ByVal ColumnsToRemove As Variant _
                             , ByVal IsUsingHeader As Boolean) As Variant

    If Not IsArrayAllocated(InputArray) Then
        Err.Raise ERR_UN_ALLOCATED_ARRAY, "Invalid Input", MSG_UN_ALLOCATED_ARRAY
        Exit Function
    End If

    Dim ColIndexToRemove As Variant
    If IsUsingHeader Then
        ColIndexToRemove = FindAllColIndex(InputArray, ColumnsToRemove)
    Else
        ColIndexToRemove = ConvertToArrayIfNotArray(ColumnsToRemove, True)
    End If
    Dim InvalidColIndex As Collection
    Set InvalidColIndex = ToCollection(ColIndexToRemove, LBound(ColIndexToRemove, 2), LBound(ColIndexToRemove, 2))

    Dim ValidColIndex As Collection
    Set ValidColIndex = New Collection

    Dim CurrentColumnIndex As Long
    For CurrentColumnIndex = LBound(InputArray, 2) To UBound(InputArray, 2)
        If Not IsExistInCollection(InvalidColIndex, CStr(CurrentColumnIndex)) Then
            ValidColIndex.Add CurrentColumnIndex, CStr(CurrentColumnIndex)
        End If
    Next CurrentColumnIndex

    If ValidColIndex.Count = 0 Then
        RemoveCols = vbEmpty
        Exit Function
    End If

    RemoveCols = ChooseCols(InputArray, FromCollection(ValidColIndex), False)

End Function

Private Function ToCollection(ByVal GivenArray As Variant, Optional ByVal KeyColumnIndex As Long = -1, _
                                  Optional ByVal ItemColumnIndex As Long = -1 _
                                  , Optional ByVal IsSuppressDuplicateError As Boolean = True) As Collection

    '@Description : By default it will not throw error for duplicate case
    '@FullyQualifiedCase: It will use given column indexes
    '@ExampleCall : Set ArrayToCollectionMapping = ArrayToCollection(SUTArray, 1, 2, IsSuppressDuplicateError:=False) >> It will throw error if any duplicate key is present.
    '@ExampleCall : Set ArrayToCollectionMapping = ArrayToCollection(SUTArray, 1, 2, IsSuppressDuplicateError:=True) >> It will skip duplicate item.

    '@DefaultValueCase : first column is the key and second column is the item
    '@ExampleCall : Set ArrayToCollectionMapping = ArrayToCollection(SUTArray, IsSuppressDuplicateError:=False) >> It will throw error if any duplicate key is present.
    '@ExampleCall : Set ArrayToCollectionMapping = ArrayToCollection(SUTArray, IsSuppressDuplicateError:=True) >> It will skip duplicate item.

    Dim FirstColumnIndex  As Long
    FirstColumnIndex = LBound(GivenArray, 2)

    If KeyColumnIndex = -1 Then KeyColumnIndex = FirstColumnIndex
    If ItemColumnIndex = -1 Then ItemColumnIndex = FirstColumnIndex + 1
    If IsSuppressDuplicateError Then On Error Resume Next

    Dim CurrentRowIndex As Long
    Dim KeyItemMapping As Collection
    Set KeyItemMapping = New Collection
    For CurrentRowIndex = LBound(GivenArray, 1) To UBound(GivenArray, 1)
        Dim Key As String
        Key = CStr(GivenArray(CurrentRowIndex, KeyColumnIndex))
        Dim Item As Variant
        Item = GivenArray(CurrentRowIndex, ItemColumnIndex)
        KeyItemMapping.Add Item, Key
    Next CurrentRowIndex
    If IsSuppressDuplicateError Then On Error GoTo 0

    Set ToCollection = KeyItemMapping

End Function

Private Function FromCollection(ByVal GivenCollection As Collection) As Variant

    '@Description("This will return the item of collection as variant array.")
    '@Dependency("No Dependency")
    '@ExampleCall : CollectionToArray(InputArray)
    '@Date : 14 October 2021 06:56:54 PM

    Dim Result() As Variant
    ReDim Result(1 To GivenCollection.Count, 1 To 1)
    Dim CurrentElement As Variant
    Dim CurrentIndex As Long
    For Each CurrentElement In GivenCollection
        CurrentIndex = CurrentIndex + 1
        Result(CurrentIndex, 1) = CurrentElement
    Next CurrentElement
    FromCollection = Result

End Function

Public Function GetRefersToRange(ByVal NamedRangeName As String _
                                 , Optional ByVal FromBook As Workbook) As Range


    If FromBook Is Nothing Then Set FromBook = ThisWorkbook
    Set GetRefersToRange = FromBook.Names(NamedRangeName).RefersToRange

End Function

Public Function GetRefersToRangeValue(ByVal NamedRangeName As String _
                                      , Optional ByVal FromBook As Workbook) As Variant


    If FromBook Is Nothing Then Set FromBook = ThisWorkbook
    GetRefersToRangeValue = FromBook.Names(NamedRangeName).RefersToRange.Value

End Function

Public Function GetNamedRangeValue(ByVal NamedRangeName As String _
                                   , Optional ByVal FromBook As Workbook) As Variant

    If FromBook Is Nothing Then Set FromBook = ThisWorkbook
    Dim CurrentName As Name
    Set CurrentName = FromBook.Names(NamedRangeName)

    On Error GoTo NoRefersToRange
    GetNamedRangeValue = CurrentName.RefersToRange.Value
    Exit Function

NoRefersToRange:
    If Err.Number = 1004 Then
        GetNamedRangeValue = Evaluate("=" & CurrentName.NameLocal)
        Err.Clear
    Else
        Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
    End If

End Function

Public Function GetRefersTo(ByVal NamedRangeName As String _
                            , Optional ByVal FromBook As Workbook) As String


    If FromBook Is Nothing Then Set FromBook = ThisWorkbook
    GetRefersTo = FromBook.Names(NamedRangeName).RefersTo

End Function

Public Function Slugify(ByVal Text As String) As String

    Dim Result As String
    Result = Text
    Result = Replace(Result, " ", "-")
    ' Only keep 0-9, a-z, A-Z, and -.
    Dim Slug As String
    Dim Index As Long
    For Index = 1 To Len(Result)
        Dim CurrentChar As String
        CurrentChar = Mid$(Result, Index, 1)
        If (CurrentChar Like "[0-9a-zA-Z-]") Then
            Slug = Slug & CurrentChar
        End If
    Next Index

    Slug = LCase$(Slug)

    Slugify = Slug

End Function

Public Function UnSlugify(ByVal Slug As String, Optional ByVal Conversion As VbStrConv = vbProperCase) As String

    Dim Result As String
    Result = Slug
    Result = Replace(Result, "-", " ")
    Result = StrConv(Result, Conversion)
    UnSlugify = Result

End Function

Public Function GetFilteredCriteria(ByVal FromRange As Range) As Variant

    ' This doesn't work with Date Filter if you use Tree view structure of date filter.
    ' If you need it to work with date and tree view structure then look into this: https://stackoverflow.com/questions/32040238/get-date-autofilter-in-excel-vba/32062115#32062115
    ' For multiple values (xlFilterValues) this function wrap each item with quotes and newline. So for example if the criteria is Dhaka, Rajshahi and Khulna then it will be like
    ' "=Dhaka"
    ' "=Rajshahi"
    ' "=Khulna"
    ' So if you want to use this and re-create the auto-filter then keep these in mind.

    If FromRange.Worksheet.AutoFilter Is Nothing Then Exit Function

    With FromRange.Worksheet.AutoFilter
        If Intersect(.Range, FromRange) Is Nothing Then Exit Function

        Dim ValidFilterIndexes As Collection
        Set ValidFilterIndexes = New Collection

        Dim Index As Long
        For Index = 1 To .Filters.Count
            If .Filters(Index).On Then ValidFilterIndexes.Add Index
        Next Index

        Dim FilterCriteria As Variant
        ReDim FilterCriteria(1 To ValidFilterIndexes.Count + 1, 1 To 4) As Variant

        FilterCriteria(1, 1) = "Col Index"
        FilterCriteria(1, 2) = "Criteria1"
        FilterCriteria(1, 3) = "Criteria2"
        FilterCriteria(1, 4) = "Operator"

        Index = 1
        Dim ValidFilterIndex As Variant
        For Each ValidFilterIndex In ValidFilterIndexes

            Index = Index + 1
            Dim CurrentFilter As Filter
            Set CurrentFilter = .Filters(CLng(ValidFilterIndex))
            FilterCriteria(Index, 1) = ValidFilterIndex
            FilterCriteria(Index, 4) = CurrentFilter.Operator
            If CurrentFilter.Operator = xlFilterValues Then
                FilterCriteria(Index, 2) = """" & Join(CurrentFilter.Criteria1, """" & vbNewLine & """") & """"
            Else
                FilterCriteria(Index, 2) = CurrentFilter.Criteria1
            End If

            If CurrentFilter.Operator = xlAnd Or CurrentFilter.Operator = xlOr Then
                FilterCriteria(Index, 3) = CurrentFilter.Criteria2
            End If

        Next ValidFilterIndex

    End With

    GetFilteredCriteria = FilterCriteria

End Function

Public Function IsSubRange(ByVal ParentRange As Range _
                           , ByVal ChildRange As Range) As Boolean

    If ChildRange Is Nothing Then Exit Function
    If ParentRange Is Nothing Then Exit Function

    Dim InterSectionRange As Range
    Set InterSectionRange = Intersect(ParentRange, ChildRange)

    If InterSectionRange Is Nothing Then Exit Function
    IsSubRange = (ChildRange.Address = InterSectionRange.Address)

End Function

Public Function IsBlankRange(ByVal CheckRange As Range) As Boolean

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

Public Function DecodeBase64EncodedString(ByVal Base64Text As String) As String

    ' Function to decode Base64 encoded string
    ' Sample Input:
    ' Base64Text:eyJTYW1wbGUiOjEyfQ==
    ' Output:{"Sample":12}

    With CreateObject("MSXML2.DOMDocument.6.0").createElement("b64")
        .DataType = "bin.base64"
        .Text = Base64Text
        DecodeBase64EncodedString = StrConv(.nodeTypedValue, vbUnicode)
    End With

End Function

Public Function EncodeToBase64(ByVal Text As String) As String

    ' Function to encode text to Base64.
    ' Sample Input:
    ' Base64Text:{"Sample":12}
    ' Output:eyJTYW1wbGUiOjEyfQ==

    With CreateObject("MSXML2.DOMDocument.6.0").createElement("b64")
        .DataType = "bin.base64"
        Dim BinaryData() As Byte
        BinaryData = StrConv(Text, vbFromUnicode)
        .nodeTypedValue = BinaryData
        EncodeToBase64 = Replace(.Text, vbLf, vbNullString)
    End With

End Function


Public Function GetASCIICodeDescriptionMap() As Variant()

    Dim ASCIICodeDescriptionMap(1 To 257, 1 To 2) As Variant

    ASCIICodeDescriptionMap(1, 1) = "ASCII Code"
    ASCIICodeDescriptionMap(1, 2) = "Description"

    Dim Counter As Long
    For Counter = 2 To 257
        ASCIICodeDescriptionMap(Counter, 1) = Counter - 2
    Next Counter

    ASCIICodeDescriptionMap(2, 2) = "Null"
    ASCIICodeDescriptionMap(3, 2) = "Start of Heading"
    ASCIICodeDescriptionMap(4, 2) = "Start of Text"
    ASCIICodeDescriptionMap(5, 2) = "End of Text"
    ASCIICodeDescriptionMap(6, 2) = "End of Transmission"
    ASCIICodeDescriptionMap(7, 2) = "Enquiry"
    ASCIICodeDescriptionMap(8, 2) = "Acknowledge"
    ASCIICodeDescriptionMap(9, 2) = "Bell"
    ASCIICodeDescriptionMap(10, 2) = "Backspace"
    ASCIICodeDescriptionMap(11, 2) = "Horizontal Tab"
    ASCIICodeDescriptionMap(12, 2) = "NL Line Feed, New Line"
    ASCIICodeDescriptionMap(13, 2) = "Vertical Tab"
    ASCIICodeDescriptionMap(14, 2) = "NP Form Feed, New Page"
    ASCIICodeDescriptionMap(15, 2) = "Carriage Return"
    ASCIICodeDescriptionMap(16, 2) = "Shift Out"
    ASCIICodeDescriptionMap(17, 2) = "Shift In"
    ASCIICodeDescriptionMap(18, 2) = "Data Link Escape"
    ASCIICodeDescriptionMap(19, 2) = "Device Control 1"
    ASCIICodeDescriptionMap(20, 2) = "Device Control 2"
    ASCIICodeDescriptionMap(21, 2) = "Device Control 3"
    ASCIICodeDescriptionMap(22, 2) = "Device Control 4"
    ASCIICodeDescriptionMap(23, 2) = "Negative Acknowledge"
    ASCIICodeDescriptionMap(24, 2) = "Synchronous Idle"
    ASCIICodeDescriptionMap(25, 2) = "End of Transmission Block"
    ASCIICodeDescriptionMap(26, 2) = "Cancel"
    ASCIICodeDescriptionMap(27, 2) = "End of Medium"
    ASCIICodeDescriptionMap(28, 2) = "Substitute"
    ASCIICodeDescriptionMap(29, 2) = "Escape"
    ASCIICodeDescriptionMap(30, 2) = "File Separator"
    ASCIICodeDescriptionMap(31, 2) = "Group Separator"
    ASCIICodeDescriptionMap(32, 2) = "Record Separator"
    ASCIICodeDescriptionMap(33, 2) = "Unit Separator"
    ASCIICodeDescriptionMap(34, 2) = "Space"
    ASCIICodeDescriptionMap(35, 2) = "Exclamation Mark"
    ASCIICodeDescriptionMap(36, 2) = "Double Quote"
    ASCIICodeDescriptionMap(37, 2) = "Hash or Number"
    ASCIICodeDescriptionMap(38, 2) = "Dollar Sign"
    ASCIICodeDescriptionMap(39, 2) = "Percentage"
    ASCIICodeDescriptionMap(40, 2) = "Ampersand"
    ASCIICodeDescriptionMap(41, 2) = "Single Quote"
    ASCIICodeDescriptionMap(42, 2) = "Left Parenthesis"
    ASCIICodeDescriptionMap(43, 2) = "Right Parenthesis"
    ASCIICodeDescriptionMap(44, 2) = "Asterisk"
    ASCIICodeDescriptionMap(45, 2) = "Plus Sign"
    ASCIICodeDescriptionMap(46, 2) = "Comma"
    ASCIICodeDescriptionMap(47, 2) = "Minus Sign"
    ASCIICodeDescriptionMap(48, 2) = "Period"
    ASCIICodeDescriptionMap(49, 2) = "Slash"
    ASCIICodeDescriptionMap(50, 2) = "Zero"
    ASCIICodeDescriptionMap(51, 2) = "Number One"
    ASCIICodeDescriptionMap(52, 2) = "Number Two"
    ASCIICodeDescriptionMap(53, 2) = "Number Three"
    ASCIICodeDescriptionMap(54, 2) = "Number Four"
    ASCIICodeDescriptionMap(55, 2) = "Number Five"
    ASCIICodeDescriptionMap(56, 2) = "Number Six"
    ASCIICodeDescriptionMap(57, 2) = "Number Seven"
    ASCIICodeDescriptionMap(58, 2) = "Number Eight"
    ASCIICodeDescriptionMap(59, 2) = "Number Nine"
    ASCIICodeDescriptionMap(60, 2) = "Colon"
    ASCIICodeDescriptionMap(61, 2) = "Semicolon"
    ASCIICodeDescriptionMap(62, 2) = "Less Than"
    ASCIICodeDescriptionMap(63, 2) = "Equals Sign"
    ASCIICodeDescriptionMap(64, 2) = "Greater Than"
    ASCIICodeDescriptionMap(65, 2) = "Question Mark"
    ASCIICodeDescriptionMap(66, 2) = "At Sign"
    ASCIICodeDescriptionMap(67, 2) = "Upper Case Letter A"
    ASCIICodeDescriptionMap(68, 2) = "Upper Case Letter B"
    ASCIICodeDescriptionMap(69, 2) = "Upper Case Letter C"
    ASCIICodeDescriptionMap(70, 2) = "Upper Case Letter D"
    ASCIICodeDescriptionMap(71, 2) = "Upper Case Letter E"
    ASCIICodeDescriptionMap(72, 2) = "Upper Case Letter F"
    ASCIICodeDescriptionMap(73, 2) = "Upper Case Letter G"
    ASCIICodeDescriptionMap(74, 2) = "Upper Case Letter H"
    ASCIICodeDescriptionMap(75, 2) = "Upper Case Letter I"
    ASCIICodeDescriptionMap(76, 2) = "Upper Case Letter J"
    ASCIICodeDescriptionMap(77, 2) = "Upper Case Letter K"
    ASCIICodeDescriptionMap(78, 2) = "Upper Case Letter L"
    ASCIICodeDescriptionMap(79, 2) = "Upper Case Letter M"
    ASCIICodeDescriptionMap(80, 2) = "Upper Case Letter N"
    ASCIICodeDescriptionMap(81, 2) = "Upper Case Letter O"
    ASCIICodeDescriptionMap(82, 2) = "Upper Case Letter P"
    ASCIICodeDescriptionMap(83, 2) = "Upper Case Letter Q"
    ASCIICodeDescriptionMap(84, 2) = "Upper Case Letter R"
    ASCIICodeDescriptionMap(85, 2) = "Upper Case Letter S"
    ASCIICodeDescriptionMap(86, 2) = "Upper Case Letter T"
    ASCIICodeDescriptionMap(87, 2) = "Upper Case Letter U"
    ASCIICodeDescriptionMap(88, 2) = "Upper Case Letter V"
    ASCIICodeDescriptionMap(89, 2) = "Upper Case Letter W"
    ASCIICodeDescriptionMap(90, 2) = "Upper Case Letter X"
    ASCIICodeDescriptionMap(91, 2) = "Upper Case Letter Y"
    ASCIICodeDescriptionMap(92, 2) = "Upper Case Letter Z"
    ASCIICodeDescriptionMap(93, 2) = "Left Square Bracket"
    ASCIICodeDescriptionMap(94, 2) = "Backslash"
    ASCIICodeDescriptionMap(95, 2) = "Right Square Bracket"
    ASCIICodeDescriptionMap(96, 2) = "Caret or Circumflex"
    ASCIICodeDescriptionMap(97, 2) = "Underscore"
    ASCIICodeDescriptionMap(98, 2) = "Grave Accent"
    ASCIICodeDescriptionMap(99, 2) = "Lower Case Letter a"
    ASCIICodeDescriptionMap(100, 2) = "Lower Case Letter b"
    ASCIICodeDescriptionMap(101, 2) = "Lower Case Letter c"
    ASCIICodeDescriptionMap(102, 2) = "Lower Case Letter d"
    ASCIICodeDescriptionMap(103, 2) = "Lower Case Letter e"
    ASCIICodeDescriptionMap(104, 2) = "Lower Case Letter f"
    ASCIICodeDescriptionMap(105, 2) = "Lower Case Letter g"
    ASCIICodeDescriptionMap(106, 2) = "Lower Case Letter h"
    ASCIICodeDescriptionMap(107, 2) = "Lower Case Letter i"
    ASCIICodeDescriptionMap(108, 2) = "Lower Case Letter j"
    ASCIICodeDescriptionMap(109, 2) = "Lower Case Letter k"
    ASCIICodeDescriptionMap(110, 2) = "Lower Case Letter l"
    ASCIICodeDescriptionMap(111, 2) = "Lower Case Letter m"
    ASCIICodeDescriptionMap(112, 2) = "Lower Case Letter n"
    ASCIICodeDescriptionMap(113, 2) = "Lower Case Letter o"
    ASCIICodeDescriptionMap(114, 2) = "Lower Case Letter p"
    ASCIICodeDescriptionMap(115, 2) = "Lower Case Letter q"
    ASCIICodeDescriptionMap(116, 2) = "Lower Case Letter r"
    ASCIICodeDescriptionMap(117, 2) = "Lower Case Letter s"
    ASCIICodeDescriptionMap(118, 2) = "Lower Case Letter t"
    ASCIICodeDescriptionMap(119, 2) = "Lower Case Letter u"
    ASCIICodeDescriptionMap(120, 2) = "Lower Case Letter v"
    ASCIICodeDescriptionMap(121, 2) = "Lower Case Letter w"
    ASCIICodeDescriptionMap(122, 2) = "Lower Case Letter x"
    ASCIICodeDescriptionMap(123, 2) = "Lower Case Letter y"
    ASCIICodeDescriptionMap(124, 2) = "Lower Case Letter z"
    ASCIICodeDescriptionMap(125, 2) = "Left Curly Bracket"
    ASCIICodeDescriptionMap(126, 2) = "Vertical Bar"
    ASCIICodeDescriptionMap(127, 2) = "Right Curly Bracket"
    ASCIICodeDescriptionMap(128, 2) = "Tilde"
    ASCIICodeDescriptionMap(129, 2) = "Delete"
    ASCIICodeDescriptionMap(130, 2) = "Latin Capital Letter C With Cedilla"
    ASCIICodeDescriptionMap(131, 2) = "Latin Small Letter U With Diaeresis"
    ASCIICodeDescriptionMap(132, 2) = "Latin Small Letter E With Acute"
    ASCIICodeDescriptionMap(133, 2) = "Latin Small Letter A With Circumflex"
    ASCIICodeDescriptionMap(134, 2) = "Latin Small Letter A With Diaeresis"
    ASCIICodeDescriptionMap(135, 2) = "Latin Small Letter A With Grave"
    ASCIICodeDescriptionMap(136, 2) = "Latin Small Letter A With Ring Above"
    ASCIICodeDescriptionMap(137, 2) = "Latin Small Letter C With Cedilla"
    ASCIICodeDescriptionMap(138, 2) = "Latin Small Letter E With Circumflex"
    ASCIICodeDescriptionMap(139, 2) = "Latin Small Letter E With Diaeresis"
    ASCIICodeDescriptionMap(140, 2) = "Latin Small Letter E With Grave"
    ASCIICodeDescriptionMap(141, 2) = "Latin Small Letter I With Diaeresis"
    ASCIICodeDescriptionMap(142, 2) = "Latin Small Letter I With Circumflex"
    ASCIICodeDescriptionMap(143, 2) = "Latin Small Letter I With Grave"
    ASCIICodeDescriptionMap(144, 2) = "Latin Capital Letter A With Diaeresis"
    ASCIICodeDescriptionMap(145, 2) = "Latin Capital Letter A With Ring Above"
    ASCIICodeDescriptionMap(146, 2) = "Latin Capital Letter E With Acute"
    ASCIICodeDescriptionMap(147, 2) = "Latin Small Letter Ae"
    ASCIICodeDescriptionMap(148, 2) = "Latin Capital Letter Ae"
    ASCIICodeDescriptionMap(149, 2) = "Latin Small Letter O With Circumflex"
    ASCIICodeDescriptionMap(150, 2) = "Latin Small Letter O With Diaeresis"
    ASCIICodeDescriptionMap(151, 2) = "Latin Small Letter O With Grave"
    ASCIICodeDescriptionMap(152, 2) = "Latin Small Letter U With Circumflex"
    ASCIICodeDescriptionMap(153, 2) = "Latin Small Letter U With Grave"
    ASCIICodeDescriptionMap(154, 2) = "Latin Small Letter Y With Diaeresis"
    ASCIICodeDescriptionMap(155, 2) = "Latin Capital Letter O With Diaeresis"
    ASCIICodeDescriptionMap(156, 2) = "Latin Capital Letter U With Diaeresis"
    ASCIICodeDescriptionMap(157, 2) = "Cent Sign"
    ASCIICodeDescriptionMap(158, 2) = "Pound Sign, Pound Sterling, Irish Punt, Lira Sign"
    ASCIICodeDescriptionMap(159, 2) = "Yen Sign, Yuan Sign"
    ASCIICodeDescriptionMap(160, 2) = "Peseta Sign"
    ASCIICodeDescriptionMap(161, 2) = "Latin Small Letter F With Hook, Florin Currency Symbol, Function Symbol"
    ASCIICodeDescriptionMap(162, 2) = "Latin Small Letter A With Acute"
    ASCIICodeDescriptionMap(163, 2) = "Latin Small Letter I With Acute"
    ASCIICodeDescriptionMap(164, 2) = "Latin Small Letter O With Acute"
    ASCIICodeDescriptionMap(165, 2) = "Latin Small Letter U With Acute"
    ASCIICodeDescriptionMap(166, 2) = "Latin Small Letter N With Tilde, Small Letter Enye"
    ASCIICodeDescriptionMap(167, 2) = "Latin Capital Letter N With Tilde, Capital Letter Enye"
    ASCIICodeDescriptionMap(168, 2) = "Feminine Ordinal Indicator"
    ASCIICodeDescriptionMap(169, 2) = "Masculine Ordinal Indicator"
    ASCIICodeDescriptionMap(170, 2) = "Inverted Question Mark, Turned Question Mark"
    ASCIICodeDescriptionMap(171, 2) = "Reversed Not Sign, Beginning Of Line"
    ASCIICodeDescriptionMap(172, 2) = "Not Sign, Angled Dash"
    ASCIICodeDescriptionMap(173, 2) = "Vulgar Fraction One Half"
    ASCIICodeDescriptionMap(174, 2) = "Vulgar Fraction One Quarter"
    ASCIICodeDescriptionMap(175, 2) = "Inverted Exclamation Mark"
    ASCIICodeDescriptionMap(176, 2) = "Left-Pointing Double Angle Quotation Mark, Left Guillemet, Chevrons"
    ASCIICodeDescriptionMap(177, 2) = "Right-Pointing Double Angle Quotation Mark, Right Guillemet"
    ASCIICodeDescriptionMap(178, 2) = "Light Shade"
    ASCIICodeDescriptionMap(179, 2) = "Medium Shade, Speckles Fill, Dotted Fill"
    ASCIICodeDescriptionMap(180, 2) = "Dark Shade"
    ASCIICodeDescriptionMap(181, 2) = "Box Drawings Light Vertical"
    ASCIICodeDescriptionMap(182, 2) = "Box Drawings Light Vertical And Left"
    ASCIICodeDescriptionMap(183, 2) = "Box Drawings Vertical Single And Left Double"
    ASCIICodeDescriptionMap(184, 2) = "Box Drawings Vertical Double And Left Single"
    ASCIICodeDescriptionMap(185, 2) = "Box Drawings Down Double And Left Single"
    ASCIICodeDescriptionMap(186, 2) = "Box Drawings Down Single And Left Double"
    ASCIICodeDescriptionMap(187, 2) = "Box Drawings Double Vertical And Left"
    ASCIICodeDescriptionMap(188, 2) = "Box Drawings Double Vertical"
    ASCIICodeDescriptionMap(189, 2) = "Box Drawings Double Down And Left"
    ASCIICodeDescriptionMap(190, 2) = "Box Drawings Double Up And Left"
    ASCIICodeDescriptionMap(191, 2) = "Box Drawings Up Double And Left Single"
    ASCIICodeDescriptionMap(192, 2) = "Box Drawings Up Single And Left Double"
    ASCIICodeDescriptionMap(193, 2) = "Box Drawings Light Down And Left"
    ASCIICodeDescriptionMap(194, 2) = "Box Drawings Light Up And Right"
    ASCIICodeDescriptionMap(195, 2) = "Box Drawings Light Up And Horizontal"
    ASCIICodeDescriptionMap(196, 2) = "Box Drawings Light Down And Horizontal"
    ASCIICodeDescriptionMap(197, 2) = "Box Drawings Light Vertical And Right"
    ASCIICodeDescriptionMap(198, 2) = "Box Drawings Light Horizontal"
    ASCIICodeDescriptionMap(199, 2) = "Box Drawings Light Vertical And Horizontal"
    ASCIICodeDescriptionMap(200, 2) = "Box Drawings Vertical Single And Right Double"
    ASCIICodeDescriptionMap(201, 2) = "Box Drawings Vertical Double And Right Single"
    ASCIICodeDescriptionMap(202, 2) = "Box Drawings Double Up And Right"
    ASCIICodeDescriptionMap(203, 2) = "Box Drawings Double Down And Right"
    ASCIICodeDescriptionMap(204, 2) = "Box Drawings Double Up And Horizontal"
    ASCIICodeDescriptionMap(205, 2) = "Box Drawings Double Down And Horizontal"
    ASCIICodeDescriptionMap(206, 2) = "Box Drawings Double Vertical And Right"
    ASCIICodeDescriptionMap(207, 2) = "Box Drawings Double Horizontal"
    ASCIICodeDescriptionMap(208, 2) = "Box Drawings Double Vertical And Horizontal"
    ASCIICodeDescriptionMap(209, 2) = "Box Drawings Up Single And Horizontal Double"
    ASCIICodeDescriptionMap(210, 2) = "Box Drawings Up Double And Horizontal Single"
    ASCIICodeDescriptionMap(211, 2) = "Box Drawings Down Single And Horizontal Double"
    ASCIICodeDescriptionMap(212, 2) = "Box Drawings Down Double And Horizontal Single"
    ASCIICodeDescriptionMap(213, 2) = "Box Drawings Up Double And Right Single"
    ASCIICodeDescriptionMap(214, 2) = "Box Drawings Up Single And Right Double"
    ASCIICodeDescriptionMap(215, 2) = "Box Drawings Down Single And Right Double"
    ASCIICodeDescriptionMap(216, 2) = "Box Drawings Down Double And Right Single"
    ASCIICodeDescriptionMap(217, 2) = "Box Drawings Vertical Double And Horizontal Single"
    ASCIICodeDescriptionMap(218, 2) = "Box Drawings Vertical Single And Horizontal Double"
    ASCIICodeDescriptionMap(219, 2) = "Box Drawings Light Up And Left"
    ASCIICodeDescriptionMap(220, 2) = "Box Drawings Light Down And Right"
    ASCIICodeDescriptionMap(221, 2) = "Full Block, Solid Block"
    ASCIICodeDescriptionMap(222, 2) = "Lower Half Block"
    ASCIICodeDescriptionMap(223, 2) = "Left Half Block"
    ASCIICodeDescriptionMap(224, 2) = "Right Half Block"
    ASCIICodeDescriptionMap(225, 2) = "Upper Half Block"
    ASCIICodeDescriptionMap(226, 2) = "Greek Small Letter Alpha"
    ASCIICodeDescriptionMap(227, 2) = "Latin Small Letter Sharp S, Eszett"
    ASCIICodeDescriptionMap(228, 2) = "Greek Capital Letter Gamma"
    ASCIICodeDescriptionMap(229, 2) = "Greek Small Letter Pi"
    ASCIICodeDescriptionMap(230, 2) = "Greek Capital Letter Sigma"
    ASCIICodeDescriptionMap(231, 2) = "Greek Small Letter Sigma"
    ASCIICodeDescriptionMap(232, 2) = "Micro Sign"
    ASCIICodeDescriptionMap(233, 2) = "Greek Capital Letter Tau"
    ASCIICodeDescriptionMap(234, 2) = "Greek Capital Letter Phi"
    ASCIICodeDescriptionMap(235, 2) = "Greek Capital Letter Theta"
    ASCIICodeDescriptionMap(236, 2) = "Greek Capital Letter Omega"
    ASCIICodeDescriptionMap(237, 2) = "Greek Small Letter Delta"
    ASCIICodeDescriptionMap(238, 2) = "Infinity"
    ASCIICodeDescriptionMap(239, 2) = "Greek Small Letter Phi"
    ASCIICodeDescriptionMap(240, 2) = "Greek Small Letter Epsilon"
    ASCIICodeDescriptionMap(241, 2) = "Intersection"
    ASCIICodeDescriptionMap(242, 2) = "Identical To"
    ASCIICodeDescriptionMap(243, 2) = "Plus-Minus Sign"
    ASCIICodeDescriptionMap(244, 2) = "Greater-Than Or Equal To"
    ASCIICodeDescriptionMap(245, 2) = "Less-Than Or Equal To"
    ASCIICodeDescriptionMap(246, 2) = "Top Half Integral"
    ASCIICodeDescriptionMap(247, 2) = "Bottom Half Integral"
    ASCIICodeDescriptionMap(248, 2) = "Division Sign, Obelus"
    ASCIICodeDescriptionMap(249, 2) = "Almost Equal To, Asymptotic To"
    ASCIICodeDescriptionMap(250, 2) = "Degree Sign"
    ASCIICodeDescriptionMap(251, 2) = "Bullet Operator"
    ASCIICodeDescriptionMap(252, 2) = "Middle Dot, Interpunct"
    ASCIICodeDescriptionMap(253, 2) = "Square Root, Radical Sign"
    ASCIICodeDescriptionMap(254, 2) = "Superscript Latin Small Letter N"
    ASCIICodeDescriptionMap(255, 2) = "Superscript Two, Squared"
    ASCIICodeDescriptionMap(256, 2) = "Black Square"
    ASCIICodeDescriptionMap(257, 2) = "Non-Breaking Space, NBSP"

    GetASCIICodeDescriptionMap = ASCIICodeDescriptionMap

End Function

Public Function GetASCIICodeMapAsArrayConst() As String

    Dim ASCIICodeDescriptionMap() As Variant
    ASCIICodeDescriptionMap = GetASCIICodeDescriptionMap()

    Dim ArrayConst As String
    ArrayConst = "{" & vbNewLine
    Dim FirstColumnIndex  As Long
    FirstColumnIndex = LBound(ASCIICodeDescriptionMap, 2)
    Dim RowIndex As Long
    Const THREE_SPACE As String = "   "
    Const ONE_SPACE As String = " "

    Dim ColSeparator As String
    ColSeparator = Application.International(XlApplicationInternational.xlColumnSeparator)

    Dim RowSeparator As String
    RowSeparator = Application.International(XlApplicationInternational.xlRowSeparator)

    For RowIndex = LBound(ASCIICodeDescriptionMap, 1) To UBound(ASCIICodeDescriptionMap, 1)

        ArrayConst = ArrayConst & THREE_SPACE
        If IsNumeric(ASCIICodeDescriptionMap(RowIndex, FirstColumnIndex)) Then
            ArrayConst = ArrayConst & ASCIICodeDescriptionMap(RowIndex, FirstColumnIndex)
        Else
            ArrayConst = ArrayConst & """" & ASCIICodeDescriptionMap(RowIndex, FirstColumnIndex) & """"
        End If

        ArrayConst = ArrayConst & ColSeparator & ONE_SPACE & """" _
                     & ASCIICodeDescriptionMap(RowIndex, FirstColumnIndex + 1) & """" _
                     & RowSeparator & vbNewLine

    Next RowIndex

    ArrayConst = Left$(ArrayConst, Len(ArrayConst) - Len(RowSeparator & vbNewLine)) & vbNewLine & "}"

    GetASCIICodeMapAsArrayConst = ArrayConst

End Function

Public Function IsNamedRangeExist(ByVal SearchInBook As Workbook _
                                  , ByVal NameOfTheNamedRange As String _
                                  , Optional ByVal ScopeSheetName As String = vbNullString) As Boolean

    ' Checks if a named range exists in the given workbook.

    Dim NamesContainer As Object
    If ScopeSheetName = vbNullString Then
        Set NamesContainer = SearchInBook
    ElseIf IsSheetExist(ScopeSheetName, SearchInBook) Then
        Set NamesContainer = SearchInBook.Worksheets(ScopeSheetName)
    Else
        IsNamedRangeExist = False
        Exit Function
    End If

    Dim IsExist As Boolean
    Dim CurrentName As Name
    For Each CurrentName In NamesContainer.Names
        If CurrentName.Name = NameOfTheNamedRange Then
            IsExist = True
            Exit For
        End If
    Next CurrentName

    IsNamedRangeExist = IsExist

End Function

Public Function IsQueryExists(ByVal SearchInBook As Workbook _
                              , ByVal QueryName As String) As Boolean

    Dim CurrentQuery As WorkbookQuery
    For Each CurrentQuery In SearchInBook.Queries
        If CurrentQuery.Name = QueryName Then
            IsQueryExists = True
            Exit Function
        End If
    Next CurrentQuery

    IsQueryExists = False

End Function

Public Function GetUniqueQueryNameByIncrementingNumber(ByVal SearchInBook As Workbook _
                                                       , ByVal QueryName As String _
                                                       , Optional ByVal NumberSeparator As String = vbNullString) As String

    ' It will check if provided query is present or not. If not present then return back same name.
    ' But if it is present then it will keep adding number starting from 2.
    ' So if you provide APIKey as name then it will be APIKey2, APIKey3 and so on.
    ' But if you specify number separator then it will use that. For example if NumberSeparator is _ then
    ' It will be like APIKey_2, APIKey_3 and so on.
    ' Don't provide like APIKey2 as QueryName as it will not extract last number from end.

    If Not IsQueryExists(SearchInBook, QueryName) Then
        GetUniqueQueryNameByIncrementingNumber = QueryName
        Exit Function
    End If

    Dim Counter As Long
    Counter = 2
    Do While IsQueryExists(SearchInBook, QueryName & NumberSeparator & Counter)
        Counter = Counter + 1
    Loop

    GetUniqueQueryNameByIncrementingNumber = QueryName & NumberSeparator & Counter

End Function

Public Function GetUniqueNamedRangeNameByIncrementingNumber(ByVal SearchInBook As Workbook _
                                                       , ByVal NamedRangeName As String _
                                                       , Optional ByVal WorksheetName As String = vbNullString _
                                                       , Optional ByVal NumberSeparator As String = vbNullString) As String

    ' It will check if provided named range is present or not. If not present then return back same name.
    ' But if it is present then it will keep adding number starting from 2.
    ' So if you provide APIKey as name then it will be APIKey2, APIKey3 and so on.
    ' But if you specify number separator then it will use that. For example if NumberSeparator is _ then
    ' It will be like APIKey_2, APIKey_3 and so on.
    ' Don't provide like APIKey2 as NamedRangeName as it will not extract last number from end.

    If Not IsNamedRangeExist(SearchInBook, NamedRangeName, WorksheetName) Then
        GetUniqueNamedRangeNameByIncrementingNumber = NamedRangeName
        Exit Function
    End If

    Dim Counter As Long
    Counter = 2
    Do While IsNamedRangeExist(SearchInBook, NamedRangeName & NumberSeparator & Counter, WorksheetName)
        Counter = Counter + 1
    Loop

    GetUniqueNamedRangeNameByIncrementingNumber = NamedRangeName & NumberSeparator & Counter

End Function

Public Function GetWords(ByVal Text As String) As Variant

    Dim CleanText As String
    CleanText = GetOnlyAlphanumericCharcter(Text, IsTrimAfter:=True)
    GetWords = Split(CleanText, " ")

End Function

Public Function CollectionOfArrayTo2DArray(ByVal Map As Collection, Optional ByVal MaxItemInEachArr As Long = -1) As Variant

    ' If MaxItemInEachArr is -1 then it will figure out the number of columns based on the maximum number of items in each array.
    ' If MaxItemInEachArr is greater than 0 then it will create that many columns and fill the data accordingly.
    Dim MaxItems As Long
    MaxItems = MaxItemInEachArr
    Dim CurrentItem As Variant
    If MaxItemInEachArr = -1 Then
        MaxItems = 1
        For Each CurrentItem In Map
            If IsArray(CurrentItem) Then
                If UBound(CurrentItem) - LBound(CurrentItem) + 1 > MaxItems Then
                    MaxItems = UBound(CurrentItem) - LBound(CurrentItem) + 1
                End If
            End If
        Next CurrentItem
    End If

    Dim Result As Variant
    ReDim Result(1 To Map.Count, 1 To MaxItems)

    Dim RowIndex As Long
    Dim ColumnIndex As Long
    For Each CurrentItem In Map
        RowIndex = RowIndex + 1
        Dim Counter As Long
        If IsArray(CurrentItem) Then
            Counter = 0
            For ColumnIndex = LBound(CurrentItem) To UBound(CurrentItem)
                Counter = Counter + 1
                If Counter > MaxItems Then Exit For
                Result(RowIndex, Counter) = CurrentItem(ColumnIndex)
            Next ColumnIndex
        Else
            Result(RowIndex, 1) = CurrentItem
        End If
    Next CurrentItem

    CollectionOfArrayTo2DArray = Result

End Function

Public Function GetVSCodePath() As String

    Dim WshShell As Object
    Dim VSCodePath As String

    ' Create a new Shell object
    Set WshShell = CreateObject("WScript.Shell")

    ' The registry key where VS Code path is stored
    Const VS_CODE_KEY As String = "HKEY_CLASSES_ROOT\Applications\Code.exe\shell\open\command\"
    On Error Resume Next
    VSCodePath = WshShell.RegRead(VS_CODE_KEY)
    On Error GoTo 0

    If InStr(VSCodePath, """") > 0 Then
        VSCodePath = Mid$(VSCodePath, InStr(VSCodePath, """") + 1)
        VSCodePath = Left$(VSCodePath, InStr(VSCodePath, """") - 1)
    End If

    GetVSCodePath = VSCodePath
    Set WshShell = Nothing

End Function

Public Function GetTable(ByVal FromBook As Workbook, ByVal TableName As String) As ListObject

    Dim CurrentSheet As Worksheet
    For Each CurrentSheet In FromBook.Worksheets
        Dim CurrentTable As ListObject
        For Each CurrentTable In CurrentSheet.ListObjects
            If CurrentTable.Name = TableName Then
                Set GetTable = CurrentTable
                Exit Function
            End If
        Next CurrentTable
    Next CurrentSheet

    Set GetTable = Nothing

End Function

Public Function GetSelectedShapes(ByVal SelectedItems As Variant) As Collection

    Dim SelectedShapes As Collection
    Set SelectedShapes = New Collection

    On Error GoTo HandleError
    Dim CurrentItem As Variant
    For Each CurrentItem In SelectedItems.ShapeRange
        If CurrentItem.Type <> msoChart And CurrentItem.Type <> msoSlicer Then
            SelectedShapes.Add CurrentItem
        End If
    Next CurrentItem

    Set GetSelectedShapes = SelectedShapes
    Exit Function

HandleError:

    Set GetSelectedShapes = New Collection
    Err.Clear

End Function

Public Function FindMatchingColHeaders(ByVal Table1 As ListObject, ByVal Table2 As ListObject) As String()

    Dim MatchingColHeaders As Collection
    Set MatchingColHeaders = New Collection

    Dim CurrentListCol As ListColumn
    For Each CurrentListCol In Table1.ListColumns
        If IsColumnExist(Table2, CurrentListCol.Name) Then
            MatchingColHeaders.Add CurrentListCol.Name
        End If
    Next CurrentListCol

    Dim Result() As String
    If MatchingColHeaders.Count <> 0 Then

        ReDim Result(1 To MatchingColHeaders.Count, 1 To 2) As String
        Dim CurrentHeader As Variant
        Dim Counter As Long
        For Each CurrentHeader In MatchingColHeaders
            Counter = Counter + 1
            Result(Counter, 1) = CurrentHeader
            Result(Counter, 2) = CurrentHeader
        Next CurrentHeader

    End If

    FindMatchingColHeaders = Result

End Function

Public Function ExpandCollectionOfVector(ByVal VectorColl As Collection) As Variant

    ' Expand 1D vector Collection to a 2D array.
    ' Row Count = Total Item in the Collection
    ' Col Count = First vector item count.
    ' So make sure that the Collection is of same size vector.

    If VectorColl.Count = 0 Then Exit Function

    Dim ColCount As Long
    ColCount = UBound(VectorColl.Item(1)) - LBound(VectorColl.Item(1)) + 1

    Dim Result As Variant
    ReDim Result(1 To VectorColl.Count, 1 To ColCount)

    Dim RowIndex As Long
    Dim TempVector As Variant
    For Each TempVector In VectorColl

        RowIndex = RowIndex + 1
        Dim CurrentItem As Variant
        Dim Counter As Long
        For Each CurrentItem In TempVector
            Counter = Counter + 1
            Result(RowIndex, Counter) = CurrentItem
        Next CurrentItem
        Counter = 0

    Next TempVector

    ExpandCollectionOfVector = Result

End Function

Public Function GetAllSheetsName(ByVal FromBook As Workbook, ByVal IsWithHeader As Boolean) As Variant

    ' This will return all sheets name as a 2D array of one col.

    Dim Result As Variant
    Dim Counter As Long

    If IsWithHeader Then
        ReDim Result(1 To FromBook.Worksheets.Count + 1, 1 To 1)
        Result(1, 1) = "Sheet Name"
        Counter = 1
    Else
        ReDim Result(1 To FromBook.Worksheets.Count, 1 To 1)
    End If

    Dim CurrentSheet As Worksheet
    For Each CurrentSheet In FromBook.Worksheets
        Counter = Counter + 1
        Result(Counter, 1) = CurrentSheet.Name
    Next CurrentSheet

    GetAllSheetsName = Result

End Function

Public Function GetAllTableInfo(ByVal FromBook As Workbook, ByVal IsWithHeader As Boolean) As Variant

    ' Create a 2D array of table name and sheet name.

    Dim InfoColl As Collection

    Set InfoColl = New Collection
    If IsWithHeader Then InfoColl.Add Array("Table Name", "Sheet Name")

    Dim CurrentSheet As Worksheet
    For Each CurrentSheet In FromBook.Worksheets

        Dim CurrentTable As ListObject
        For Each CurrentTable In CurrentSheet.ListObjects
            InfoColl.Add Array(CurrentTable.Name, CurrentSheet.Name)
        Next CurrentTable

    Next CurrentSheet

    GetAllTableInfo = ExpandCollectionOfVector(InfoColl)

End Function

Public Function GetAllPivotTableInfo(ByVal FromBook As Workbook, ByVal IsWithHeader As Boolean) As Variant

    ' Create a 2D array of pivot table name and sheet name.

    Dim InfoColl As Collection

    Set InfoColl = New Collection
    If IsWithHeader Then InfoColl.Add Array("Pivot Table Name", "Sheet Name")

    Dim CurrentSheet As Worksheet
    For Each CurrentSheet In FromBook.Worksheets

        Dim CurrentPivotTable As PivotTable
        For Each CurrentPivotTable In CurrentSheet.PivotTables
            InfoColl.Add Array(CurrentPivotTable.Name, CurrentSheet.Name)
        Next CurrentPivotTable

    Next CurrentSheet

    GetAllPivotTableInfo = ExpandCollectionOfVector(InfoColl)

End Function

Public Function GetAllChartInfo(ByVal FromBook As Workbook, ByVal IsWithHeader As Boolean) As Variant

    ' Create a 2D array of chart name and sheet name.

    Dim InfoColl As Collection

    Set InfoColl = New Collection
    If IsWithHeader Then InfoColl.Add Array("Chart Name", "Sheet Name")

    Dim CurrentSheet As Worksheet
    For Each CurrentSheet In FromBook.Worksheets

        Dim CurrentChartObject As ChartObject
        For Each CurrentChartObject In CurrentSheet.ChartObjects
            InfoColl.Add Array(CurrentChartObject.Name, CurrentSheet.Name)
        Next CurrentChartObject

    Next CurrentSheet

    GetAllChartInfo = ExpandCollectionOfVector(InfoColl)

End Function

Public Function GetAllShapeInfo(ByVal FromBook As Workbook, ByVal IsWithHeader As Boolean) As Variant

    ' Create a 2D array of shape name and sheet name.

    Dim InfoColl As Collection

    Set InfoColl = New Collection
    If IsWithHeader Then InfoColl.Add Array("Shape Name", "Sheet Name")

    Dim CurrentSheet As Worksheet
    For Each CurrentSheet In FromBook.Worksheets

        Dim CurrentShape As Shape
        For Each CurrentShape In CurrentSheet.Shapes
            If Not (CurrentShape.Type = msoChart Or CurrentShape.Type = msoSlicer) Then
                InfoColl.Add Array(CurrentShape.Name, CurrentSheet.Name)
            End If
        Next CurrentShape

    Next CurrentSheet

    GetAllShapeInfo = ExpandCollectionOfVector(InfoColl)

End Function

Public Function GetAllSlicerInfo(ByVal FromBook As Workbook, ByVal IsWithHeader As Boolean) As Variant

    ' Create a 2D array of shape name and sheet name.

    Dim InfoColl As Collection

    Set InfoColl = New Collection
    If IsWithHeader Then InfoColl.Add Array("Slicer Name", "Cache Name")

    Dim CurrentSlicerCache As SlicerCache
    For Each CurrentSlicerCache In FromBook.SlicerCaches

        Dim CurrentSlicer As Slicer
        For Each CurrentSlicer In CurrentSlicerCache.Slicers
            InfoColl.Add Array(CurrentSlicer.Name, CurrentSlicerCache.Name)
        Next CurrentSlicer

    Next CurrentSlicerCache

    GetAllSlicerInfo = ExpandCollectionOfVector(InfoColl)

End Function

Public Function GetAllPowerQueryInfo(ByVal FromBook As Workbook, ByVal IsWithHeader As Boolean) As Variant

    ' Create a 2D array of shape name and sheet name.

    Dim InfoColl As Collection

    Set InfoColl = New Collection
    If IsWithHeader Then InfoColl.Add Array("Query Name")

    Dim CurrentQuery As WorkbookQuery
    For Each CurrentQuery In FromBook.Queries
        InfoColl.Add Array(CurrentQuery.Name)
    Next CurrentQuery

    GetAllPowerQueryInfo = ExpandCollectionOfVector(InfoColl)

End Function

Public Function GetSpillRangeValue(ByVal FormulaCell As Range) As Variant

    If FormulaCell.HasSpill Then
        GetSpillRangeValue = FormulaCell.SpillParent.SpillingToRange.Value
    Else
        GetSpillRangeValue = FormulaCell.Value
    End If

End Function

Public Function ArrayToText(ByVal ArrayOrVector As Variant _
                            , Optional ByVal Delimiter As String = "," _
                             , Optional ByVal IsUniqueOnly As Boolean _
                              , Optional ByVal ColIndexIf2DArr As Long = -1) As String
    
    ' This is going to concate a vector or array col.
    ' You can remove duplicates by passing IsUniqueOnly to true.
    ' This is case insensitive if you want to consider unique once only.
    '@ExampleCall: ArrayToText(array("A","B","a"),,True) >> A,B
    
    
    If Not IsArray(ArrayOrVector) Then
        ArrayToText = ArrayOrVector
        Exit Function
    End If

    Dim Items As Collection
    Set Items = New Collection

    On Error Resume Next
    Dim CurrentItem As Variant
    If Is2DArray(ArrayOrVector) Then
        
        If ColIndexIf2DArr = -1 Then ColIndexIf2DArr = LBound(ArrayOrVector, 2)

        Dim RowIndex As Long
        For RowIndex = LBound(ArrayOrVector, 1) To UBound(ArrayOrVector, 1)
            CurrentItem = ArrayOrVector(RowIndex, ColIndexIf2DArr)
            If IsUniqueOnly Then
                Items.Add CurrentItem, CStr(CurrentItem)
            Else
                Items.Add CurrentItem
            End If
        Next RowIndex

    Else
        
        For Each CurrentItem In ArrayOrVector
            If IsUniqueOnly Then
                Items.Add CurrentItem, CStr(CurrentItem)
            Else
                Items.Add CurrentItem
            End If
        Next CurrentItem

    End If
    
    On Error GoTo 0
    
    ArrayToText = ConcatenateCollection(Items, Delimiter)

End Function

Public Function RepeatString(ByVal Text As String _
                             , Optional ByVal NumberOfTimes As Long = 1) As String
    
    Dim Result As String
    If Text = vbNullString Or NumberOfTimes <= 0 Then
        Result = vbNullString
    ElseIf NumberOfTimes = 1 Then
        Result = Text
    Else
        Result = Replace(Space(NumberOfTimes), Space(1), Text)
    End If
    
    RepeatString = Result
    
End Function

Public Function WrapWithDoubleQuote(ByVal InputText As String) As String
    
    WrapWithDoubleQuote = """" & Replace(InputText, """", """""") & """"
    
End Function

Public Function GetURLResponseText(ByVal URL As String _
                                   , Optional ByVal ContentTypeReqHeader As String = "application/json") As String
    
    '        Dim HTTPCaller As MSXML2.XMLHTTP60
    '        Set HTTPCaller = New MSXML2.XMLHTTP60
    '
    Dim HTTPCaller As Object
    Set HTTPCaller = CreateObject("MSXML2.XMLHTTP.6.0")
    
    Dim Result As String
    With HTTPCaller
        .Open "GET", URL, False
        .setRequestHeader "Content-Type", ContentTypeReqHeader
        .send
        Result = .responseText
    End With
    
    GetURLResponseText = Result
    
End Function

Public Function GetAccountFromEmailAddress(ByVal OutApp As Object _
                                    , ByVal EmailAddress As String) As Object
                                    
    Dim Account As Object
    
    ' Loop through the accounts in the Outlook session
    For Each Account In OutApp.Session.Accounts
        If Account.SmtpAddress = EmailAddress Then
            Set GetAccountFromEmailAddress = Account
            Exit Function
        End If
    Next Account
    
End Function

Public Function RoundUpToNextIntNumber(ByVal ForNumber As Double) As Long
    
    Dim Result As Long
    Debug.Print TypeName(ForNumber)
    If Not IsTextPresent(CStr(ForNumber), ".") Then
        Result = CLng(ForNumber)
    ElseIf ForNumber < 0 Then
        ' If negative number fix will transform -99.2 or -99.8 = -99
        ' But int will transform them to 100
        Result = Fix(ForNumber)
    Else
        ' In case of positive number int will extract only the int part. So Add + 1
        Result = Int(ForNumber) + 1
    End If
    
    RoundUpToNextIntNumber = Result
        
End Function

Public Function RoundDownToPreviousIntNumber(ByVal ForNumber As Double) As Long
    
    Dim Result As Long
    If Not IsTextPresent(CStr(ForNumber), ".") Then
        Result = CLng(ForNumber)
    ElseIf ForNumber < 0 Then
        ' If negative number fix will transform -99.2 or -99.8 = -99
        ' But int will transform them to 100
        Result = Int(ForNumber)
    Else
        ' In case of positive number int will extract only the int part. So Add + 1
        Result = Fix(ForNumber)
    End If
    
    RoundDownToPreviousIntNumber = Result
        
End Function

Public Function UnescapeXML(ByVal Text As String) As String
    
    '@Description: Not bullet proof. watch out for issue.
    
    Dim Result As String
    Result = Text
    Result = Replace(Result, "&amp;", "&")
    Result = Replace(Result, "&lt;", "<")
    Result = Replace(Result, "&gt;", ">")
    Result = Replace(Result, "&quot;", """")
    Result = Replace(Result, "&apos;", "'")
    
    UnescapeXML = Result

End Function

Public Function EscapeXML(ByVal Text As String) As String
    
    '@Description: Not bullet proof. watch out for issue.
    
    Dim Result As String
    Result = Text
    Result = Replace(Result, "&", "&amp;")
    Result = Replace(Result, "<", "&lt;")
    Result = Replace(Result, ">", "&gt;")
    Result = Replace(Result, """", "&quot;")
    Result = Replace(Result, "'", "&apos;")
    
    EscapeXML = Result

End Function

Public Function GetUserInputs(ByVal Prompt As String _
                              , ByVal InputType As InputBoxEnum _
                               , Optional ByVal Count As Integer = 1 _
                                , Optional ByVal DefaultValue As Variant) As Variant
    
    ' This function is for taking user input again and again.
    
    Dim UserInputs As Collection
    Dim Counter As Integer
    
    Set UserInputs = New Collection
    For Counter = 1 To Count
        
        Dim CurrentUserInput As Variant
        If InputType = 8 Then
            On Error Resume Next
            Set CurrentUserInput = Application.InputBox(Prompt, "User Input", DefaultValue, Type:=InputType)
            On Error GoTo 0
        Else
            CurrentUserInput = Application.InputBox(Prompt, "User Input", DefaultValue, Type:=InputType)
        End If
        
        ' As UserInput is variant type don't use Not UserInput
        If Not IsObject(CurrentUserInput) Then
            If CurrentUserInput = False Then Exit For
        End If
        
        UserInputs.Add CurrentUserInput
        
        CurrentUserInput = False
        
    Next Counter
    
    Dim Result As Variant
    If UserInputs.Count = 0 Then
        ' Update for default value for each data type
        
        Select Case InputType

            Case InputBoxEnum.IBFormula, InputBoxEnum.IBString
                Result = vbNullString
            
            Case InputBoxEnum.IBNumber
                Result = 0

            Case InputBoxEnum.IBBoolean
                Result = False

            Case InputBoxEnum.IBRange
                Set Result = Nothing

            Case InputBoxEnum.IBError
                Result = vbEmpty

            Case InputBoxEnum.IBArray
                Result = vbEmpty

            Case Else
                Err.Raise 13, "Wrong Input Argument"

        End Select
        
    ElseIf UserInputs.Count = 1 Then
        If IsObject(UserInputs.Item(1)) Then
            Set Result = UserInputs.Item(1)
        Else
            Result = UserInputs.Item(1)
        End If
    Else
        Result = CollectionToArray(UserInputs)
    End If
    
    If IsObject(Result) Then
        Set GetUserInputs = Result
    Else
        GetUserInputs = Result
    End If
    
End Function

Public Function GetResizedRange(ByVal TopCell As Range, ByVal Arr As Variant) As Range
    
    If IsNothing(TopCell) Then Exit Function
    
    Dim Result As Range
    If Is2DArray(Arr) Then
        Set Result = TopCell.Resize(NumberOfRowIn2DArray(Arr), NumberOfColumnIn2DArray(Arr))
    ElseIf IsArray(Arr) Then
        Set Result = TopCell.Resize(UBound(Arr) - LBound(Arr) + 1)
    Else
        Set Result = TopCell
    End If
    
    Set GetResizedRange = Result
    
End Function
