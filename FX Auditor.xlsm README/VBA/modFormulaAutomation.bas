Attribute VB_Name = "modFormulaAutomation"
Option Explicit

Public Const COL_SEPERATOR As String = ":"

Private Sub Test()
    
    Dim Extractor As FormulaColRefExtractor
    Set Extractor = FormulaColRefExtractor.Create(ActiveCell.Formula2)
    Dim NewFormula As String
    NewFormula = Extractor.UpdatedFormulaWithTrimRange
    Debug.Print NewFormula
    
End Sub

Private Sub TestReplaceFullColumnRefWithTrimRangeForSheet()
    
    ReplaceFullColumnRefWithTrimRangeForSheet ActiveSheet
    
    Dim V As Variant
    V = Selection.Formula2R1C1
    
End Sub

Public Sub ReplaceFullColumnRefWithTrimRangeForBook(ByVal OnBook As Workbook)
    
    Dim CurrentSheet As Worksheet
    For Each CurrentSheet In OnBook.Worksheets
        ReplaceFullColumnRefWithTrimRangeForSheet CurrentSheet
        DoEvents
        Logger.Log DEBUG_LOG, vbNewLine
    Next CurrentSheet
    
End Sub

Public Sub ExtractAndListFormulaInfo()
    
    Dim Result As Variant
    Result = ExtractFormulaInfo(ActiveSheet)
    
    Dim DumpRange As Range
    Set DumpRange = GetResizedRange(Sheet2.Range("C4"), Result)
    DumpRange.Value = Result
    DumpRange.AutoFilter
    AutoFitRangeCols DumpRange, 50
    
End Sub

Public Function ExtractFormulaInfo(ByVal FormulaSheet As Worksheet) As Variant
    
    Dim R1C1FormulaVsRangeMap As Scripting.Dictionary
    Set R1C1FormulaVsRangeMap = GetR1C1FormulaVsRangeMapFromSheet(FormulaSheet)
    
    Dim Result As Variant
    ReDim Result(1 To R1C1FormulaVsRangeMap.Count + 1, 1 To 5)
    
    Result(1, 1) = "Address"
    Result(1, 2) = "Formula"
    Result(1, 3) = "Any Full Col Ref"
    Result(1, 4) = "Is TRIMRANGE Applied"
    Result(1, 5) = "Updated Formula With TRIMRANGE"
    
    Dim CurrentExtractor As FormulaColRefExtractor
    Dim CurrentItem As Variant
    Dim TempRange As Range
    Dim Counter As Long
    Counter = 2
    For Each CurrentItem In R1C1FormulaVsRangeMap.Items
        Set TempRange = CurrentItem
        Result(Counter, 1) = TempRange.Address
        Result(Counter, 2) = "'" & TempRange.Cells(1).Formula2
        Set CurrentExtractor = FormulaColRefExtractor.Create(TempRange.Cells(1).Formula2)
        Result(Counter, 3) = CurrentExtractor.IsAnyFullColRefFound
        If CurrentExtractor.IsAnyFullColRefFound Then
            Result(Counter, 4) = CurrentExtractor.IsAlreadyTrimRangeApplied
            If Not CurrentExtractor.IsAlreadyTrimRangeApplied Then
                Result(Counter, 5) = "'" & CurrentExtractor.UpdatedFormulaWithTrimRange
            End If
        End If
        Counter = Counter + 1
    Next CurrentItem
    
    ExtractFormulaInfo = Result
    
End Function


Public Sub ReplaceFullColumnRefWithTrimRangeForSheet(ByVal FormulaSheet As Worksheet)
    
    ' In case of early binding
    Dim R1C1FormulaVsRangeMap As Scripting.Dictionary
    Set R1C1FormulaVsRangeMap = GetR1C1FormulaVsRangeMapFromSheet(FormulaSheet)
    
    Dim CurrentExtractor As FormulaColRefExtractor
    Dim CurrentItem As Variant
    Dim TempRange As Range
    For Each CurrentItem In R1C1FormulaVsRangeMap.Items
        
        Set TempRange = CurrentItem
        
        Dim OldFormula As String
        OldFormula = TempRange.Cells(1).Formula2
        
        Set CurrentExtractor = FormulaColRefExtractor.Create(OldFormula)
        
        If CurrentExtractor.IsAnyFullColRefFound And Not CurrentExtractor.IsAlreadyTrimRangeApplied Then
            Logger.Log DEBUG_LOG, "Formula will be changed on: " & TempRange.Address
            Logger.Log DEBUG_LOG, "Old Formula:     " & OldFormula
            Logger.Log DEBUG_LOG, "Updated formula: " & CurrentExtractor.UpdatedFormulaWithTrimRange
            Logger.Log DEBUG_LOG, vbNewLine
            TempRange.Cells(1).Formula2 = CurrentExtractor.UpdatedFormulaWithTrimRange
            TempRange.Formula2R1C1 = TempRange.Cells(1).Formula2R1C1
            DoEvents
        End If
        
    Next CurrentItem
    
End Sub

Private Function GetR1C1FormulaVsRangeMapFromSheet(ByVal FormulaSheet As Worksheet) As Scripting.Dictionary
    
    On Error GoTo HandleError
    Dim FormulaCells As Range
    Set FormulaCells = FormulaSheet.UsedRange.SpecialCells(xlCellTypeFormulas)
    
    Logger.Log DEBUG_LOG, "Sheet Name: " & FormulaSheet.Name
    Logger.Log DEBUG_LOG, "Formula Cells: " & FormulaCells.Address
    
    ' In case of late binding
'    Dim R1C1FormulaVsRangeMap As Object
'    Set R1C1FormulaVsRangeMap = CreateObject("Scripting.Dictionary")
    
    ' In case of early binding
    Dim R1C1FormulaVsRangeMap As Scripting.Dictionary
    Set R1C1FormulaVsRangeMap = New Scripting.Dictionary
    
    Dim CurrentArea As Range
    For Each CurrentArea In FormulaCells.Areas
        
        Dim CurrentAreaFormulas As Variant
        CurrentAreaFormulas = CurrentArea.Formula2R1C1
        
        If CurrentArea.Cells.CountLarge = 1 Then
            UpdateFormulaDic R1C1FormulaVsRangeMap, CurrentAreaFormulas, CurrentArea
        Else
            Dim RowIndex As Long
            For RowIndex = LBound(CurrentAreaFormulas, 1) To UBound(CurrentAreaFormulas, 1)
                Dim ColumnIndex As Long
                For ColumnIndex = LBound(CurrentAreaFormulas, 2) To UBound(CurrentAreaFormulas, 2)
                    UpdateFormulaDic R1C1FormulaVsRangeMap, CurrentAreaFormulas(RowIndex, ColumnIndex), CurrentArea.Cells(RowIndex, ColumnIndex)
                Next ColumnIndex
                DoEvents
            Next RowIndex
            
        End If
        
        DoEvents
        
    Next CurrentArea
    
    Set GetR1C1FormulaVsRangeMapFromSheet = R1C1FormulaVsRangeMap
    Exit Function
    
HandleError:
    Logger.Log DEBUG_LOG, Err.Description
    Set GetR1C1FormulaVsRangeMapFromSheet = New Scripting.Dictionary
    
    
End Function

Private Sub UpdateFormulaDic(ByRef R1C1FormulaVsRangeMap As Scripting.Dictionary _
                             , ByVal R1C1Formula As String, ByVal ForCell As Range)
    
    If R1C1FormulaVsRangeMap.Exists(R1C1Formula) Then
        Set R1C1FormulaVsRangeMap.Item(R1C1Formula) = Union(R1C1FormulaVsRangeMap.Item(R1C1Formula), ForCell)
    Else
        R1C1FormulaVsRangeMap.Add R1C1Formula, ForCell
    End If
    
End Sub

Public Function GetSpecialCharactersForSheetName() As Collection
    
    
    Dim SpecialCharacters As Collection
    Set SpecialCharacters = New Collection
    Dim Characters As Variant
    Characters = Array("`", "~", "!", "@", "#", "$", "%", "^", "&", "(", ")", "-", "_", "=", "+", "{", "}", "|", ";", ":", ",", "<", ".", ">")
     
    Dim CurrentCharacter As Variant
    For Each CurrentCharacter In Characters
        SpecialCharacters.Add CurrentCharacter, CStr(CurrentCharacter)
    Next CurrentCharacter
    
    Set GetSpecialCharactersForSheetName = SpecialCharacters
    
End Function
