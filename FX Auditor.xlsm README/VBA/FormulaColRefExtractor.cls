VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FormulaColRefExtractor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
Option Explicit


Private Type TFormulaColRefExtractor
    Formula As String
    UpdatedFormulaWithTrimRange As String
    ColRefBeanColl As Collection
    IsAlreadyTrimRangeApplied As Boolean
    IsAnyFullColRefFound As Boolean
End Type

Private This As TFormulaColRefExtractor

Public Property Get IsAnyFullColRefFound() As Boolean
    IsAnyFullColRefFound = This.IsAnyFullColRefFound
End Property

Public Property Get UpdatedFormulaWithTrimRange() As String
    UpdatedFormulaWithTrimRange = This.UpdatedFormulaWithTrimRange
End Property

Public Property Get ColRefBeanColl() As Collection
    Set ColRefBeanColl = This.ColRefBeanColl
End Property

Public Property Get IsAlreadyTrimRangeApplied() As Boolean
    IsAlreadyTrimRangeApplied = This.IsAlreadyTrimRangeApplied
End Property

Public Property Get Formula() As String
    Formula = This.Formula
End Property

Public Property Let Formula(ByVal RHS As String)
    This.Formula = RHS
    ParseFormula
End Property

Public Property Get Self() As FormulaColRefExtractor
    Set Self = Me
End Property

Public Function Create(Formula As String) As FormulaColRefExtractor

    With New FormulaColRefExtractor
        .Formula = Formula
        Set Create = .Self
    End With

End Function

Private Sub ParseFormula()
    
    ' Check if the formula contains any ':'
    If Not Text.Contains(This.Formula, COL_SEPERATOR) Then
        This.IsAnyFullColRefFound = False
        Exit Sub
    End If
    
    ' Extract from end so that we can easily replace them from the end.
    Dim ColonPositions As Collection
    Set ColonPositions = Text.FindAllIndexOf(Formula, COL_SEPERATOR, FROM_end)
    
    Dim Pos As Variant

    Dim NewFormula As String
    NewFormula = Formula
    
    Set This.ColRefBeanColl = New Collection
    
    ' Loop through positions and process each ':'
    For Each Pos In ColonPositions
        
        Dim CurrentColRefBean As ColRefBean
        Set CurrentColRefBean = GetCurrentColRefBeanFromColPos(This.Formula, CLng(Pos))
        
        If IsNotNothing(CurrentColRefBean) Then
            If Not CurrentColRefBean.IsTrimRangePresent Then
                Dim StartPos As Long
                StartPos = Pos - Len(Split(CurrentColRefBean.ColRef, COL_SEPERATOR)(0))
                NewFormula = Left(NewFormula, StartPos - 1) _
                             & "TRIMRANGE(" & CurrentColRefBean.ColRef & ", 2)" _
                             & Mid(NewFormula, StartPos + Len(CurrentColRefBean.ColRef))
            End If
            If This.ColRefBeanColl.Count = 0 Then
                This.ColRefBeanColl.Add CurrentColRefBean
            Else
                This.ColRefBeanColl.Add CurrentColRefBean, , 1
            End If
        End If
        
    Next Pos
    
    This.UpdatedFormulaWithTrimRange = NewFormula
    
    If This.ColRefBeanColl.Count = 0 Then
        This.IsAnyFullColRefFound = False
        This.IsAlreadyTrimRangeApplied = False
    Else
        This.IsAnyFullColRefFound = True
        This.IsAlreadyTrimRangeApplied = True
        For Each CurrentColRefBean In This.ColRefBeanColl
            If Not CurrentColRefBean.IsTrimRangePresent Then
                This.IsAlreadyTrimRangeApplied = False
                Exit For
            End If
        Next CurrentColRefBean
    End If
    
End Sub

Private Function GetCurrentColRefBeanFromColPos(ByVal Formula As String, ByVal ColPos As Long) As ColRefBean
    
    Dim CurrentColRefBean As ColRefBean
    Dim Ref As String
    ' Extract the column reference
    Ref = ExtractColumnRefWithoutSheetName(Formula, ColPos)
        
    ' Validate the column reference
    If IsValidColumnRef(Ref) Then
        ' Check if the reference is inside a quotation
        If Not IsInsideQuotation(Formula, ColPos) Then
                
            Ref = ExtractSheetNameQualifiedRef(Formula, Ref, ColPos)
            Dim StartPos As Long
            StartPos = ColPos - Len(Split(Ref, COL_SEPERATOR)(0))
            If Not Text.IsStartsWith(Ref, "#REF!", IGNORE_CASE) Then
                Set CurrentColRefBean = ColRefBean.Create(Ref, StartPos, IsAlreadyTrimRangeAdded(Formula, StartPos))
            End If
                
        End If
    End If
    
    Set GetCurrentColRefBeanFromColPos = CurrentColRefBean
    
End Function

Private Function ExtractColumnRefWithoutSheetName(ByVal Formula As String, ColonPos As Long) As String
    
    Dim StartPos As Long
    Dim EndPos As Long
    Dim Ref As String
    
    ' Check left part
    StartPos = ColonPos - 1
    Do While IsAlphabetOrDollarSign(Formula, StartPos)
        StartPos = StartPos - 1
    Loop
    
    StartPos = StartPos + 1
    
    ' Check right part
    EndPos = ColonPos + 1
    Do While IsAlphabetOrDollarSign(Formula, EndPos)
        EndPos = EndPos + 1
    Loop
    
    EndPos = EndPos - 1
    
    ' Extract the reference
    Ref = Mid(Formula, StartPos, EndPos - StartPos + 1)
    
    ExtractColumnRefWithoutSheetName = Ref
    
End Function

Private Function IsAlphabetOrDollarSign(ByVal Text As String, CheckAtIndex As Long) As Boolean
    
    Dim Result As Boolean
    If CheckAtIndex <= 0 Then
        Result = False
    ElseIf CheckAtIndex > Len(Text) Then
        Result = False
    Else
        
        Dim CurrentChar As String
        CurrentChar = Mid$(Text, CheckAtIndex, 1)
        Result = (CurrentChar Like "[A-Za-z$]")
    End If
    
    IsAlphabetOrDollarSign = Result
    
End Function

Private Function IsInsideQuotation(Formula As String, Pos As Long) As Boolean
    
    Dim i As Long
    Dim InQuotes As Boolean
    
    InQuotes = False
    For i = 1 To Pos
        If Mid(Formula, i, 1) = """" Then
            InQuotes = Not InQuotes
        End If
    Next i
    
    IsInsideQuotation = InQuotes
    
End Function

Private Function IsValidColumnRef(ByVal FullColAddress As String) As Boolean
    
    Dim Parts() As String
    Parts = Split(FullColAddress, ":")
    
    IsValidColumnRef = (IsPossibleColRef(CStr(Parts(0))) _
                        And IsPossibleColRef(CStr(Parts(1))))
    
End Function

Private Function IsPossibleColRef(ByVal ColumnOrdinal As String) As Boolean
    
    Dim Result As Boolean
    
    Dim WithoutStartDolSign As String
    WithoutStartDolSign = Text.RemoveFromStartIfPresent(ColumnOrdinal, "$")

    Select Case Len(WithoutStartDolSign)

        Case 1
            Result = (WithoutStartDolSign Like "[A-Za-z]")
    
        Case 2
            Result = (WithoutStartDolSign Like "[A-Za-z][A-Za-z]")
    
        Case 3
            Result = (WithoutStartDolSign Like "[A-Xa-x][A-Za-z][A-Za-z]")
    
        Case Else
            Result = False
            
    End Select

    If Result Then
        On Error GoTo NotValidRef
        Result = (Not Range(WithoutStartDolSign & "1") Is Nothing)
    End If

    IsPossibleColRef = Result
    Exit Function
    
NotValidRef:
    
End Function


Private Function ExtractSheetNameQualifiedRef(ByVal Formula As String _
                                              , ByVal ColRef As String _
                                               , ByVal ColonPos As Long) As String
    
    
    
    Dim StartPos As Long
    StartPos = ColonPos - Len(Split(ColRef, ":")(0))
    
    Dim EndPos As Long
    EndPos = ColonPos + Len(Split(ColRef, ":")(1))
    
    Dim Result As String
    Dim CurrentChar As String
    If Mid(Formula, StartPos - 1, 1) = "!" Then
        ' Sheet name present.
        ' Now we may have single quote or not. Based on the single quote it will change the behaviour
        
        If Mid(Formula, StartPos - 2, 1) = "'" Then
            
            StartPos = StartPos - 3
            
            CurrentChar = Mid(Formula, StartPos, 1)
            Dim IsStop As Boolean
            IsStop = (CurrentChar = "'" And Mid(Formula, StartPos - 1, 1) <> "'" And Mid(Formula, StartPos + 1, 1) <> "'")
            Do While Not IsStop
                StartPos = StartPos - 1
                CurrentChar = Mid(Formula, StartPos, 1)
                IsStop = (CurrentChar = "'" And Mid(Formula, StartPos - 1, 1) <> "'" And Mid(Formula, StartPos + 1, 1) <> "'")
            Loop
            
        Else
            
            StartPos = StartPos - 2
            CurrentChar = Mid(Formula, StartPos, 1)
            ' In case of #REF! error
            Do While (IsAlphanumericOnly(CurrentChar) Or CurrentChar = "#") And StartPos > 0
                StartPos = StartPos - 1
                CurrentChar = Mid(Formula, StartPos, 1)
            Loop
            
            StartPos = StartPos + 1
            
        End If
    
        Result = Mid(Formula, StartPos, EndPos - StartPos + 1)
        
    Else
        Result = ColRef
    End If
    
    ExtractSheetNameQualifiedRef = Result
    
End Function

Private Function IsAlreadyTrimRangeAdded(ByVal Formula As String, ByVal CellRefStartPos As Long) As Boolean
    
    Const TRIM_RANGE_WITH_PAREN As String = "TRIMRANGE("
    
    Dim Result As Boolean
    If CellRefStartPos < Len(TRIM_RANGE_WITH_PAREN) Then
        Result = False
    Else
        Result = Text.IsStartsWith(Mid(Formula, CellRefStartPos - Len(TRIM_RANGE_WITH_PAREN)), TRIM_RANGE_WITH_PAREN)
    End If
    
    IsAlreadyTrimRangeAdded = Result
    
End Function
