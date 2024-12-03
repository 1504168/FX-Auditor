VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ColRefBean"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
Option Explicit

Private Type TColRefBean
    ColRef As String
    StartIndex As Long
    IsTrimRangePresent As Boolean
End Type

Private This As TColRefBean

Public Property Get ColRef() As String
    ColRef = This.ColRef
End Property

Public Property Let ColRef(ByVal RHS As String)
    This.ColRef = RHS
End Property

Public Property Get StartIndex() As Long
    StartIndex = This.StartIndex
End Property

Public Property Let StartIndex(ByVal RHS As Long)
    This.StartIndex = RHS
End Property

Public Property Get IsTrimRangePresent() As Boolean
    IsTrimRangePresent = This.IsTrimRangePresent
End Property

Public Property Let IsTrimRangePresent(ByVal RHS As Boolean)
    This.IsTrimRangePresent = RHS
End Property

Public Property Get Self() As ColRefBean
    Set Self = Me
End Property

Public Function Create(ByVal ColRef As String _
                       , ByVal StartIndex As Long _
                        , ByVal IsTrimRangePresent As Boolean) As ColRefBean

    With New ColRefBean
        .ColRef = ColRef
        .StartIndex = StartIndex
        .IsTrimRangePresent = IsTrimRangePresent
        Set Create = .Self
    End With

End Function



