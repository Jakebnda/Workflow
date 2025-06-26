Option Explicit

Public Const COL_WO As String = "WO"

' Returns the ordered list of workflow stages.
Public Function GetStageNames() As Variant
    GetStageNames = Array("Design", "Printing", "Production", "Shipping")
End Function

