Attribute VB_Name = "mCommon"
Option Explicit

Public Function VariantArr(Optional ByVal Begin_ As Long, Optional ByVal End_ As Long)
    Dim ret()
    ReDim ret(Begin_ To End_)
    VariantArr = ret
End Function

Public Sub AppendToVariantArr(ByRef SourceArr, ByVal NewItem As Variant)
    Dim n As Long
    n = ArrUbound(SourceArr) + 1
    If n > 0 Then
        ReDim Preserve SourceArr(n)
        SourceArr(n) = NewItem
    Else
        Dim arr()
        ReDim arr(0)
        arr(0) = NewItem
        SourceArr = arr
    End If
End Sub

Public Function ArrUbound(ByVal SourceArr As Variant) As Long
    On Error GoTo eh
    ArrUbound = UBound(SourceArr)
    Exit Function
eh:
    Err.Clear
    ArrUbound = -1
End Function


