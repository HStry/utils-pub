Attribute VB_Name = "utils_variables"
Option Explicit

Public Function IterableArray(arr As Variant) As Boolean
    On Error GoTo Err
    IterableArray = UBound(arr) >= LBound(arr)
Err:
End Function

Private Function HasEmptyWorker(val As Variant) As Boolean
    If IterableArray(val) Then
        Dim ival As Variant
        For Each ival In val
            If HasEmptyWorker(ival) Then
                HasEmptyWorker = True
                Exit Function
            End If
        Next ival
    ElseIf TypeName(val) = "Range" Then
        Dim icell As Range
        For Each icell In val.Cells
            If IsEmpty(icell) Then
                HasEmptyWorker = True
                Exit Function
            End If
        Next icell
    Else
        HasEmptyWorker = IsEmpty(val)
    End If
End Function
Public Function HasEmpty(ParamArray vals() As Variant) As Boolean
    HasEmpty = HasEmptyWorker(Array(vals))
End Function

Private Function AllEmptyWorker(val As Variant) As Boolean
    If IterableArray(val) Then
        Dim ival As Variant
        For Each ival In val
            If Not AllEmptyWorker(ival) Then Exit Function
        Next ival
        AllEmptyWorker = True
    ElseIf TypeName(val) = "Range" Then
        Dim icell As Range
        For Each icell In val.Cells
            If Not IsEmpty(icell) Then Exit Function
        Next icell
        AllEmptyWorker = True
    Else
        AllEmptyWorker = IsEmpty(val)
    End If
End Function
Public Function AllEmpty(ParamArray vals() As Variant) As Boolean
    AllEmpty = AllEmptyWorker(Array(vals))
End Function

Public Function RangeSelector(ParamArray rngs() As Variant) As Range
    Dim rng As Variant
    For Each rng In rngs
        If TypeName(rng) = "Range" Then
            If Not rng Is Nothing Then
                Set RangeSelector = rng
                Exit Function
            End If
        End If
    Next rng
End Function
