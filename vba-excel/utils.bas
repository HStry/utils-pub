Attribute VB_Name = "utils"
Public Function cDeriv() As String
    cDeriv = ChrW(8706)
End Function

Public Function Floor(val As Double) As Long
    Floor = CLng(Int(val))
End Function
Public Function Ceil(val As Double) As Long
    Ceil = Floor(val)
    If val > Ceil Then Ceil = Ceil + 1
End Function
Private Function AggregateWorker(arr As Variant) As Variant
    AggregateWorker = Array(0, Empty, Empty, 0, 0)
    
    If IsNumeric(arr) Then
        If Not IsEmpty(arr) Then AggregateWorker = Array(1, arr, arr, arr, arr ^ 2)
        Exit Function
    End If
    
    If Not IterableArray(arr) Then Exit Function
    
    Dim val As Variant
    Dim aval As Variant
    
    Dim n As Long
    Dim amin As Variant
    Dim amax As Variant
    Dim asum As Variant
    Dim asum2 As Variant
    
    n = 0
    asum = 0
    asum2 = 0
    
    For Each val In arr
        aval = AggregateWorker(val)
        If aval(0) > 0 Then
            n = n + aval(0)
            asum = asum + aval(3)
            asum2 = asum2 + aval(4)
            
            If IsEmpty(amin) Then
                amin = aval(1)
                amax = aval(2)
            Else
                If aval(1) < amin Then amin = aval(1)
                If aval(2) > amax Then amax = aval(2)
            End If
        End If
    Next val
    
    AggregateWorker = Array(n, amin, amax, asum, asum2)
End Function
Public Function Aggregate(ParamArray vals() As Variant) As Variant()
    Aggregate = AggregateWorker(Array(vals))
End Function
Public Function Min(ParamArray vals() As Variant) As Variant
    Min = AggregateWorker(Array(vals))(1)
End Function
Public Function Max(ParamArray vals() As Variant) As Variant
    Max = AggregateWorker(Array(vals))(2)
End Function
Public Function Sum(ParamArray vals() As Variant) As Variant
    Sum = AggregateWorker(Array(vals))(3)
End Function
Public Function Sum2(ParamArray vals() As Variant) As Variant
    Sum2 = AggregateWorker(Array(vals))(4)
End Function
Public Function Avg(ParamArray vals() As Variant) As Variant
    Dim a As Variant
    a = AggregateWorker(Array(vals))
    If a(0) > 0 Then Avg = a(3) / a(0)
End Function
Public Function Avg2(ParamArray vals() As Variant) As Variant
    Dim a As Variant
    a = AggregateWorker(Array(vals))
    If a(0) > 0 Then Avg2 = (a(4) / a(0)) ^ (1 / 2)
End Function

Public Function UnixTimeToSerial(timestamp As Double) As Double
    UnixTimeToSerial = 25569 + (timestamp / 86400)
End Function

Public Function EpochToSerial(timestamp As Double) As Double
    EpochToSerial = 25569 + (timestamp / 86400)
End Function
Public Function EpochToDate(timestamp As Double) As Long
    EpochToDate = Floor(EpochToSerial(timestamp))
End Function
Public Function EpochToTime(timestamp As Double) As Double
    EpochToTime = (timestamp Mod 86400) / 86400
End Function

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

Public Function HasListObject(rng As Range) As Boolean
    On Error GoTo Err
    HasListObject = rng.ListObject.Name <> ""
Err:
End Function
Public Function ListObjectRowNum(Optional rng As Range) As Long
    ' Returns the row number of rng within its listobject. Row 0 and -1 are
    ' the header and total row respectively.
    ListObjectRowNum = -9
    On Error GoTo Err
    Set rng = RangeSelector(rng, Application.Caller)
    
    Dim dbr As Range
    Set dbr = rng.ListObject.DataBodyRange
    
    ListObjectRowNum = rng.Row - dbr.Row + 1
    If ListObjectRowNum > dbr.Rows.Count Then ListObjectRowNum = -1
Err:
End Function
Public Function ListObjectNamedCol(rng As Range, colname As String) As Range
    On Error GoTo Err
    Set ListObjectNamedCol = Range(rng.ListObject.Name & "[" & colname & "]")
Err:
End Function

Public Function GetCnWorksheet(cname As String) As Worksheet
    Dim ws As Worksheet
    For Each ws In Application.Worksheets
        If ws.CodeName = cname Then
            Set GetCnWorksheet = ws
            Exit Function
        End If
    Next ws
End Function

Public Function GetListColumn(lo As ListObject, cname As String) As ListColumn
    Dim lc As ListColumn
    For Each lc In lo.ListColumns
        If lc.Name = cname Then
            Set GetListColumn = lc
            Exit Function
        End If
    Next lc
End Function
