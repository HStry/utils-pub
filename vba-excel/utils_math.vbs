Attribute VB_Name = "utils_math"
Option Explicit

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