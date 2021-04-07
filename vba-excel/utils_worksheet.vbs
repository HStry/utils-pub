Attribute VB_Name = "utils_worksheet"
Option Explicit

Public Function GetCodeNameWorksheet(wscodename As String, _
                                     Optional match_like As Boolean) _
                                     As Worksheet
    Dim ws As Worksheet
    For Each ws In Application.Worksheets
        If ws.CodeName = wscodename Or _
           (match_like And ws.CodeName Like wscodename) Then
            Set GetCodeNameWorksheet = ws
            Exit Function
        End If
    Next ws
End Function

Public Function GetListObject(ws As Worksheet, loname As String, _
                              Optional match_like As Boolean) _
                              As ListObject
    Dim lo As ListObject
    For Each lo In ws.ListObjects
        If lo.Name = loname Or _
           (match_like And lo.Name Like loname) Then
            Set GetListObject = lo
            Exit Function
        End If
    Next lo
End Function

Public Function GetListColumn(lo As ListObject, colname As String, _
                              Optional match_like As Boolean) _
                              As ListColumn
    Dim lc As ListColumn
    For Each lc In lo.ListColumns
        If lc.Name = colname Or _
           (match_like And lc.Name Like colname) Then
            Set GetListColumn = lc
            Exit Function
        End If
    Next lc
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