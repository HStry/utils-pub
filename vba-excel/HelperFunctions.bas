Option Explicit

'This file provides the following functions:
' getWorksheetFromCodeName
' getCellName
' tblRowNum
' tblRangeLookup
' tblValueLookup
' tblValueIntersectLookup
' AnyInListObject
' AllInListObject
' AllInSameListObject
' CountInListObject
' getListObjectName
' getCellBGColor

Public Function getWorksheetFromCodeName(codename As String) As Worksheet
  Dim ws As Worksheet
  
  'Runs through all worksheets and checks if its codename is equal to the
  'provided codename. Will return matching worksheet, and exit function if
  'found.
  For Each ws In ThisWorkbook.Worksheets
    If ws.codename = codename Then
      Set getWorksheetFromCodeName = ws
      Exit Function
    End If
  Next ws
End Function

Public Function getCellName(cell As Range) As String
  On Error GoTo reject
  If cell.Count = 1 Then
    getCellName = cell.Name.Name
    Exit Function
  End If
reject:
  getCellName = ""
End Function

Public Function tblRowNum(Optional r As Range) As Long
  On Error GoTo reject
  If (r Is Nothing) Then Set r = Application.Caller
  tblRowNum = r.Row - r.ListObject.DataBodyRange.Row + 1
  Exit Function
reject:
  tblRowNum = -1
End Function

Public Function tblRangeLookup( _
                  table As Range, _
                  lookup_column As Range, _
                  lookup_value As Variant, _
                  Optional operator As String = "=" _
                ) As Range
  Dim result As Range
  Dim row As Range
  
  For Each row In table.Rows()
    cell_value = Application.Intersect(row, lookup_column).Value
    If operator = "<>" And cell_value <> lookup_value Or _
       operator = "<" And cell_value < lookup_value Or _
       operator = "<=" And cell_value <= lookup_value Or _
       operator = "=" And cell_value = lookup_value Or _
       operator = ">=" And cell_value >= lookup_value Or _
       operator = ">" And cell_value > lookup_value Then
      If result Is Nothing Then
        Set result = row
      Else
        Set result = Union(result, row)
      End If
    End If
  Next row
  Set tblRangeLookup = result
End Function

Public Function tblValueLookup( _
                  result_column As Range, _
                  lookup_column As Range, _
                  lookup_value As Variant, _
                  Optional operator As String = "=" _
                ) As Variant
  Dim cell As Range
  
  For Each cell In lookup_column
    If operator = "<>" And cell.Value <> lookup_value Or _
       operator = "<" And cell.Value < lookup_value Or _
       operator = "<=" And cell.Value <= lookup_value Or _
       operator = "=" And cell.Value = lookup_value Or _
       operator = ">=" And cell.Value >= lookup_value Or _
       operator = ">" And cell.Value > lookup_value Then
      tblValueLookup = Application.Intersect(cell.EntireRow, result_column).Value
      Exit Function
    End If
  Next cell
  tblValueLookup = CVErr(xlErrNA)
End Function

Public Function tblValueIntersectLookup( _
                  table As Range, _
                  result_column As Range, _
                  lookup_column As Range, _
                  lookup_value As Variant, _
                  Optional operator As String = "=" _
                ) As Variant
  Dim rc As Range
  Dim lc As Range
  
  Set rc = Application.Intersect(table, result_column)
  Set lc = Application.Intersect(table, lookup_column)
  
  tblValueIntersectLookup = tblValueLookup(rc, lc, lookup_value, operator)
End Function

Public Function AnyInListObject(r As Range) As Boolean
  Dim c As Range
  For Each c In r
    If (Not c.ListObject Is Nothing) = True Then
      AnyInListObject = True
      Exit Function
    End If
  Next c
  AnyInListObject = False
End Function

Public Function AllInListObject(r As Range) As Boolean
  Dim c As Range
  For Each c In r
    If (Not c.ListObject Is Nothing) = False Then
      AllInListObject = False
      Exit Function
    End If
  Next c
  AllInListObject = True
End Function

Public Function AllInSameListObject( _
                  r As Range,  _
                  Optional IgnoreCellsNotInListObject As Boolean = True _
                ) As Boolean
  Dim c As Range
  Dim i As Long: i = 0
  Dim loname As String: loname = ""
  
  For Each c In r
    If (Not c.ListObject Is Nothing) = True Then
      If loname = "" Then
        loname = c.ListObject.Name
      ElseIf loname <> c.ListObject.Name Then
        AllInSameListObject = False
        Exit Function
      End If
    ElseIf IgnoreCellsNotInListObject = False Then
      AllInSameListObject = False
      Exit Function
    End If
  Next c
  If loname = "" Then
    AllInSameListObject = False
  Else
    AllInSameListObject = True
  End If
End Function

Public Function CountInListObject(r As Range) As Long
  Dim c As Range
  Dim n As Long: n = 0
  
  For Each c In r
    If (Not c.ListObject Is Nothing) = True Then n = n + 1
  Next c
  CountInListObject = n
End Function

Public Function getListObjectName(r As Range) As String
  If (Not r.ListObject Is Nothing) = True Then
    getListObjectName = r.ListObject.Name
  Else
    getListObjectName = ""
  End If
End Function

Public Function getCellBGColor(r As Range) As Variant
  Dim color As Long
  Dim c As Range
  Dim i As Long: i = 0
  For Each c In r
    If i = 0 Then
      color = c.Interior.color
    ElseIf color <> c.Interior.color Then
      getCellBGColor = CVErr(xlErrNA)
      Exit Function
    End If
    i = i + 1
  Next c
  getCellBGColor = color
End Function

Public Function arrDim(arr As Variant) As Variant
  On Error GoTo finish
  
  If Not IsArray(arr) Then
    arrDim = CVErr(xlErrValue)
    Exit Function
  End If
  
  Dim i As Long
  Dim u As Long
  
  i = 1
  Do While True
    u = UBound(arr, i)
    i = i + 1
  Loop
finish:
  arrDim = i - 1
End Function

Public Function arrLen(arr As Variant, Optional dimension As Long = 0) As Variant
  ' If dimension = -1, arrLen will return a 1D array containing the lengths for
  ' each dimension. If arrLen is 0, it will return the sum of each dimension's
  ' length. Otherwise it will return the length of the requested dimension.
  ' If the array is undefined, it will return an empty array if dimension = -1,
  ' or -1 if dimension is >= 0
  
  Dim dimensions As Long
  Dim ilen As Long
  Dim len_ As Long
  Dim alen() As Long
  Dim i As Long
  
  ' Check if provided array is actually an array
  If Not IsArray(arr) Then
    arrLen = CVErr(xlErrValue)
    Exit Function
  End If
  
  ' Determine number of dimensions in array
  dimensions = arrDim(arr)
  
  ' Check if requested dimension is valid.
  If dimension < -1 Or dimension > dimensions Then
    arrLen = CVErr(xlErrNA)
    Exit Function
  End If
  
  ' Calculate the array's length
  If dimension = -1 And dimensions = 0 Then
    'do nothing, not necessary
  ElseIf dimension = -1 Then
    ReDim alen(1 To dimensions)
    For i = 1 To dimensions
      alen(i) = UBound(arr, i) - LBound(arr, i) + 1
    Next i
  ElseIf dimensions = 0 Then
    ilen = -1
  ElseIf dimension = 0 Then
    ilen = 1
    For i = 1 To dimensions
      ilen = ilen * (UBound(arr, i) - LBound(arr, i) + 1)
    Next i
  Else
    ilen = UBound(arr, dimension) - LBound(arr, dimension) + 1
  End If
  
  ' Set arrLen
  If dimension = -1 Then
    arrLen = alen
  Else
    arrLen = ilen
  End If
  
End Function


Public Function arrIndex(arr As Variant, val As Variant, Optional lock As Variant = -1) As Variant
  
  ' arr must be an array
  If Not IsArray(arr) Then
    arrIndex = CVErr(xlErrValue)
    Exit Function
  End If

  Dim dimensions As Long
  Dim struct() As Long
  Dim coords() As Variant
  Dim coord() As Long
  dim i As Long
  
  ' Determine number of dimensions in array
  dimensions = arrDim(arr)
  
  ' An empty array will not contain the value.
  If dimensions = 0 Then
    arrIndex = coord
    Exit Function
  End If
  
  ' Get the size structure
  struct = arrLen(arr, -1)
  
  ' Set up the coordinate structure
  
  
  i = 0
  
  For Each v in arr
    
    
    i += 0
  Next v
  If dimensions = 0 Then
    
  
  
  For i = 1 To dimensions
    For j = LBound(arr, i) To UBound(arr, i)
      
    Next j
  Next i
  
  
  If fixed_coordinates = -1 Then
    'do stuff
    
  End If
  Dim alen() As Long
  alen = arrLen(val, -1)




Public Function arrAppend(arr As Variant, val As Variant) As Variant
  On Error GoTo reject
  ReDim Preserve arr(UBound(arr)+1)
  arr(UBound(arr)) = val
  arrAppend = arr
  Exit Function
reject:
  arrAppend = CVErr(xlErrValue)
End Function

