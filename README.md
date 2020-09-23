<div align="center">

## Collection Helper/Sorter


</div>

### Description

This method enables you to pass any collection 'byref' and then sort it by any property either ascending or descending.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Paul Gorman](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/paul-gorman.md)
**Level**          |Advanced
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/paul-gorman-collection-helper-sorter__1-36320/archive/master.zip)





### Source Code

```
Public Enum Enum_colSortType
  colSortTypeString = 0
  colSortTypeNumeric = 1
  colSortTypeDateTime = 2
End Enum
Public Enum Enum_colSortOrder
  colSortOrderAscending = 0
  colSortOrderDescending = 1
End Enum
Public Sub SortCollection(ByRef oCollection As Collection, ByVal SortPropertyName As String, Optional ByVal KeyPropertyName As String, Optional ByVal SortType As Enum_colSortType = colSortTypeString, Optional ByVal SortOrder As Enum_colSortOrder = colSortOrderAscending)
Dim RS As ADODB.Recordset, oObj As Object, i As Long
Dim oSorted As Collection, sKeyField As String
  Set RS = New ADODB.Recordset
  RS.CursorLocation = adUseClient
  RS.CursorType = adOpenStatic
  Select Case SortType
    Case colSortTypeString
      RS.Fields.Append SortPropertyName, adVarChar, 100
    Case colSortTypeNumeric
      RS.Fields.Append SortPropertyName, adDouble
    Case colSortTypeDateTime
      RS.Fields.Append SortPropertyName, adDate
  End Select
  If KeyPropertyName <> "" Then
    sKeyField = "Key" & KeyPropertyName
    RS.Fields.Append sKeyField, adVarChar, 100
  End If
  Set oSorted = New Collection
  RS.Open
  For i = oCollection.Count To 1 Step -1
    Set oObj = oCollection.Item(i)
    RS.AddNew
    RS.Fields(SortPropertyName).Value = CallByName(oObj, SortPropertyName, VbGet)
    If KeyPropertyName <> "" Then
      RS.Fields(sKeyField).Value = CallByName(oObj, KeyPropertyName, VbGet)
    End If
    RS.Update
    If KeyPropertyName <> "" Then
      oSorted.Add oObj, CallByName(oObj, KeyPropertyName, VbGet)
    Else
      oSorted.Add oObj
    End If
    oCollection.Remove i
  Next
  If SortOrder = colSortOrderAscending Then
    RS.Sort = SortPropertyName & " ASC"
  Else
    RS.Sort = SortPropertyName & " DESC"
  End If
  RS.MoveFirst
  i = 1
  Do Until RS.EOF
    If KeyPropertyName <> "" Then
      oCollection.Add oSorted.Item(RS.Fields(sKeyField).Value), RS.Fields(sKeyField).Value
    Else
      oCollection.Add oSorted.Item(i)
    End If
    RS.MoveNext
    i = i + 1
  Loop
End Sub
```

