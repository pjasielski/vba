Attribute VB_Name = "VbaSharedRepo"
' VBA Shared Repo
' Version 23.09.2019


Function LastRow(StartCell As Range) As Long
    'Get the number of last filled row
    Dim StartRow As Long

    StartRow = StartCell.Row

    LastRow = StartCell.Offset(Rows.Count - StartRow, 0).End(xlUp).Row

End Function


Function DataRows(StartCell As Range) As Long
    'Get the number of last filled row in data table
    Dim StartRow As Long

    StartRow = StartCell.Row

    DataRows = StartCell.Offset(Rows.Count - StartRow, 0).End(xlUp).Row - StartRow

End Function


Function LastCol(StartCell As Range) As Long
    
    StartCol = StartCell.Column

    LastCol = StartCell.Offset(0, Columns.Count - StartCol).End(xlToLeft).Column

End Function


Function DataCols(StartCell As Range) As Long

    StartCol = StartCell.Column

    DataCols = StartCell.Offset(0, Columns.Count - StartCol).End(xlToLeft).Column - StartCol

End Function


Function SortDataVertically(StartCell As Range, TargetCell As Range) As Long

    Dim i As Long
    Dim nextcol As Long
    Dim nextrow As Long

    StartRow = StartCell.Row
    nextcol = 0
    nextrow = 1

    For i = StartRow - 1 To DataRows(StartCell)
        If IsNumeric(StartCell.Offset(i, 0)) = False And StartCell.Offset(i, 0) <> "" Then
                TargetCell.Offset(0, nextcol) = StartCell.Offset(i, 0)
                nextcol = nextcol + 1
                nextrow = 1
        Else
            TargetCell.Offset(nextrow, nextcol - 1) = StartCell.Offset(i, 0)
            nextrow = nextrow + 1
        End If
    Next i

    SortDataVertically = i

End Function


Function SortDataHorizontally(StartCell As Range, TargetCell As Range) As Long

    Dim i As Long
    Dim nextcol As Long
    Dim nextrow As Long

    StartRow = StartCell.Row
    nextcol = 0
    nextrow = 0

    For i = StartRow - 1 To DataRows(StartCell)
        If IsNumeric(StartCell.Offset(i, 0)) = False And StartCell.Offset(i, 0) <> "" Then
                TargetCell.Offset(nextrow, 0) = StartCell.Offset(i, 0)
                nextrow = nextrow + 1
                nextcol = 1
        Else
            TargetCell.Offset(nextrow - 1, nextcol) = StartCell.Offset(i, 0)
            nextcol = nextcol + 1
        End If
    Next i

    SortDataHorizontally = i

End Function


'========================

Sub CombineArrays(Arr1 As Variant, Arr2 As Variant)

    Dim d As Scripting.Dictionary
    Set d = New Scripting.Dictionary
    Dim c As clsID
    Dim vKey As Variant
    
    'Load Arr1 to dictionary:
    For i = 1 To UBound(Arr1)
        Set c = New clsID
            c.ID = Arr1(i, 1)
            c.FirstName = Arr1(i, 2)
        
            vKey = c.ID
            If Not d.Exists(vKey) Then
                d.Add vKey, c
            End If
        Set c = Nothing
    Next i
        
    'Load Arr2 to dictionary:
    For i = 1 To UBound(Arr2)
        vKey = Arr2(i, 1)
        If d.Exists(vKey) Then
            Set c = d.Item(vKey)
                c.LastName = Arr2(i, 2)
        Else
            Set c = New clsID
                c.ID = Arr2(i, 1)
                c.LastName = Arr2(i, 2)
                
                d.Add vKey, c
        End If
        Set c = Nothing
    Next i
        
   'For Each vKey In d.Keys
   'do stuff

End Sub


