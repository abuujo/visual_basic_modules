Attribute VB_Name = "AppendSheets"
' Made by Jonathan Abood - feel free to use in any way. Cheers.

Option Explicit
' This subroutine will use the general subroutine "Append" and append data_1, data_2 and data_3
Sub AppendSheets()

' Note that in this case, the master_sheet will have the following column array:
Dim SheetName As String
SheetName = "master_data"
Sheets.Add.Name = SheetName

' Add in worksheet names
With Worksheets(SheetName)
    .Range("A1").Value = "name"
    .Range("B1").Value = "age"
    .Range("C1").Value = "height"
    .Range("D1").Value = "birth_date"
    .Range("E1").Value = "colour"
End With

' Now we can mapp the three data sheets as such:
' master_Data   : A B C D E
' data_1        : A B C
' data_2        : A B   C
' data_3        : A     B C
' which will get us the correct mapping!


' Construct mapping array as needed:
' I use ind_0 as master_sheet index
Dim ColumnMap(3, 5) As Variant

' Master Sheet
ColumnMap(0, 0) = SheetName
ColumnMap(0, 1) = "A"
ColumnMap(0, 2) = "B"
ColumnMap(0, 3) = "C"
ColumnMap(0, 4) = "D"
ColumnMap(0, 5) = "E"

' data_1 | Note here the trick - since were essentially skipping over columns
' we can add in blank column data to fill in said columns
ColumnMap(1, 0) = "data_1"
ColumnMap(1, 1) = "A"
ColumnMap(1, 2) = "B"
ColumnMap(1, 3) = "C"
ColumnMap(1, 4) = "D" ' Blank Column
ColumnMap(1, 5) = "D" ' Blank Column

' data_2
ColumnMap(2, 0) = "data_2"
ColumnMap(2, 1) = "A"
ColumnMap(2, 2) = "B"
ColumnMap(2, 3) = "D" ' Blank Column
ColumnMap(2, 4) = "C"
ColumnMap(2, 5) = "D" ' Blank Column

' data_2
ColumnMap(3, 0) = "data_3"
ColumnMap(3, 1) = "D"
ColumnMap(3, 2) = "A" ' Blank Column
ColumnMap(3, 3) = "D" ' Blank Column
ColumnMap(3, 4) = "B"
ColumnMap(3, 5) = "C"

' Call the Append Code
Call Append(ColumnMap)

' Example with condition
Dim SheetName1 As String
SheetName1 = "master_data_filter"
Sheets.Add.Name = SheetName1

' Add in worksheet names
With Worksheets(SheetName1)
    .Range("A1").Value = "name"
    .Range("B1").Value = "age"
    .Range("C1").Value = "height"
    .Range("D1").Value = "birth_date"
    .Range("E1").Value = "colour"
End With

Dim ColumnMap1 As Variant
ColumnMap1 = ColumnMap
ColumnMap1(0, 0) = SheetName1

' Create Condition Array -> (filter_value,col_ind,row_ind)
' Can change into 2D array for multiple filters on multiple sheets - need to change handling inside Append tho
Dim Condition(2) As Variant
Condition(0) = "red"
Condition(1) = 3
Condition(2) = 3

Call Append(ColumnMap1, Condition)

End Sub


' This Subroutine appends a new sheet's data to the master sheet's data.
' Matched columns based on column mapping done.
'
' master_sheet  | Insert's the append_sheet below the master_sheet's last row.
' append_sheet(s)  <
'
' This can be applied to any tables that need to be appeneded in that way as long as the columns can be mapped
' Onto the mastersheets column.
Private Sub Append(ColumnMap As Variant, Optional ByVal Condition As Variant)
   'MsgBox (ColumnMap(0, 0) & " " & ColumnMap(1, 0))
    Dim row As Integer
    Dim LastRow As Long
    Dim LastRowMaster As Long
    For row = LBound(ColumnMap, 1) + 1 To UBound(ColumnMap, 1)
        
        ' If condition is applied
        If IsMissing(Condition) = False Then
            If row = Condition(2) Then
                ThisWorkbook.Worksheets(ColumnMap(row, 0)).Range("A1").AutoFilter Field:=Condition(1), Criteria1:=Condition(0)
            End If
        End If
        
    
        ' Target Source Data
        LastRow = ThisWorkbook.Worksheets(ColumnMap(row, 0)).Range("A" & ThisWorkbook.Worksheets(ColumnMap(row, 0)).Rows.Count).End(xlUp).row
        LastRowMaster = ThisWorkbook.Worksheets(ColumnMap(0, 0)).Range("A" & ThisWorkbook.Worksheets(ColumnMap(row, 0)).Rows.Count).End(xlUp).row + 1
        
        ' For each column - Copy the data to merge onto master sheet
        Dim col As Integer
        For col = LBound(ColumnMap, 2) To UBound(ColumnMap, 2) - 1
            
            ' Use Col2 to avoid headers - assumes headers are in place
            ThisWorkbook.Worksheets(ColumnMap(row, 0)).Range(ColumnMap(row, col + 1) & "2:" & ColumnMap(row, col + 1) & LastRow).Copy
            ThisWorkbook.Worksheets(ColumnMap(0, 0)).Range(ColumnMap(0, col + 1) & LastRowMaster).PasteSpecial Paste:=xlPasteValues
        Next col
    Next row
    
Cleanup:
    LastRow = 0
    LastRowMaster = 0
End Sub
