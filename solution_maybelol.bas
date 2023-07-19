Attribute VB_Name = "Module6"
Sub go_again()

Range("C6:C2000").Clear

Dim dataRange As Range
Dim subColumn As Range
Dim psuColumn As Range

Dim currentCell As Range

Dim binDictionary As Object 'Dictionary to store the bins
Dim outputRange As Range

Dim dataRow As Range

Dim subValue As String

Dim intRows As Integer
intRows = 5

lastRow = Cells(Rows.count, 1).End(xlUp).row

Set subColumn = Range("A6:A" & lastRow)
Set psuColumn = Range("B6:B" & lastRow)

Set dataRange = Range(subColumn.Cells(-4), subColumn.Cells(lastRow))
Set outputRange = Range("D6")


Dim rand As Long



Dim userInput As Variant

Dim distrib_vals As Long


Dim i As Long


Dim row As Integer







Dim psu_dupeCount As Long
Dim psu_UniqueValues As Object
Dim psu_columnRange As Range

Dim p_cell As Range
Dim psu As Object
Dim psu_count As Double
Dim psu_cell As Range

Dim psu_countValues As Object

Dim vArray As Variant





Dim duplicatePsu As String

Dim count As Long
Dim address As String
Dim subKey As Variant
Dim nextCell As Range



Dim sub_count As Long
Dim sub_Dict As Object
Dim s_cell As Range
Dim sub_columnRange As Range

Dim sub_Coll As Object

Dim count_loop As Long

Dim count_sub As Long
Dim count_sub2 As Long



count_loop = 0
count_sub = 0
count_sub2 = 0

While count_sub < 15

    For row = 6 To lastRow
        If row = lastRow Then
            count_sub2 = count_sub2 + 1
                ' used to count the iterations of the loop
            
    
        If Cells(row + 1, 1).Value <> Cells(row, 1).Value Then
        count_sub = count_sub + 1
        
        
        ' if count_sub2 > 1 -> rnd value & coinflip (?) to decide if it will assign
        
     End If
    End If
Next row


Wend








End Sub
