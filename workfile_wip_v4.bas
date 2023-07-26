Attribute VB_Name = "Module7"
'~MTC 2023
                                                                                                        


Sub random_sample2()

Range("C6:C1000").ClearContents
Range("D6:D1000").ClearContents
Range("E6:C1000").ClearContents
Range("F6:C1000").ClearContents

Dim dataRange As Range
Dim subColumn As Range
Dim psuColumn As Range

Dim currentCell As Range

Dim binDictionary As Object 'Dictionary to store the bins
Dim outputRange As Range

Dim dataRow As Range

Dim subValue As String

Dim intRows As Integer
'intRows = 5

lastRow = Cells(Rows.count, 1).End(xlUp).row

Set subColumn = Range("A6:A" & lastRow)
Set psuColumn = Range("B6:B" & lastRow)

Set dataRange = Range(subColumn.Cells(-4), subColumn.Cells(lastRow))
Set outputRange = Range("D6")


Dim rand As Long



Dim userInput As Variant

Dim distrib_vals As Long


Dim i As Integer
Dim row As Integer
Dim rows3 As Integer




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
Dim sub_UniqueValues As Object
Dim s_cell As Range
Dim sub_columnRange As Range





' setting up variables & dictionaries to check for unique & duplicate names within Col A & Col B

Set sub_columnRange = Range("A6:A" & Cells(Rows.count, "A").End(xlUp).row)
Set sub_UniqueValues = CreateObject("Scripting.Dictionary")

Set psu_countValues = CreateObject("Scripting.Dictionary")
Set psu_columnRange = Range("B6:B" & Cells(Rows.count, "B").End(xlUp).row)
Set psu = CreateObject("Scripting.Dictionary")
duplicatePsu = ""






'end user inputs the desired amount of sub-disctrict

userInput = InputBox("Enter number of Sub-districts")
    ' Validate the input
    
    If Not IsNumeric(userInput) Or userInput <= 0 Then
        MsgBox "Invalid input. Please enter a positive numeric value.", vbExclamation
    End If

Range("B4") = userInput
'Checking for unique Values

lastRow = Cells(Rows.count, 1).End(xlUp).row


For Each s_cell In sub_columnRange
    If Not sub_UniqueValues.Exists(s_cell.Value) Then
        sub_UniqueValues.Add s_cell.Value, 1
    End If
Next s_cell
sub_count = sub_UniqueValues.count


' checking for duplicate values
' message box pops up with duplicate names if dupe is present

For Each p_cell In psu_columnRange
    If psu.Exists(p_cell.Value) Then

        If InStr(1, duplicatePsu, p_cell.Value, vbTextCompare) = 0 Then
            duplicatePsu = duplicatePsu & p_cell.Value & ", "
            MsgBox "Duplicate PSU: " & duplicatePsu
            'colors dupe PSU if present
            
            p_cell.Interior.ColorIndex = 3
            'Exits sub when dupe is present - is it possible to highlight all the dupes THEN exit sub?
            
            Exit Sub
            
            
        End If
    Else
        psu.Add p_cell.Value, 1
    End If

Next p_cell


If Len(duplicatePsu) > 0 Then
    duplicatePsu = Left(duplicatePsu, Len(duplicatePsu) - 2)
End If
' ^^ uhhh.. i forget why i included this.. hmmm

    
    
psu_count = psu.count
    
'MsgBox "total PSUs: " & psu_count




' checks if the user input matches the number of unique sub-district names

If userInput <> sub_count Then
    MsgBox "Error: Number of Sub-Districts entered does not equal number of Sub-Districts present"
    MsgBox "Number of unique Sub-Districts counted: " & sub_count
    MsgBox "Number of unique Sub-Districts entered: " & userInput
    Exit Sub
    End If


' !! When handling 15 or less subdistricts


'the counter
Dim counter As Integer
Dim selector As Integer
Dim rand2 As Integer
Dim rand3 As Integer

 

distrib_vals = 15

Dim random_val As Variant

Dim rand_die As Long




counter = 1
Dim p As Long
Dim tracker As Long
tracker = 5

Dim assignment As Long


For row = 6 To lastRow
    
    If IsEmpty(Range("A6:A" & row)) Then
        Exit For
    MsgBox "Error: There is a blank row"
    End If
    
    If Range("A" & row).Value = Range("A" & row - 1).Value Then
    
        counter = counter + 1
        Else
        
        counter = 1
    End If
    
    'tracker =
    Range("D" & row).Value = counter
    
Next row

Dim loop_counter As Long
loop_counter = 0



If sub_count <= 15 And distrib_vals = 15 Then

    For row = 6 To lastRow

        'this for loop adds to & stores the values of: *selector* (col D), *rand* (col E) , & *assignment* (col F) for each sub-district
            
            If Range("A" & row + 1).Value <> Range("A" & row).Value Then

                selector = Cells(row, 4).Value
                        'hidden row - contains max value of counter due to <> argument in if/then conditional above
                    
                    Debug.Print "selector1: " & selector
        
                    'Debug.Print "tracker1: " & tracker
                
                rand = Int((selector * Rnd + 1))
                Range("E" & row).Value = rand
                    
                    Debug.Print "random: " & rand
                            ' rand = 1 to (max value of count)
                
                        ' used to toggle to a random row within the range  of PSUs for each subdistrict
                assignment = row + 1 - rand
                Range("F" & row).Value = assignment
                    
                If IsEmpty(Cells(assignment, 3).Value) Then
                    Cells(assignment, 3).Value = "x"
                    distrib_vals = distrib_vals - 1
                    loop_counter = loop_counter + 1
                            'counts how many times the <> loop has spotted a change in sub-district names
                    
                End If

            End If
            
    Next row
    

         

                
        
End If


            

            
If sub_count <= 15 And distrib_vals > 0 Then
            
    For intRows = 6 To lastRow
    
        
    
    
                                         'this for loop adds to & stores the values of: *selector* (col D), *rand* (col E) , & *assignment* (col F) for each sub-district
            
        If Range("A" & intRows + 1).Value <> Range("A" & intRows).Value Then
                
        
            selector = Cells(intRows, 4).Value
                        'hidden row - contains max value of counter due to <> argument in if/then conditional above
                    
                         Debug.Print "selector2: " & selector
        
                        'Debug.Print "tracker: " & tracker
                
                
            rand = Int((selector * Rnd + 1))
            Range("E" & intRows).Value = rand
        
                      Debug.Print "random2: " & rand
                    ' rand = 1 to (max value of count)
    
                    ' used to toggle to a random row within the range  of PSUs for each subdistrict
            assignment = intRows + 1 - rand
            Range("F" & intRows).Value = assignment
        
        
            
        
                If IsEmpty(Cells(selector + assignment, 3).Value) Then
                    Cells(selector + assignment, 3).Value = "x"
                    Debug.Print selector + assignment
                    distrib_vals = distrib_vals - 1
                    loop_counter = loop_counter + 1
                'counts how many times the <> loop has spotted a change in sub-district names
        
                End If
                
                If distrib_vals = 0 Then
                Exit For
                End If
            End If
            

            
            
            
                  

                    'Debug.Print loop_counter
        
        Next intRows
    'End If
    
    
    
Else









    

'    End If
    'End If

'end if

       
       
    
     



    
    
If sub_count > 15 And distrib_vals = 15 Then
    While distrib_vals > 0

        For rows3 = 6 To lastRow
                                   '        If Cells(row + 1, 4).Value < Cells(row, 4).Value Then
    

                rand = Int((psu_count * Rnd + 1))
                If IsEmpty(Cells(5 + rand, 3).Value) Then
        
                                     '        Debug.Print rand
                Cells(5 + rand, 3).Value = "x"
                distrib_vals = distrib_vals - 1
                    
                End If
                
                If distrib_vals = 0 Then Exit For
                
                
       '
            
        Next rows3
            'End If
    
        'Next rows3
            
            
    Wend
    
End If
'Stop



End If
Stop

End Sub





