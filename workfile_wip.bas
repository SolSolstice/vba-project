Attribute VB_Name = "Module1"
'~MTC 2023
                                                                                                        


Sub randomizer()

'Range("B3:H20").ClearContents
'Range("A9:A24").ClearContents

Dim rand As Long
Dim remainder As Long
Dim exactNum As Long
Dim allocatedNum As Double
Dim allocatedNumMod As Long
Dim userInput As Variant

Dim vArray As Variant

Dim userInputDists() As Variant

Dim distrib_vals As Long

distrib_vals = 15



Dim i As Integer
Dim row As Long

Dim sub_lastrow As Long

Dim psu_avail As Long


Dim psu_dupeCount As Long
Dim psu_UniqueValues As Object
Dim psu_columnRange As Range

Dim p_cell As Range
Dim psu As Object
Dim psu_count As Double
Dim psu_cell As Range

Dim psuRows As Double
psuRows = 6
Dim psu_countValues As Object
Dim rand_assign As Long





Dim duplicatePsu As String



Dim sub_count As Long
Dim sub_UniqueValues As Object
Dim s_cell As Range
Dim sub_columnRange As Range






Set sub_columnRange = Range("A6:A" & Cells(Rows.Count, "A").End(xlUp).row)
Set sub_UniqueValues = CreateObject("Scripting.Dictionary")

Set psu_countValues = CreateObject("Scripting.Dictionary")
Set psu_columnRange = Range("B6:B" & Cells(Rows.Count, "B").End(xlUp).row)
Set psu = CreateObject("Scripting.Dictionary")
duplicatePsu = ""



'end user inputs the desired amount of sub-disctrict samples

userInput = InputBox("Enter number of Sub-districts")
    ' Validate the input
    If Not IsNumeric(userInput) Or userInput <= 0 Then
        MsgBox "Invalid input. Please enter a positive numeric value.", vbExclamation
        Exit Sub
    End If


'Checking for unique Values

lastrow = Cells(Rows.Count, 1).End(xlUp).row


For Each s_cell In sub_columnRange
    If Not sub_UniqueValues.Exists(s_cell.Value) Then
        sub_UniqueValues.Add s_cell.Value, 1
    End If
Next s_cell
sub_count = sub_UniqueValues.Count

' checking for duplicate values
' message box pops up with duplicate names if duplicate is present

For Each p_cell In psu_columnRange
    If psu.Exists(p_cell.Value) Then

        If InStr(1, duplicatePsu, p_cell.Value, vbTextCompare) = 0 Then
            duplicatePsu = duplicatePsu & p_cell.Value & ", "
            MsgBox "Duplicate PSU: " & duplicatePsu
        End If
    Else
        psu.Add p_cell.Value, 1
    End If

    
psu_dupeCount = p_cell.Count

Next p_cell


 ' !! ~~ trying to count unique values again
 
'For Each psu_cell In psu_columnRange
'    If Not psu_countValues.Exists(psu_cell.Value) Then
'        psu_countValues.Add psu_cell.Value, 1
'    End If
'Next psu_cell
'psu_avail = psu_countValues.Count
'MsgBox psu_avail

' ~ !! Trying to find unique PSU value count


If Len(duplicatePsu) > 0 Then
    duplicatePsu = Left(duplicatePsu, Len(duplicatePsu) - 2)
End If
    
psu_count = psu.Count

'While remainder > 0
'    If i > userInput Then i = i - userInput
'    vArray(i) = vArray(i) + 1
'    remainder = remainder - 1
'    i = i + rand
'Wend

'While distrib_vals > 0
'For i = 1 To psu_count
'    rand = Int((psu_count * Rnd) + 1) 'number between 1 and all PSUs
'    Cells(rand + 5, 3).Value = "x" '
'    rand = rand + 1
'Next i

    
    
MsgBox "total PSUs: " & psu_count




Range("B4") = userInput

ReDim userInputDists(1 To userInput)

If userInput <> sub_count Then
    MsgBox "Error: Number of Sub-Districts entered does not equal number of Sub-Districts present"
    MsgBox "Number of unique Sub-Districts counted: " & sub_count
    MsgBox "Number of unique Sub-Districts entered: " & userInput
    End If







'For row = 6 To psu_columnRange
psu_avail = 0
 
distrib_vals = 15
    
For row = 6 To lastrow
    If Cells(row + 1, 1).Value <> Cells(row, 1).Value Then
        psuRows = row
        rand = Int((psuRows * Rnd) + 1)

        'use rand value to place an X on a SINGLE cell that will be between the first and last PSU
        
        'For Each rand_assign In Range("c6:c19")
        '        rand_assign.Value = "x"
                
       
     '   MsgBox rand_assign
     
        
        MsgBox psuRows
        psuRows = psuRows + 1
        
   ' Next rand_assign


    
    
        'rand = Int((psu_avail * Rnd) + 1)
        'Range("h5").Value = rand
        'psuRows = psuRows + 1
    'Else
    
    'If Cells(row + 1, 1).Value = Cells(row, 1).Value Then
    '    psu_avail = psu_avail + 1
        
    '    End If
        
 
   ' End If
End If
Next row


distrib_vals = 15

    While distrib_vals > 0
        rand = Int((psu_count * Rnd) + 1)
        Cells(i * rand + 5, 3).Value = "x"
        distrib_vals = distrib_vals - 1
        rand = rand + 1
    Wend
    



'While remainder > 0
'    If i > userInput Then i = i - userInput
'    vArray(i) = vArray(i) + 1
'    remainder = remainder - 1
'    i = i + rand
'Wend


    
    
            

'While distrib_vals > 0
'    If i > userInput Then i = i - userInput
'    vArray(i) = vArray(i) + 1
'    remainder = remainder - 1
'    i = i + rand
'Wend
        
        
    
'Next i

    
    
'Next row

    
        
    
'rand is any value between 1 and remainder
' ex.) 15/6 = 2 remainder 3..  TF; rand = 1 - 3

    


'For row = 9 To lastrow
'    Cells(lastrow, "c9") = x
'Next row
    

    
    

  
    
   ' For i = 1 To userInput
   '     Cells(i + 8, 1).Value = userInputDists(i)
    
   ' Next i


'using remainder to allocate left over values into random sub-districts
    'remainder = 15 Mod userInput
'Range("F4") = remainder

'    allocatedNum = 15 / userInput
'round PSU sample values
 '   allocatedNumMod = Fix(allocatedNum)
'Range("G4") = allocatedNumMod


    


Range("H4") = rand

'ReDim vArray(1 To userInput)

'For subdist = LBound(vArray) To UBound(vArray)
'    vArray(subdist) = allocatedNumMod
'Next subdist

'i = rand

'While distrib_vals > 0
'    If i > userInput Then i = i - userInput
'    vArray(i) = vArray(i) + 1
'    remainder = remainder - 1
'    i = i + rand
'Wend


'Range("b9").Resize(UBound(vArray)).Value = Application.Transpose(vArray)


'Wend


End Sub
