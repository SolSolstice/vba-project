Attribute VB_Name = "Module1"
'~MTC 2023
                                                                                                        


Sub randomizer()

'Range("B3:H20").ClearContents
'Range("A9:A24").ClearContents
Range("C6:C2000").Clear


Dim rand As Long



Dim userInput As Variant

Dim distrib_vals As Long


Dim i As Integer
Dim row As Long




Dim psu_dupeCount As Long
Dim psu_UniqueValues As Object
Dim psu_columnRange As Range

Dim p_cell As Range
Dim psu As Object
Dim psu_count As Double
Dim psu_cell As Range

Dim psu_countValues As Object




Dim duplicatePsu As String



Dim sub_count As Long
Dim sub_UniqueValues As Object
Dim s_cell As Range
Dim sub_columnRange As Range



' setting up variables & dictionaries to check for unique & duplicate names within Col A & Col B

Set sub_columnRange = Range("A6:A" & Cells(Rows.Count, "A").End(xlUp).row)
Set sub_UniqueValues = CreateObject("Scripting.Dictionary")

Set psu_countValues = CreateObject("Scripting.Dictionary")
Set psu_columnRange = Range("B6:B" & Cells(Rows.Count, "B").End(xlUp).row)
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

lastrow = Cells(Rows.Count, 1).End(xlUp).row


For Each s_cell In sub_columnRange
    If Not sub_UniqueValues.Exists(s_cell.Value) Then
        sub_UniqueValues.Add s_cell.Value, 1
    End If
Next s_cell
sub_count = sub_UniqueValues.Count


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

    
    
psu_count = psu.Count
    
MsgBox "total PSUs: " & psu_count




' checks if the user input matches the number of unique sub-district names

If userInput <> sub_count Then
    MsgBox "Error: Number of Sub-Districts entered does not equal number of Sub-Districts present"
    MsgBox "Number of unique Sub-Districts counted: " & sub_count
    MsgBox "Number of unique Sub-Districts entered: " & userInput
    Exit Sub
    End If



'the counter
distrib_vals = 15


   
While distrib_vals > 0
    rand = Int((psu_count * Rnd) + 1)
       If IsEmpty(Cells(5 + rand, 3).Value) Then
    Cells(5 + rand, 3).Value = "x"
    MsgBox rand
    
    ' ^^ maybe add a way to place an x & have a message box pop that says the name of the PSU?
    distrib_vals = distrib_vals - 1
    rand = rand + 1
    
    End If
    
Wend

' the above section is for randomly assigning PSUs to be sampled. an X is placed by each PSU that has been chosen.
    ' starts at row 6 and goes to the last row populated row in the spreadsheet (i.e., the last PSU)
    ' variable rand is any number between 1 and total amount of PSUs (psu_count)
        ' rand select a new number in that range iteration of the loop
    ' we use rand + 5 to randomly populate a cell in column C starting at row 6
    ' (if IsEmpty) is used so that a new cell is chosen if an x is already present



End Sub
