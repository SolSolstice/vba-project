Attribute VB_Name = "Module4"
'~MTC 2023
                                                                                                        


Sub random_sample()

'Range("B3:H20").ClearContents
'Range("A9:A24").ClearContents
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


Dim i As Integer
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


counter = 1
Dim p As Long


For row = 6 To lastRow
    If IsEmpty(Range("A6:A" & row)) Then
        Exit For
    MsgBox "Error: There is a blank row)"
    End If
    
    If Range("A" & row).Value = Range("A" & row - 1).Value Then
    
        counter = counter + 1
        
        
        Else
        
        counter = 1
        
    End If
    Range("D" & row).Value = counter
    'MsgBox counter
        'rand = Int((counter * Rnd + 1))
        
        
        
    
    
    
    
'End If

distrib_vals = 15


'While distrib_vals > 0
Dim random_val As Variant
Dim location As Long
location = 5

Dim totalCounter As Long

        





 If Range("A" & row + 1).Value <> Range("A" & row).Value Then
 
 
       
        selector = Cells(row, 4).Value
        'selector = selector +
        
        Debug.Print "selector: " & elector
        
        
 
        rand = Int((selector * Rnd + 1))
        
        'Range("B5:B" & rand).Value = "x"
        
'        Debug.Print rand
        
        location = location + rand
       ' Debug.Print locayt
        
        
        'Debug.Print "selector: " & selector
        
      '  Debug.Print "row location: " & location
        
        

        'MsgBox rand
'        cells(rand, 3).Value = "x"
'        distrib_vals = distrib_vals - 1
      '  ReDim vArray(1 To selector)
        
        
     '   MsgBox selector
       ' Range("E" & row).Value = selector
        
       ' Debug.Print "row location: " & location
    
        'Range("C" & row).Value = rand
        
        'location = Cells(row + selector, 2).Value
        'Range("G" & row).Value = location
        
        
        
       ' For subdist = LBound(vArray) To UBound(vArray)
            'Debug.Print LBound(vArray)
           ' Debug.Print UBound(vArray)
           
            
           ' rand = int((LBound
            
        '    MsgBox vArray(subdist) 'create individual msgboxs 1,2,3,4,5,6,7,8,9,10,11,12,13-> 1,2,3,etc.
            
        '    rand = WorksheetFunction.RandBetween(LBound(vArray), UBound(vArray))
        '    vArray(subdist) = rand
            
'            random_val = vArray(rand)
        '    MsgBox rand
            
        'Next subdist
    
  
    
    
    
    
    
    
    
    'Debug.Print random_val
    
    
    
        
   ' i = rand
   ' distrib_vals = 15
    
   ' While distrib_vals > 0
        
totalCounter = 0
        
        
End If




'Wend

Next row



        
    



        
  '  Rnd = Int((1 - count * Rnd + 1))
   ' Range(address).Offset(0, 2) = "x"
      
        

'End If
'Wend




    
    'While distrib_vals > 0
    
    
    
    
    
    
    
    

    
    

    'rand = Int((psu_count * Rnd + 1))
     '   If IsEmpty(cells(5 + rand, 3).Value) Then
      '      cells(5 + rand, 3).Value = "x"
       '     distrib_vals = distrib_vals - 1
            
        
       ' End If
    




'Next currentCell

'Next currentCell
    

'Wend


            
            
'While distrib_vals > 0
'    For Each subKey In binDictionary.keys
'        If binDictionary(subKey) = 1 Then
'        rand = Int((2 - 1 + 1) * Rnd + 1)
'            If rand = 1 Then
            
'    rand = Int((psu_count * Rnd) + 1)
'       If IsEmpty(Cells(5 + rand, 3).Value) Then
'    Cells(5 + rand, 3).Value = "x"
'    MsgBox rand
    
    ' ^^ maybe add a way to place an x & have a message box pop that says the name of the PSU?
'    distrib_vals = distrib_vals - 1
'    rand = rand + 1
    
'    End If
    
'Wend
    
                
                
                
            
            
  







' the above section is for randomly assigning PSUs to be sampled. an X is placed by each PSU that has been chosen.
    ' starts at row 6 and goes to the last row populated row in the spreadsheet (i.e., the last PSU)
    ' variable rand is any number between 1 and total amount of PSUs (psu_count)
        ' rand select a new number in that range iteration of the loop
    ' we use rand + 5 to randomly populate a cell in column C starting at row 6
    ' (if IsEmpty) is used so that a new cell is chosen if an x is already present




End Sub

