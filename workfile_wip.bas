Attribute VB_Name = "Module1"
'~MTC 2023
                                                                                                        
Sub randomizer()

Range("B4:H20").Clear
Range("A9:A24").ClearContents

Dim rand As Long
Dim remainder As Long
Dim exactNum As Long
Dim allocatedNum As Double
Dim allocatedNumMod As Long
Dim userInput As Variant

Dim vArray As Variant
Dim yArray As Variant
Dim userInputDists() As Variant



Dim i As Integer




'end user inputs the desired amount of sub-disctrict samples

userInput = InputBox("Enter number of Sub-districts")
    ' Validate the input
    If Not IsNumeric(userInput) Or userInput <= 0 Then
        MsgBox "Invalid input. Please enter a positive numeric value.", vbExclamation
        Exit Sub
    End If
    
Range("B3") = userInput

ReDim yArray(1 To userInput)
ReDim userInputDists(1 To userInput)
    
    
' Prompt the user to enter names of subdistricts

        
    ' Display the input boxes
    For i = 1 To userInput
        
            userInputDists(i) = InputBox("Enter name of Subdistrict " & i & ":", "Sub-district Name")
        
        ' Process the user input (you can modify this part as needed)
        
            If userInputDists(i) = "" Then
                MsgBox "Input process was canceled at Input" & i, vbInformation
                Exit Sub
            Else
                MsgBox "Input Box " & i & " value: " & userInputDists(i), vbInformation
            End If
        
            
    Next i
    
    For i = 1 To userInput
        Cells(i + 8, 1).Value = userInputDists(i)
    
    Next i


'using remainder to allocate left over values into random sub-districts
    remainder = 15 Mod userInput
Range("F4") = remainder

    allocatedNum = 15 / userInput
'round PSU sample values
    allocatedNumMod = Fix(allocatedNum)
Range("G4") = allocatedNumMod

    Set startCell = Range("B9")
    
'rand is any value between 1 and remainder
' ex.) 15/6 = 2 remainder 3..  TF; rand = 1 - 3
    rand = Int((remainder * Rnd) + 1)

Range("H4") = rand

ReDim vArray(1 To userInput)

For subdist = LBound(vArray) To UBound(vArray)
    vArray(subdist) = allocatedNumMod
Next subdist

i = rand

While remainder > 0
    If i > userInput Then i = i - userInput
    vArray(i) = vArray(i) + 1
    remainder = remainder - 1
    i = i + rand
Wend


Range("b9").Resize(UBound(vArray)).Value = Application.Transpose(vArray)




End Sub
