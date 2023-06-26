VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub randomizer()

    Dim PSU As Double
    Dim totalSubDists As Double
     
    
    
    Dim randomVal As Double
    Dim assignedRandomVal As Double
    Dim sht As Worksheet
    UserInput = InputBox("Enter # of Subdistricts")
    
    'lastrow = ws.Cells(Rows.Count, 1).End(xlUp).row
    '!! cells(rows,columns).value !!
    
        For PSU = 1 To UserInput
            If Cells(PSU + 2, 1).Value <> 0 Then
                
        
        
        Next PSU
        



End Sub
