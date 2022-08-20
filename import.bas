Sub importSMDorNAMDENERGYanyPUFA()
Dim ArrayA As Variant
ArrayA = Array("A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "AA", "AB", "AC")
Dim dir As String
Dim filename As String
Dim sheetname As String
Dim num As Long


'dir = "F:/namd/membrane/ciproPUFA/excel/"
'dir = "F:/namd/membrane/tacPUFA/excel/"
dir = "F:/namd/membrane/HBDPUFA/excel/"


Sheets.Add(After:=Sheets(Sheets.Count)).name = "main"

For i = 1 To 3

    filename = dir & "equ" & i & "_namdenergy.csv"
    'filename = dir & "equ" & i & ".csv"
    sheetname = i & "-temp"
    
    Sheets.Add(After:=Sheets(Sheets.Count)).name = sheetname
     Set ws = ActiveWorkbook.Sheets(sheetname)
    With ws.QueryTables.Add(Connection:="TEXT;" & filename, Destination:=ws.Range("A1"))
           .TextFileParseType = xlDelimited
            .TextFileCommaDelimiter = True
            .TextFileSpaceDelimiter = True
            .Refresh
    End With
   
   
    For j = 0 To 7
        num = i + (3 * j)
              
        Sheets(sheetname).Range(ArrayA(j) & ":" & ArrayA(j)).Copy Sheets("main").Range(ArrayA(num) & ":" & ArrayA(num))
        
    Next j
      
Next i


End Sub
