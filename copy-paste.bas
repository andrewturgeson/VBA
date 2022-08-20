Sub copyfromto()

Dim ArrayA As Variant
ArrayA = Array("P", "L", "B", "Y", "H", "E", "A", "D")


For i = 0 To 7
    Worksheets(ArrayA(i) & "-residues (2)").Range("B2:AD1502").Copy Worksheets(ArrayA(i) & "-residues").Range("B2:AD1502")
    Worksheets(ArrayA(i) & "-atoms (2)").Range("B2:AD1502").Copy Worksheets(ArrayA(i) & "-atoms").Range("B2:AD1502")
Next i

End Sub
