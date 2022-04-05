Attribute VB_Name = "Módulo1"
Sub matrices()
Dim c As Integer
c = 1
suma = 0
suma2 = 0
For i = 1 To 10
 For j = 1 To 10
  Worksheets(1).Cells(i, j).Value = c
  c = c + 1
  If (i = j) Then
     Worksheets(1).Cells(i, j).Interior.Color = RGB(255, 197, 0)
     suma = suma + Worksheets(1).Cells(i, j).Value
     End If
     
  If (i + j = 11) Then
   Worksheets(1).Cells(j, i).Interior.Color = RGB(251, 255, 0)
     suma2 = suma2 + Worksheets(1).Cells(j, i).Value
     End If
     
  Next j
  
  Cells(13, 1).Value = "La suma es: "
  Cells(13, 4).Value = suma2
  
Next i
Cells(12, 1).Value = "La suma es: "
Cells(12, 4).Value = suma

End Sub

