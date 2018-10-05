Attribute VB_Name = "Modul1"
Sub binDec()

Dim userInput As String
Dim result As String

userInput = Cells(4, 2).Value
result = myBin2Dec(userInput)
Cells(4, 7).Value = result

End Sub

Public Function myBin2Dec(userInput As String) As Double
 Dim decValue As Double
 For i = 0 To Len(userInput) - 1
  decValue = decValue + Mid(userInput, i + 1, 1) * (2 ^ (Len(userInput) - 1 - i))
 Next
 myBin2Dec = decValue
End Function
