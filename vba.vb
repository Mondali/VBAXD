Public Sub Task01()
    
    Dim number1 As Integer
    Dim number2 As Integer
    
    number1 = Application.InputBox(prompt:="Insert 1st number", Type:=1)
    number2 = Application.InputBox(prompt:="Insert 2nd number", Type:=1)
    
    If number1 > number2 Then
        MsgBox ("1st number is grater than 2nd")
    ElseIf number2 > number1 Then
        MsgBox ("2nd number is grater than 1st")
    Else
        MsgBox ("Both numbers are equals")
    End If
    
End Sub

Public Sub Task02()
    
    Dim rg As Range
    Dim multiplier As Integer
    
    Set rg = Application.InputBox(prompt:="Select range", Type:=8)
    multiplier = Application.InputBox(prompt:="Insert multiplier", Type:=1)
    
    For Each c In rg.cells
        c.Value = c.Value * multiplier
    Next
    
End Sub
Public Sub Task03()
    
    Dim rg As Range
    Dim multiplier As Integer
    
    Set rg = Application.InputBox(prompt:="Select range", Type:=8)
    multiplier = Application.InputBox(prompt:="Insert multiplier", Type:=1)
    
    For Each c In rg.cells
        If c.Value < 20 Then
            c.Value = c.Value * multiplier
        End If
    Next
    
End Sub

Public Sub Task04()
    
    Dim rg As Range
    Dim multiplier As Integer
    
    Set rg = Application.InputBox(prompt:="Select range", Type:=8)
    multiplier = Application.InputBox(prompt:="Insert multiplier", Type:=1)
    
    For Each c In rg.cells
        If c.Value < 20 And c.Value > 10 Then
            c.Value = c.Value * multiplier
        End If
    Next
    
End Sub

Public Sub Task06()

    Dim i As Integer
    Dim num1 As Integer
    Dim num2 As Integer
    Dim rg As Range
    Dim output As Range
    
    num1 = Range("F7").Value
    num2 = Range("F8").Value
    
    For i = 7 To 15
        Set rg = cells(i, 2)
        Set output = cells(i, 3)

        If rg.Value < num1 Or rg.Value > num2 Then
            output.Value = "Liczba " & rg.Value & " spelnia kryterium"
        End If
    Next i

End Sub

Public Sub Task07()

    Dim i As Integer
    Dim cellB As Range
    Dim cellC As Range
    Dim cellD As Range
    Dim lowValue As Double

    lowValue = Range("F7").Value

    For i = 7 To 15
        Set cellB = cells(i, 2)
        Set cellC = cells(i, 3)
        Set cellD = cells(i, 4)

        If cellB.Value < lowValue And cellC.Value < lowValue Then
            cellD.Value = "Obie mniejsze"
        End If
    Next i

End Sub

Public Sub Task08()

    Dim i As Variant
    Dim rng As Range
    
    Set rng = Range("B5:I5")
    
    For Each i In rng
    
        If i Mod i.Offset(1, 0) = 0 Then
            i.Offset(2, 0).Value = "Podzielna"
        Else: i.Offset(2, 0).Value = "Nie podzielna"
        End If
    Next i

End Sub

Public Sub Task09()

    Dim i As Integer
    Dim cell As Range
    Dim number As Double
    Dim inputNumber As Double

    inputNumber = InputBox("Please enter a number", "Enter Number")

    For i = 2 To 9
        Set cell = cells(5, i)

        number = cell.Value

        If number Mod inputNumber = 0 Then
            cells(6, i).Value = "Podzielna"
        Else
            cells(6, i).Value = "Nie podzielna"
        End If
    Next i

End Sub

Public Sub Task10()

    Dim i As Integer
    Dim sum As Integer
    Dim nNumber As Integer
    Dim cell As Range
    
    nNumber = InputBox("Prosze podac liczbe n", "Podaj liczbę")
    
    sum = 0
    
    For i = 1 To nNumber
        Set cell = cells(5, i)
        sum = sum + cell.Value
    Next i
    
    MsgBox "Suma " & nNumber & " pierwszych liczb wiersza 5 wynosi: " & sum

End Sub
