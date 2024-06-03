Option Explicit

Global C As Integer
Global R As Integer

Sub find_value()
    Do While IsEmpty(ActiveCell.Value)
        ActiveCell.Offset(0, 1).Select
    Loop
End Sub

Sub take_value()
    With ActiveCell
        C = .Column
        R = .Row
    End With
End Sub

Sub Main()
    Dim ws As Worksheet
    Dim pasta As String
    Dim arquivo As String
    Dim wb As Workbook
    Dim Rng As Range
    Dim termos() As String
    Dim termo As Variant
    Dim cont1 As Integer
    Dim cont2 As Integer
    Dim cont3 As Integer
    Dim i As Integer

    pasta = "C:\Extratos\Fontes\" 'nome caminho desejado
    arquivo = Dir(pasta)

    Do Until arquivo = ""
        Set wb = Workbooks.Open(pasta & arquivo)
        Set ws = wb.Sheets("Sheet1")
        termos = Split("Saldo fin. inicial ajustado:|Receita orçamentária:|Inscrição de consignação:|Ingresso:|Baixa de consignação:|Baixa realizável por|Egresso:|Despesa orçamentária:|Restos a pagar:|Resultado do ajuste final:|Saldo a pagar:", "|")
        cont1 = 1
        For Each termo In termos
            Cells.Find(What:=termo, After:=ActiveCell, LookIn:=xlFormulas, _
            LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
            MatchCase:=True, SearchFormat:=False).Activate
            take_value
            Set Rng = ws.Cells(R, C)
            If cont1 = 1 Then
                Rng.Cut Destination:=Range("A1")
            ElseIf cont1 = 2 Then
                Rng.Cut Destination:=Range("A2")
            ElseIf cont1 = 3 Then
                Rng.Cut Destination:=Range("A3")
            ElseIf cont1 = 4 Then
                Rng.Cut Destination:=Range("A4")
            ElseIf cont1 = 5 Then
                Rng.Cut Destination:=Range("A5")
            ElseIf cont1 = 6 Then
                Rng.Cut Destination:=Range("A6")
            ElseIf cont1 = 7 Then
                Rng.Cut Destination:=Range("A7")
            ElseIf cont1 = 8 Then
                Rng.Cut Destination:=Range("A8")
            ElseIf cont1 = 9 Then
                Rng.Cut Destination:=Range("A9")
            ElseIf cont1 = 10 Then
                Rng.Cut Destination:=Range("A10")
            ElseIf cont1 = 11 Then
                Rng.Cut Destination:=Range("A11")
            End If
            find_value
            take_value
            Set Rng = ws.Cells(R, C)
            If cont1 = 1 Then
                Rng.Cut Destination:=Range("B1")
            ElseIf cont1 = 2 Then
                Rng.Cut Destination:=Range("B2")
            ElseIf cont1 = 3 Then
                Rng.Cut Destination:=Range("B3")
            ElseIf cont1 = 4 Then
                Rng.Cut Destination:=Range("B4")
            ElseIf cont1 = 5 Then
                Rng.Cut Destination:=Range("B5")
            ElseIf cont1 = 6 Then
                Rng.Cut Destination:=Range("B6")
            ElseIf cont1 = 7 Then
                Rng.Cut Destination:=Range("B7")
            ElseIf cont1 = 8 Then
                Rng.Cut Destination:=Range("B8")
            ElseIf cont1 = 9 Then
                Rng.Cut Destination:=Range("B9")
            ElseIf cont1 = 10 Then
                Rng.Cut Destination:=Range("B10")
            ElseIf cont1 = 11 Then
                Rng.Cut Destination:=Range("B11")
            End If
            cont1 = cont1 + 1
        Next termo
        
        Range("A12").Value = "Restos a pagar Saldo"
            Cells.Find(What:="Inscritos", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        True, SearchFormat:=False).Activate
        take_value
        Cells(R, C).End(xlToLeft).Select
        ActiveCell.Offset(1, 0).Select
        
        ' tratamento especial nessa célula
        cont2 = 1
        For i = 1 To 5
            find_value
            take_value
            Set Rng = ws.Cells(R, C)
            If cont2 = 1 Then
                Rng.Cut Destination:=Range("B12")
            ElseIf cont2 = 2 Then
                Rng.Cut Destination:=Range("C12")
            ElseIf cont2 = 3 Then
                Rng.Cut Destination:=Range("D12")
            ElseIf cont2 = 4 Then
                Rng.Cut Destination:=Range("E12")
            ElseIf cont2 = 5 Then
                Rng.Cut Destination:=Range("F12")
            End If
            cont2 = cont2 + 1
        Next i
        
        ' tratamento especial nessa célula
        cont3 = 1
        Do While IsEmpty(ActiveCell.Value)
            ActiveCell.Offset(0, 1).Select
            If cont3 = 5 Then
                Exit Do
            End If
            cont3 = cont3 + 1
        Loop
        take_value
        Set Rng = ws.Cells(R, C)
        Rng.Cut Destination:=Range("G12")
        
        Range("A1").Select
        Columns("A:A").EntireColumn.AutoFit
        Application.DisplayAlerts = False
        wb.Close Savechanges:=1
        arquivo = Dir
    Loop
End Sub
