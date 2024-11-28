Sub copy_history()

'inserir data atual
Dim CurrDate As Date
Dim LR As Long
Dim LR2 As Long
Dim LR3 As Long
Dim LR4 As Long
Dim LR5 As Long

'abrir planilha e copiar informações
pasta = "C:\Users\seu_usuario\Documents\1-Contas_Livres\Superávit Financeiro.xlsx"
pasta2 = "C:\Users\seu_usuario\Documents\1-Contas_Livres\HISTORICO-SALDOS.xlsx"

Set wb2 = Workbooks.Open(pasta2)

    Sheets("GERAL").Select
    CurrDate = Date
    LR = Cells(Rows.Count, 1).End(xlUp).Row
    Range("A" & LR + 1) = CurrDate
    


Set wb = Workbooks.Open(pasta)
    Windows("Superávit Financeiro.xlsx").Activate
    Sheets("DADOS").Select
    Range("F2:I2").Select
    Selection.Copy
    
    Windows("HISTORICO-SALDOS.xlsx").Activate
    Sheets("GERAL").Select
    LR2 = Cells(Rows.Count, 2).End(xlUp).Row
    Range("B" & LR2 + 1).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    Windows("Superávit Financeiro.xlsx").Activate
    Sheets("DADOS").Select
    Range("B2:B16,C2:C16,E2:E16").Select
    Selection.Copy
    
    Windows("HISTORICO-SALDOS.xlsx").Activate
    Sheets("ANALITICO").Select
    LR3 = Cells(Rows.Count, 1).End(xlUp).Row
    Range("B" & LR3 + 1).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    LR = Cells(Rows.Count, 1).End(xlUp).Row
    Range("A" & LR + 1) = CurrDate
    Windows("HISTORICO-SALDOS.xlsx").Activate
    LR4 = Cells(Rows.Count, 1).End(xlUp).Row
    LR5 = Cells(Rows.Count, 2).End(xlUp).Row
    Range("A" & LR4).Copy
    Range("A" & LR4 & ":" & "A" & LR5).PasteSpecial
    Range("A" & LR5).Select
        
    Windows("Superávit Financeiro.xlsx").Activate
    Application.DisplayAlerts = False
    wb.Close
    NovoAjuste
    wb2.Close Savechanges:=1

    
End Sub

Sub NovoAjuste()

    Dim LR6 As Long
    Dim LR7 As Long

    Windows("HISTORICO-SALDOS.xlsx").Activate
    Sheets("GERAL").Select
    LR6 = Cells(Rows.Count, 3).End(xlUp).Row
    Cells(LR6, 3).Copy Destination:=Cells(LR6, 6)
    Cells(LR6, 3).FormulaR1C1 = "=RC[3]/2"
    LR7 = Cells(Rows.Count, 5).End(xlUp).Row
    Range("E" & LR7 - 1 & ":E" & LR7 + 1).FillDown

End Sub
