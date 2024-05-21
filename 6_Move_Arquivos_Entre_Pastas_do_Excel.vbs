Sub cp()

'inserir data atual
Dim CurrDate As Date
Dim LR As Long
Dim LR2 As Long
Dim LR3 As Long
Dim LR4 As Long
Dim LR5 As Long

    Sheets("GERAL").Select
    CurrDate = Date
    LR = Cells(Rows.Count, 1).End(xlUp).Row
    Range("A" & LR + 1) = CurrDate
    
	'abrir planilha e copiar informações
	pasta = "C:\Users\usuario\Documents\" ' inserir caminho correto

Set wb = Workbooks.Open(pasta)
    Windows("Superávit Financeiro.xlsx").Activate
    Sheets("DADOS").Select
    Range("F2:I2").Select
    Selection.Copy
    
    Windows("HISTORICO-SALDOS.xlsm").Activate
    Sheets("GERAL").Select
    LR2 = Cells(Rows.Count, 2).End(xlUp).Row
    Range("B" & LR2 + 1).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    Windows("Superávit Financeiro.xlsx").Activate
    Sheets("DADOS").Select
    Range("B2:B16,C2:C16,E2:E16").Select
    Selection.Copy
    
    Windows("HISTORICO-SALDOS.xlsm").Activate
    Sheets("ANALITICO").Select
    LR3 = Cells(Rows.Count, 1).End(xlUp).Row
    Range("B" & LR3 + 1).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    LR = Cells(Rows.Count, 1).End(xlUp).Row
    Range("A" & LR + 1) = CurrDate
    Windows("HISTORICO-SALDOS.xlsm").Activate
    LR4 = Cells(Rows.Count, 1).End(xlUp).Row
    LR5 = Cells(Rows.Count, 2).End(xlUp).Row
    Range("A" & LR4).Copy
    Range("A" & LR4 & ":" & "A" & LR5).PasteSpecial
    Range("A" & LR5).Select
        
    Windows("Superávit Financeiro.xlsx").Activate
    Application.DisplayAlerts = False
    wb.Close Savechanges:=1
    

End Sub
