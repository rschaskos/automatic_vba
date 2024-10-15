Sub bb()

'@Lang VBA

	Dim i As Integer

	pasta = "C:\Users\usuario\Documents\" ' inserir caminho correto
		
	arquivo = Dir(pasta)
    
    'Faça 'até' arquivo ser vazio
    'Enquanto não for vazio ficará sendo executado
    Do Until arquivo = ""
    
    Set wb = Workbooks.Open(pasta & arquivo)
    
    Range("A12").Select
    Selection.Copy
    Range("A1").Select
    ActiveSheet.Paste
    
    Range("G1").Select
    ActiveCell.FormulaR1C1 = "=COUNTIF(C[-6],""*Resumo do mês"")"
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    
    If Cells(1, 7) = 1 Then
    
        Cells.Find(What:="resumo", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Activate
        
    Row = ActiveCell.Row
    Range("A" & Row + 9).Select
    Selection.Copy
    Range("A2").Select
    ActiveSheet.Paste
    
    Else
      For i = 2 To 3
            Cells.Find(What:="resumo", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Activate
        
    Row = ActiveCell.Row
    Range("A" & Row + 9).Select
    Selection.Copy
    Range("A" & i).Select
    ActiveSheet.Paste
    Range("A" & Row + 9).Select
    Next i
    

    End If
        
    Application.DisplayAlerts = False
    wb.Close Savechanges:=1
    
' Esse comando permite encontrar o próximo arquivo
arquivo = Dir()

Loop

End Sub
