Sub substitui()

	Dim LR As Integer

	pasta = "C:\Users\usuario\Documents\" ' inserir caminho correto
		
	arquivo = Dir(pasta)

    Do Until arquivo = ""
    
    Set wb = Workbooks.Open(pasta & arquivo)

    Cells.Replace What:="C", Replacement:="", LookAt:=xlPart, SearchOrder:= _
        xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False, _
	FormulaVersion:=xlReplaceFormula2
	LR = Cells(Rows.Count, 1).End(xlUp).Row
	Range("B1").Select
	ActiveCell.FormulaR1C1 = "=RC[-1]*1"
	Range("B1").Copy
	Range("B1:B" & LR).PasteSpecial
	Columns("B:B").Select
	Selection.Copy
	Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
		:=False, Transpose:=False
	LR = Cells(Rows.Count, 2).End(xlUp).Row
	Range("B1:B" & LR).Copy
	Range("A1:A" & LR).PasteSpecial
	Columns("B:B").Clear
	Range("A:A").Select
    Selection.Style = "Comma"
	Columns("A:A").EntireColumn.AutoFit
	
	Application.DisplayAlerts = False
    wb.Close Savechanges:=1
	
arquivo = Dir()

Loop
	
End Sub
