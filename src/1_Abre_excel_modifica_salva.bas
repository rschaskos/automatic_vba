Sub Ajuste()

'@Lang VBA

	pasta = "C:\Users\usuario\Documents\" ' inserir caminho correto

	arquivo = Dir(pasta)

	Do Until arquivo = ""

    Set wb = Workbooks.Open(pasta & arquivo)
        
        Dim LR As Long
        LR = Cells(Rows.Count, 16).End(xlUp).Row
        Range("P" & LR).Select
        Selection.Copy
        Range("A1").Select
        ActiveSheet.Paste
	Columns("A").AutoFit
        
    Application.DisplayAlerts = False
    wb.Close Savechanges:=1
    
arquivo = Dir()

Loop

End Sub

