Sub caixa()

'@Lang VBA

	pasta = "C:\Users\usuario\Documents\" ' inserir caminho correto
		
	arquivo = Dir(pasta)
    
    'Fa�a 'at�' arquivo ser vazio
    'Enquanto n�o for vazio ficar� sendo executado
    Do Until arquivo = ""
    
    Set wb = Workbooks.Open(pasta & arquivo)

    Range("A6").Select
    Selection.Copy
    Range("A1").Select
    ActiveSheet.Paste

    Cells.Find(What:="saldo disponivel", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Activate
        
    Row = ActiveCell.Row
    Range("A" & Row).Select
    Selection.Copy
    Range("A2").Select
    ActiveSheet.Paste
        
    Application.DisplayAlerts = False
    wb.Close Savechanges:=1
    
' Esse comando permite encontrar o pr�ximo arquivo
arquivo = Dir()

Loop

End Sub

