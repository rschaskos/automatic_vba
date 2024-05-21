Sub empenhos()

    pasta = "C:\Users\usuario\Documents\" ' inserir caminho correto
    
    arquivo = Dir(pasta)

    Set wb = Workbooks.Open(pasta & arquivo)

    Cells.Find(What:="Saldo a pagar", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Activate
        
    With ActiveCell
    C = .Column
    End With
    
        Cells.Find(What:="TOTAL GERAL:", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Activate
        
    With ActiveCell
    R = .Row
    End With
    
    Cells(R, C).Select
    Selection.Copy
    Range("A1").Select
    ActiveSheet.Paste
    Columns("A").AutoFit
    
    Application.DisplayAlerts = False
    wb.Close Savechanges:=1

End Sub
